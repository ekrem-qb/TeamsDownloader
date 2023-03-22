using System.Diagnostics;
using System.Runtime.InteropServices;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Nini.Config;
using TextCopy;
using Process = System.Diagnostics.Process;

namespace TeamsDownloader;

internal abstract class Program
{
	private const string TenantId = "common";
	private const string ClientId = "24fefc72-2336-4c45-9518-5d61c3a6306e";
	private const string DriveRootPath = "root";
	private const string FilePathPrefix = "root:/";
	private const string SettingsFileName = "Settings.ini";
	private const string ConfigName = "Main";
	private const string SaveFolderKey = "SaveFolder";

	private static readonly string[] Scopes = { "Team.ReadBasic.All", "Files.Read.All" };
	private static readonly string SettingsFilePath = Path.GetFullPath(SettingsFileName);

	private static IniConfigSource? _settings;
	private static GraphServiceClient? _graphClient;
	private static string? _saveFolderPath;

	private static async Task Main()
	{
		Setup();

		try
		{
			await FetchMeetingRecordings();
		}
		catch (ODataError odataError)
		{
			Console.WriteLine(odataError.Error?.Code);
			Console.WriteLine(odataError.Error?.Message);
		}
	}

	private static void Setup()
	{
		if (!File.Exists(SettingsFilePath))
		{
			FileStream file = File.Create(SettingsFilePath);
			file.Close();
		}

		_settings = new(SettingsFilePath)
		{
			AutoSave = true,
		};

		if (_settings.Configs[ConfigName] == null)
		{
			_settings.AddConfig(ConfigName);
		}

		_saveFolderPath = _settings.Configs[ConfigName].Get(SaveFolderKey);
		if (string.IsNullOrEmpty(_saveFolderPath))
		{
			_saveFolderPath = KnownFolders.GetPath(KnownFolder.Downloads);
			_settings.Configs[ConfigName].Set(SaveFolderKey, _saveFolderPath);
		}

		TokenCredentialOptions options = new() { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };
		DeviceCodeCredential deviceCodeCredential = new(Authenticate, TenantId, ClientId, options);
		_graphClient = new(deviceCodeCredential, Scopes);
	}

	private static void OpenUrl(string url)
	{
		try
		{
			Process.Start(url);
		}
		catch
		{
			// hack because of this: https://github.com/dotnet/corefx/issues/10361
			if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
			{
				url = url.Replace("&", "^&");
				Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
			}
			else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
			{
				Process.Start("xdg-open", url);
			}
			else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
			{
				Process.Start("open", url);
			}
			else
			{
				throw;
			}
		}
	}

	private static Task Authenticate(DeviceCodeInfo code, CancellationToken cancellation)
	{
		OpenUrl(code.VerificationUri.AbsoluteUri);
		ClipboardService.SetText(code.UserCode);
		Console.WriteLine(code.Message);
		return Task.FromResult(0);
	}

	private static async Task FetchMeetingRecordings()
	{
		if (_graphClient == null) return;

		TeamCollectionResponse? teams = await _graphClient.Me.JoinedTeams.GetAsync();
		if (teams?.Value == null) return;

		foreach (Team team in teams.Value)
		{
			try
			{
				await FetchDrive(team);
			}
			catch (ODataError odataError)
			{
				Console.WriteLine(odataError.Error?.Code);
				Console.WriteLine(odataError.Error?.Message);
			}
		}
	}

	private static async Task FetchDrive(Team team)
	{
		if (_graphClient == null) return;

		Drive? drive = await _graphClient.Groups[team.Id].Drive.GetAsync();
		if (drive?.Id == null) return;

		await SearchForVideoFiles(DriveRootPath, drive.Id, team.DisplayName ?? "");
	}

	private static async Task SearchForVideoFiles(string itemId, string driveId, string teamName)
	{
		if (_graphClient == null) return;

		DriveItemCollectionResponse? children = await _graphClient.Drives[driveId].Items[itemId].Children.GetAsync();
		if (children?.Value == null) return;

		foreach (DriveItem child in children.Value)
		{
			if (child.Video != null)
			{
				await DownloadRecording(child, driveId, teamName);
			}
			else if (child.Folder != null)
			{
				if (child.Id == null) continue;

				await SearchForVideoFiles(child.Id, driveId, teamName);
			}
		}
	}

	private static async Task DownloadRecording(DriveItem file, string driveId, string teamName)
	{
		if (file.ParentReference?.Path == null) return;
		if (file.Name == null) return;

		int indexOfPathWithoutPrefix = file.ParentReference.Path.IndexOf(FilePathPrefix, StringComparison.Ordinal);
		if (indexOfPathWithoutPrefix < 0) return;

		string filePathInTeams = file.ParentReference.Path[(indexOfPathWithoutPrefix + FilePathPrefix.Length)..];

		if (_saveFolderPath == null) return;

		string filePath = Path.GetFullPath(Path.Combine(_saveFolderPath, teamName, filePathInTeams, file.Name));
		if (File.Exists(filePath))
		{
			Console.WriteLine("File already exists: " + file.Name);
			return;
		}

		string? directoryName = Path.GetDirectoryName(filePath);
		if (directoryName == null) return;

		DirectoryInfo directoryInfo = Directory.CreateDirectory(directoryName);
		if (!directoryInfo.Exists)
		{
			Console.WriteLine($"Can't create directory: {directoryName}");
			return;
		}

		try
		{
			if (_graphClient == null) return;

			Stream? cloudFileStream = await _graphClient.Drives[driveId].Items[file.Id].Content.GetAsync();
			if (cloudFileStream == null) return;

			FileStream localFileStream = File.Create(filePath);

			await cloudFileStream.CopyToAsync(localFileStream);

			localFileStream.Close();
			Console.WriteLine($"Successfully downloaded: {file.Name}");
		}
		catch (ODataError odataError)
		{
			Console.WriteLine($"Error downloading: {file.WebUrl}");
			Console.WriteLine(odataError.Error?.Code);
			Console.WriteLine(odataError.Error?.Message);
		}
		catch (Exception e)
		{
			Console.WriteLine($"Error downloading: {file.WebUrl}");
			Console.WriteLine(e);
		}
	}
}