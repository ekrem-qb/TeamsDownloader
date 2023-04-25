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
	private static readonly List<Func<Task>> FileTasks = new();

	private static async Task Main()
	{
		Setup();

		try
		{
			await FetchMeetingRecordings();

			for (int i = 0; i < FileTasks.Count; i++)
			{
				await FileTasks[i].Invoke();
			}
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

		Task[] tasks = new Task[teams.Value.Count];

		for (int i = 0; i < teams.Value.Count; i++)
		{
			Team team = teams.Value[i];
			try
			{
				tasks[i] = FetchDrive(team);
			}
			catch (ODataError odataError)
			{
				Console.WriteLine(odataError.Error?.Code);
				Console.WriteLine(odataError.Error?.Message);
			}
		}

		await Task.WhenAll(tasks);
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

		List<Task> tasks = new();

		for (int i = 0; i < children.Value.Count; i++)
		{
			DriveItem child = children.Value[i];
			if (child.Video != null)
			{
				CheckFileExistence(child, driveId, teamName);
			}
			else if (child.Folder != null)
			{
				if (child.Id == null) continue;

				tasks.Add(SearchForVideoFiles(child.Id, driveId, teamName));
			}
		}

		await Task.WhenAll(tasks);
	}

	private static void CheckFileExistence(DriveItem file, string driveId, string teamName)
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

		FileTasks.Add(() => DownloadFile(file, driveId, filePath));
	}

	private static async Task DownloadFile(DriveItem file, string driveId, string filePath)
	{
		try
		{
			if (_graphClient == null) return;

			Stream? cloudFileStream = await _graphClient.Drives[driveId].Items[file.Id].Content.GetAsync();
			if (cloudFileStream == null) return;

			FileStream localFileStream = File.Create(filePath);

			await cloudFileStream.CopyToAsync(localFileStream);

			localFileStream.Close();
			Console.WriteLine($"Successfully downloaded: {file.Name}");

			if (file.CreatedDateTime.HasValue)
			{
				File.SetCreationTimeUtc(filePath, file.CreatedDateTime.Value.DateTime);
			}

			if (file.LastModifiedDateTime.HasValue)
			{
				File.SetLastWriteTimeUtc(filePath, file.LastModifiedDateTime.Value.DateTime);
			}
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