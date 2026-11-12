using System.Text.Json;

namespace WinUIExample;

/// <summary>Per-user sample settings stored under LocalApplicationData, outside the repository.</summary>
public sealed class UserSettings
{
	public string ClientId { get; set; } = "";
	public string TenantId { get; set; } = "";

	private static string FilePath => Path.Combine(
		Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
		"MsGraphMcpProvider", "WinUIExample", "settings.json");

	public static UserSettings Load()
	{
		try
		{
			return JsonSerializer.Deserialize<UserSettings>(File.ReadAllText(FilePath)) ?? new UserSettings();
		}
		catch (Exception)
		{
			return new UserSettings();
		}
	}

	public void Save()
	{
		try
		{
			Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
			File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
		}
		catch (Exception)
		{
			// Settings are a convenience; ignore storage failures.
		}
	}
}
