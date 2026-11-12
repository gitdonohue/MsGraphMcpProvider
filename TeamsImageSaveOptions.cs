namespace MsGraphMcpProvider;

/// <summary>
/// Opt-in configuration that lets MCP clients save Teams inline images to disk.
/// Files are only ever written inside <see cref="Directory"/>; callers cannot choose paths.
/// </summary>
public sealed class TeamsImageSaveOptions
{
	/// <summary>Folder that saved images are written to. Created on first save.</summary>
	public required string Directory { get; init; }
}
