using System.Net;
using System.Text.Json;

namespace MsGraphMcpProvider;

public record struct GraphResponse(string Content, HttpStatusCode Status);

public record struct GraphBinaryResponse(byte[] Content, string ContentType, HttpStatusCode Status);

/// <summary>
/// Host-supplied Microsoft Graph transport. The host owns authentication and token acquisition.
/// Endpoints are relative to https://graph.microsoft.com/v1.0/ or absolute Graph v1.0 continuation URLs.
/// </summary>
public interface ITeamsGraphClient
{
	Task<GraphResponse> GetAsync(string endpoint, string scope, CancellationToken cancellationToken);

	Task<GraphResponse> PostAsync(string endpoint, JsonElement body, string scope, CancellationToken cancellationToken);

	/// <summary>Downloads binary content (for example hosted images). Implementations should enforce <paramref name="maxBytes"/>.</summary>
	Task<GraphBinaryResponse> GetBytesAsync(string endpoint, string scope, long maxBytes, CancellationToken cancellationToken) =>
		throw new NotSupportedException("This Graph client does not support binary downloads.");
}

public interface ITeamsGraphReader
{
	Task<JsonElement> ReadAsync(string endpoint, string scope, CancellationToken cancellationToken);

	Task<GraphBinaryResponse> ReadBytesAsync(string endpoint, string scope, long maxBytes, CancellationToken cancellationToken) =>
		throw new NotSupportedException("This Graph reader does not support binary downloads.");
}

public interface ITeamsGraphWriter
{
	Task<JsonElement> PostAsync(string endpoint, JsonElement body, string scope, CancellationToken cancellationToken);
}

public sealed class TeamsChannelWriteAccess
{
	private bool enabled;

	public bool Enabled
	{
		get => Volatile.Read(ref enabled);
		set => Volatile.Write(ref enabled, value);
	}
}
