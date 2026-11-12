using ModelContextProtocol;
using System.Net;
using System.Text.Json;

namespace MsGraphMcpProvider;

public sealed class TeamsGraphReader(Func<ITeamsGraphClient?> getClient) : ITeamsGraphReader, ITeamsGraphWriter
{
	public Task<JsonElement> ReadAsync(string endpoint, string scope, CancellationToken cancellationToken) =>
		ExecuteAsync(client => client.GetAsync(endpoint, scope, cancellationToken), scope, false, cancellationToken);

	public async Task<GraphBinaryResponse> ReadBytesAsync(string endpoint, string scope, long maxBytes, CancellationToken cancellationToken)
	{
		var client = getClient() ?? throw new McpException(NotConfigured);
		GraphBinaryResponse response;
		try
		{
			response = await client.GetBytesAsync(endpoint, scope, maxBytes, cancellationToken);
		}
		catch (HttpRequestException)
		{
			throw new McpException("Could not reach Microsoft Graph. Check your internet connection and retry.");
		}
		catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
		{
			throw new McpException("The Microsoft Graph request timed out. Retry later.");
		}
		catch (InvalidDataException ex)
		{
			throw new McpException(ex.Message);
		}
		CheckStatus(response.Status, scope, false);
		return response;
	}

	private const string NotConfigured = "Microsoft Graph is not configured. Set the Client Id and Tenant Id in the application's settings first.";

	public Task<JsonElement> PostAsync(string endpoint, JsonElement body, string scope, CancellationToken cancellationToken) =>
		ExecuteAsync(client => client.PostAsync(endpoint, body, scope, cancellationToken), scope, true, cancellationToken);

	private async Task<JsonElement> ExecuteAsync(Func<ITeamsGraphClient, Task<GraphResponse>> request,
		string scope, bool isWrite, CancellationToken cancellationToken)
	{
		var client = getClient();
		if (client == null)
			throw new McpException(NotConfigured);

		GraphResponse response;
		try
		{
			response = await request(client);
		}
		catch (HttpRequestException)
		{
			throw new McpException(isWrite
				? "The Microsoft Graph connection failed. Delivery is uncertain; check the channel before retrying to avoid duplicate messages."
				: "Could not reach Microsoft Graph. Check your internet connection and retry.");
		}
		catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
		{
			throw new McpException(isWrite
				? "The Microsoft Graph request timed out. Delivery is uncertain; check the channel before retrying to avoid duplicate messages."
				: "The Microsoft Graph request timed out. Retry later.");
		}
		CheckStatus(response.Status, scope, isWrite);

		using var document = JsonDocument.Parse(response.Content);
		return document.RootElement.Clone();
	}

	private static void CheckStatus(HttpStatusCode status, string scope, bool isWrite)
	{
		switch (status)
		{
			case HttpStatusCode.Unauthorized:
				throw new McpException("Microsoft sign-in/consent is required. Sign in in the application, then retry. Tenant administrator approval may still be required.");
			case HttpStatusCode.Forbidden:
				throw new McpException($"Microsoft Graph denied access. Ensure the app registration has delegated {scope} permission, tenant consent, and that you can access this conversation.");
			case HttpStatusCode.NotFound:
				throw new McpException("The Teams conversation was not found or is not accessible to the signed-in user.");
			case HttpStatusCode.TooManyRequests:
				throw new McpException("Microsoft Graph is throttling requests. Wait before retrying.");
		}
		if ((int)status < 200 || (int)status >= 300)
			throw new McpException(isWrite
				? $"Microsoft Graph returned HTTP {(int)status}. Check the channel before retrying to avoid duplicate messages."
				: $"Microsoft Graph returned HTTP {(int)status}. Retry later.");
	}
}
