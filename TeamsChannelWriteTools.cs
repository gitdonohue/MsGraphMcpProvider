using ModelContextProtocol;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text;
using System.Text.Json;

namespace MsGraphMcpProvider;

[McpServerToolType]
public sealed class TeamsChannelWriteTools(ITeamsGraphWriter graph, TeamsChannelWriteAccess access)
{
	[McpServerTool(Name = "send_channel_message", ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = true)]
	[Description("Post a new plain-text message to a Teams channel as the signed-in user. Requires the application's channel-write opt-in and Microsoft ChannelMessage.Send consent. This publishes a visible message; confirm the destination and content with the user first. Do not automatically retry: a repeated call can publish a duplicate.")]
	public Task<JsonElement> SendChannelMessage(
		[Description("Team id from list_teams.")] string teamId,
		[Description("Channel id from list_channels.")] string channelId,
		[Description("Plain-text message to publish, up to 24,000 UTF-8 bytes. HTML and mentions are not interpreted.")] string content,
		CancellationToken cancellationToken = default)
	{
		// Also enforce opt-in at invocation, including calls from clients with cached tool definitions.
		if (!access.Enabled)
			throw new McpException("Channel posting is disabled. Enable 'Allow MCP to post channel messages' in the application's Settings first.");

		var endpoint = $"teams/{TeamsConversationTools.Id(teamId)}/channels/{TeamsConversationTools.Id(channelId)}/messages";
		if (string.IsNullOrWhiteSpace(content))
			throw new McpException("Message content must not be empty.");
		if (Encoding.UTF8.GetByteCount(content) > 24000)
			throw new McpException("Message content must not exceed 24,000 UTF-8 bytes.");

		var body = JsonSerializer.SerializeToElement(new { body = new { contentType = "text", content } });
		return graph.PostAsync(endpoint, body, "ChannelMessage.Send", cancellationToken);
	}
}
