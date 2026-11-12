using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;

namespace MsGraphMcpProvider;

[McpServerToolType]
public sealed class TeamsConversationTools(ITeamsGraphReader graph)
{
    [McpServerTool(Name = "list_chats", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("List the signed-in user's Teams chats (one-to-one, group, and meeting chats), including members. Returns a Graph page; pass @odata.nextLink as nextLink to read more.")]
    public Task<JsonElement> ListChats(
        [Description("Number of chats per page, from 1 to 50.")] int pageSize = 25,
        [Description("@odata.nextLink returned by this tool; omit for the first page.")] string? nextLink = null,
        CancellationToken cancellationToken = default) =>
        ReadPage("me/chats", "Chat.Read", pageSize, nextLink, "&$expand=members", cancellationToken);

    [McpServerTool(Name = "read_chat_messages", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("Read messages in a Teams chat, newest first. Includes sender, timestamps, body (often HTML), and attachment metadata, but does not download attachments. Returns a Graph page with @odata.nextLink.")]
    public Task<JsonElement> ReadChatMessages(
        [Description("Chat id from list_chats.")] string chatId,
        [Description("Number of messages per page, from 1 to 50.")] int pageSize = 25,
        [Description("@odata.nextLink returned for this chat; omit for the first page.")] string? nextLink = null,
        CancellationToken cancellationToken = default) =>
        ReadPage($"chats/{Id(chatId)}/messages", "Chat.Read", pageSize, nextLink,
            "&$orderby=lastModifiedDateTime desc", cancellationToken);

    [McpServerTool(Name = "list_teams", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("List Teams the signed-in user belongs to. Use each team's id with list_channels. Returns a Graph page.")]
    public Task<JsonElement> ListTeams(
        [Description("@odata.nextLink returned by this tool; omit for the first page.")] string? nextLink = null,
        CancellationToken cancellationToken = default) =>
        graph.ReadAsync(PageEndpoint("me/joinedTeams", nextLink), "Team.ReadBasic.All", cancellationToken);

    [McpServerTool(Name = "list_channels", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("List channels in a Team. Access to private/shared channel messages is subject to the signed-in user's permissions. Returns a Graph page.")]
    public Task<JsonElement> ListChannels(
        [Description("Team id from list_teams.")] string teamId,
        [Description("@odata.nextLink returned for this Team; omit for the first page.")] string? nextLink = null,
        CancellationToken cancellationToken = default) =>
        graph.ReadAsync(PageEndpoint($"teams/{Id(teamId)}/channels", nextLink), "Channel.ReadBasic.All", cancellationToken);

    [McpServerTool(Name = "read_channel_messages", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("Read root posts in a Teams channel. Replies are not included; use read_channel_replies for a thread. Returns a Graph page with message bodies and attachment metadata.")]
    public Task<JsonElement> ReadChannelMessages(
        [Description("Team id from list_teams.")] string teamId,
        [Description("Channel id from list_channels.")] string channelId,
        [Description("Number of posts per page, from 1 to 50.")] int pageSize = 25,
        [Description("@odata.nextLink returned for this channel; omit for the first page.")] string? nextLink = null,
        CancellationToken cancellationToken = default) =>
        ReadPage($"teams/{Id(teamId)}/channels/{Id(channelId)}/messages", "ChannelMessage.Read.All",
            pageSize, nextLink, "", cancellationToken);

    [McpServerTool(Name = "read_channel_replies", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("Read replies to a root post in a Teams channel. Returns a Graph page with @odata.nextLink.")]
    public Task<JsonElement> ReadChannelReplies(
        [Description("Team id from list_teams.")] string teamId,
        [Description("Channel id from list_channels.")] string channelId,
        [Description("Root message id from read_channel_messages.")] string messageId,
        [Description("Number of replies per page, from 1 to 50.")] int pageSize = 25,
        [Description("@odata.nextLink returned for this thread; omit for the first page.")] string? nextLink = null,
        CancellationToken cancellationToken = default) =>
        ReadPage($"teams/{Id(teamId)}/channels/{Id(channelId)}/messages/{Id(messageId)}/replies",
            "ChannelMessage.Read.All", pageSize, nextLink, "", cancellationToken);

    internal const long MaxImageBytes = 5 * 1024 * 1024;

    [McpServerTool(Name = "read_chat_image", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("Download an inline image (hosted content) from a Teams chat message. Get messageId and hostedContentId from the <img src=\".../messages/{messageId}/hostedContents/{hostedContentId}/$value\"> in a message body from read_chat_messages. Images over 5 MB are rejected.")]
    public Task<ImageContentBlock> ReadChatImage(
        [Description("Chat id from list_chats.")] string chatId,
        [Description("Message id containing the image.")] string messageId,
        [Description("Hosted content id from the image URL.")] string hostedContentId,
        CancellationToken cancellationToken = default) =>
        ReadImage($"chats/{Id(chatId)}/messages/{Id(messageId)}/hostedContents/{Id(hostedContentId)}/$value",
            "Chat.Read", cancellationToken);

    [McpServerTool(Name = "read_channel_image", ReadOnly = true, Destructive = false, OpenWorld = true)]
    [Description("Download an inline image (hosted content) from a Teams channel message or reply. Use the ids from the <img src> URL in the message body. For a reply, pass the reply id as messageId and the root post id as rootMessageId. Images over 5 MB are rejected.")]
    public Task<ImageContentBlock> ReadChannelImage(
        [Description("Team id from list_teams.")] string teamId,
        [Description("Channel id from list_channels.")] string channelId,
        [Description("Message id containing the image (the reply id for a reply).")] string messageId,
        [Description("Hosted content id from the image URL.")] string hostedContentId,
        [Description("Root post id when messageId is a reply; omit otherwise.")] string? rootMessageId = null,
        CancellationToken cancellationToken = default)
    {
        var message = rootMessageId == null
            ? $"messages/{Id(messageId)}"
            : $"messages/{Id(rootMessageId)}/replies/{Id(messageId)}";
        return ReadImage($"teams/{Id(teamId)}/channels/{Id(channelId)}/{message}/hostedContents/{Id(hostedContentId)}/$value",
            "ChannelMessage.Read.All", cancellationToken);
    }

    private async Task<ImageContentBlock> ReadImage(string endpoint, string scope, CancellationToken cancellationToken)
    {
        var response = await graph.ReadBytesAsync(endpoint, scope, MaxImageBytes, cancellationToken);
        if (!response.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            throw new McpException($"The content is not an image (type: {response.ContentType}).");
        return ImageContentBlock.FromBytes(response.Content, response.ContentType);
    }

    private Task<JsonElement> ReadPage(string path, string scope, int pageSize, string? nextLink,
        string query, CancellationToken cancellationToken)
    {
        if (pageSize < 1 || pageSize > 50)
            throw new McpException("pageSize must be between 1 and 50.");

        return graph.ReadAsync(nextLink == null ? $"{path}?$top={pageSize}{query}" : PageEndpoint(path, nextLink),
            scope, cancellationToken);
    }

    internal static string Id(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "." || value == "..")
            throw new McpException("A non-empty Teams identifier is required.");
        return Uri.EscapeDataString(value);
    }

    private static string PageEndpoint(string path, string? nextLink)
    {
        if (nextLink == null)
            return path;

        var expected = new Uri($"https://graph.microsoft.com/v1.0/{path}");
        // Never forward a Microsoft access token to a caller-supplied host or another resource.
        if (!Uri.TryCreate(nextLink, UriKind.Absolute, out var next) ||
            next.Scheme != "https" || next.Host != expected.Host || !next.IsDefaultPort ||
            next.UserInfo.Length != 0 || next.Fragment.Length != 0 ||
            !SamePath(next, expected))
            throw new McpException("nextLink must be a Microsoft Graph continuation URL for the same conversation/list.");

        return next.AbsoluteUri;
    }

    private static bool SamePath(Uri actual, Uri expected)
    {
        // Graph may return reserved characters in IDs (e.g. ':' and '@') unescaped.
        // Compare decoded segments, not the entire decoded path, to preserve path boundaries.
        var actualSegments = actual.AbsolutePath.Split('/');
        var expectedSegments = expected.AbsolutePath.Split('/');
        return actualSegments.Length == expectedSegments.Length &&
            actualSegments.Zip(expectedSegments).All(pair =>
                string.Equals(Uri.UnescapeDataString(pair.First), Uri.UnescapeDataString(pair.Second), StringComparison.Ordinal));
    }
}
