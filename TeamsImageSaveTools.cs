using ModelContextProtocol;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;

namespace MsGraphMcpProvider;

[McpServerToolType]
public sealed class TeamsImageSaveTools(ITeamsGraphReader graph, TeamsImageSaveOptions options)
{
	[McpServerTool(Name = "save_chat_image", ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = true)]
	[Description("Download an inline image (hosted content) from a Teams chat message and save it to the application's configured image folder. Returns the saved file path. Get the ids from the <img src=\".../messages/{messageId}/hostedContents/{hostedContentId}/$value\"> in a message body. Images over 5 MB are rejected.")]
	public Task<JsonElement> SaveChatImage(
		[Description("Chat id from list_chats.")] string chatId,
		[Description("Message id containing the image.")] string messageId,
		[Description("Hosted content id from the image URL.")] string hostedContentId,
		CancellationToken cancellationToken = default) =>
		SaveImage($"chats/{TeamsConversationTools.Id(chatId)}/messages/{TeamsConversationTools.Id(messageId)}/hostedContents/{TeamsConversationTools.Id(hostedContentId)}/$value",
			"Chat.Read", cancellationToken);

	[McpServerTool(Name = "save_channel_image", ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = true)]
	[Description("Download an inline image (hosted content) from a Teams channel message or reply and save it to the application's configured image folder. Returns the saved file path. For a reply, pass the reply id as messageId and the root post id as rootMessageId. Images over 5 MB are rejected.")]
	public Task<JsonElement> SaveChannelImage(
		[Description("Team id from list_teams.")] string teamId,
		[Description("Channel id from list_channels.")] string channelId,
		[Description("Message id containing the image (the reply id for a reply).")] string messageId,
		[Description("Hosted content id from the image URL.")] string hostedContentId,
		[Description("Root post id when messageId is a reply; omit otherwise.")] string? rootMessageId = null,
		CancellationToken cancellationToken = default)
	{
		var message = rootMessageId == null
			? $"messages/{TeamsConversationTools.Id(messageId)}"
			: $"messages/{TeamsConversationTools.Id(rootMessageId)}/replies/{TeamsConversationTools.Id(messageId)}";
		return SaveImage($"teams/{TeamsConversationTools.Id(teamId)}/channels/{TeamsConversationTools.Id(channelId)}/{message}/hostedContents/{TeamsConversationTools.Id(hostedContentId)}/$value",
			"ChannelMessage.Read.All", cancellationToken);
	}

	private async Task<JsonElement> SaveImage(string endpoint, string scope, CancellationToken cancellationToken)
	{
		var response = await graph.ReadBytesAsync(endpoint, scope, TeamsConversationTools.MaxImageBytes, cancellationToken);
		if (!response.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
			throw new McpException($"The content is not an image (type: {response.ContentType}).");

		// The file name is generated here; nothing from the caller or the message is used in the path.
		var fileName = $"teams-image-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}{Extension(response.ContentType)}";
		string path;
		try
		{
			var directory = Path.GetFullPath(options.Directory);
			Directory.CreateDirectory(directory);
			path = Path.Combine(directory, fileName);
			await using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
			await stream.WriteAsync(response.Content, cancellationToken);
		}
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
		{
			throw new McpException("The image could not be saved to the configured folder.");
		}

		return JsonSerializer.SerializeToElement(new { path, bytes = response.Content.Length, contentType = response.ContentType });
	}

	private static string Extension(string contentType) => contentType.ToLowerInvariant() switch
	{
		"image/png" => ".png",
		"image/jpeg" or "image/jpg" => ".jpg",
		"image/gif" => ".gif",
		"image/bmp" => ".bmp",
		"image/webp" => ".webp",
		_ => ".bin"
	};
}
