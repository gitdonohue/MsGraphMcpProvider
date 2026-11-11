# MsGraphMcpProvider

Embeddable [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that gives AI clients access to the signed-in user's Microsoft Teams chats and channels through Microsoft Graph. Host it inside your own .NET app and point any MCP client at it.

- Loopback-only HTTP listener (`127.0.0.1`), stateless MCP transport at `/mcp`.
- Rejects non-local `Host` headers and any request with an `Origin` header (DNS-rebinding and browser protection).
- Read-only by default. Channel posting is an explicit opt-in.
- Sign-in with MSAL: **ClientId + TenantId** (Windows broker) or **ClientId + device code**.
- Or bring your own Graph transport by implementing `ITeamsGraphClient`.

## Install

```
dotnet add package MsGraphMcpProvider
```

Requires .NET 10 (uses the ASP.NET Core shared framework). The broker sign-in mode is Windows-only; device code works anywhere.

## Quick start

```csharp
using MsGraphMcpProvider;

// Option 1: device code (ClientId only; tenant optional, defaults to "organizations")
var client = new MsalGraphClient(clientId, tenantId, useDeviceCode: true);

// Option 2: ClientId + TenantId with the Windows broker (interactive sign-in on first use)
// var client = new MsalGraphClient(clientId, tenantId, hwnd: windowHandle);

await using var provider = new MsGraphMcpProvider.MsGraphMcpProvider();
await provider.StartAsync(5123, client);   // MCP endpoint: http://127.0.0.1:5123/mcp

// Device code mode: sign-in is never triggered by an MCP caller, only by your app.
await client.SignInWithDeviceCodeAsync(code => Console.WriteLine(code.Message));
```

Configure your MCP client with the URL `http://127.0.0.1:5123/mcp` (Streamable HTTP transport).

### Enabling channel posting

```csharp
var writeAccess = new TeamsChannelWriteAccess { Enabled = true };
await provider.StartAsync(5123, client, writeAccess: writeAccess);
```

This registers the `send_channel_message` tool. Setting `writeAccess.Enabled = false` later makes the tool refuse calls without restarting the server. The signed-in user must have consented to `ChannelMessage.Send`.

## Tools

| Tool | Description | Graph scope |
| --- | --- | --- |
| `list_chats` | List the user's chats, with members | `Chat.Read` |
| `read_chat_messages` | Read messages in a chat | `Chat.Read` |
| `list_teams` | List joined teams | `Team.ReadBasic.All` |
| `list_channels` | List channels in a team | `Channel.ReadBasic.All` |
| `read_channel_messages` | Read root posts in a channel | `ChannelMessage.Read.All` |
| `read_channel_replies` | Read replies to a post | `ChannelMessage.Read.All` |
| `send_channel_message` | Post a plain-text message (opt-in) | `ChannelMessage.Send` |

List and read tools return Graph pages; pass `@odata.nextLink` back as `nextLink` to continue. Only Graph continuation URLs for the same list are accepted.

## Entra app registration

1. Register an app in Microsoft Entra ID and note the **Application (client) ID** (and tenant ID).
2. Under **Authentication**, add a mobile/desktop platform and enable **Allow public client flows** (required for device code).
3. Add the delegated Graph permissions from the table above. Some require tenant administrator consent.

## Bring your own Graph client

Implement `ITeamsGraphClient` to supply tokens and HTTP yourself:

```csharp
public interface ITeamsGraphClient
{
	Task<GraphResponse> GetAsync(string endpoint, string scope, CancellationToken ct);
	Task<GraphResponse> PostAsync(string endpoint, JsonElement body, string scope, CancellationToken ct);
}
```

`endpoint` is relative to `https://graph.microsoft.com/v1.0/` (or an absolute v1.0 continuation URL). Return a `GraphResponse(content, statusCode)`; Graph errors are translated into MCP errors for you. `StartAsync` also accepts a `Func<ITeamsGraphClient?>` if the client is created later.

## Security notes

- Message bodies and attachments are untrusted content; the server tells clients not to treat them as instructions.
- Logging is disabled so message content and credentials are not written to logs.
- Tokens are cached in memory only; device-code sign-in must be repeated after a restart.

## Sample

`samples/WinUIExample` is a WinUI 3 app that starts the server, signs in with device code (code shown in bold with a copy button) and optionally enables posting.

## License

MIT
