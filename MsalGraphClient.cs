using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace MsGraphMcpProvider;

/// <summary>
/// MSAL-backed <see cref="ITeamsGraphClient"/>. Supports either ClientId + TenantId with the Windows broker
/// (interactive sign-in), or ClientId + device code (sign-in is started explicitly via
/// <see cref="SignInWithDeviceCodeAsync"/>, never by an MCP caller).
/// </summary>
public sealed class MsalGraphClient : ITeamsGraphClient
{
	public static readonly string[] DefaultScopes =
	[
		"User.Read", "Chat.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All", "Files.Read.All"
	];

	private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(30) };
	private static readonly Uri GraphBase = new("https://graph.microsoft.com/v1.0/");

	private readonly SemaphoreSlim authLock = new(1, 1);
	private readonly IntPtr hwnd;
	private AuthenticationResult? authCache;

	public bool UseDeviceCode { get; }
	public IPublicClientApplication PublicClientApp { get; }

	/// <param name="clientId">Entra app registration (client) id.</param>
	/// <param name="tenantId">Tenant id or domain. Defaults to "organizations" when empty.</param>
	/// <param name="useDeviceCode">Use device-code sign-in instead of the Windows broker.</param>
	/// <param name="hwnd">Parent window for interactive broker prompts (broker mode only).</param>
	/// <param name="persistTokenCache">Persist the device-code token cache (encrypted by the OS where supported) so sign-in survives restarts.</param>
	/// <param name="cacheFilePath">Cache file location. Defaults to a per-client file under LocalApplicationData.</param>
	public MsalGraphClient(string clientId, string? tenantId = null, bool useDeviceCode = false, IntPtr hwnd = default,
		bool persistTokenCache = true, string? cacheFilePath = null)
	{
		ArgumentException.ThrowIfNullOrWhiteSpace(clientId);
		UseDeviceCode = useDeviceCode;
		this.hwnd = hwnd;
		if (this.hwnd == IntPtr.Zero && !useDeviceCode && OperatingSystem.IsWindows())
			this.hwnd = GetConsoleWindow();

		var tenant = string.IsNullOrWhiteSpace(tenantId) ? "organizations" : tenantId;
		var builder = PublicClientApplicationBuilder.Create(clientId)
			.WithAuthority($"https://login.microsoftonline.com/{tenant}")
			.WithDefaultRedirectUri();
		// Device-code clients must not use the Windows broker.
		PublicClientApp = useDeviceCode
			? builder.Build()
			: builder.WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows)).Build();

		// The broker manages its own cache; only device-code mode needs explicit persistence.
		if (useDeviceCode && persistTokenCache)
			RegisterPersistentCache(clientId, cacheFilePath);
	}

	private void RegisterPersistentCache(string clientId, string? cacheFilePath)
	{
		try
		{
			var path = cacheFilePath ?? Path.Combine(
				Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
				"MsGraphMcpProvider", $"msal_{clientId}.cache");
			var directory = Path.GetDirectoryName(Path.GetFullPath(path))!;
			Directory.CreateDirectory(directory);
			var properties = new StorageCreationPropertiesBuilder(Path.GetFileName(path), directory)
				.WithLinuxKeyring("com.gitdonohue.msgraphmcpprovider", MsalCacheHelper.LinuxKeyRingDefaultCollection,
					"MsGraphMcpProvider token cache",
					new KeyValuePair<string, string>("Version", "1"), new KeyValuePair<string, string>("ProductGroup", "MsGraphMcpProvider"))
				.WithMacKeyChain("MsGraphMcpProvider", clientId)
				.Build();
			var helper = MsalCacheHelper.CreateAsync(properties).GetAwaiter().GetResult();
			helper.RegisterCache(PublicClientApp.UserTokenCache);
		}
		catch (Exception)
		{
			// Secure storage unavailable: fall back to an in-memory cache.
		}
	}

	[DllImport("kernel32.dll")]
	private static extern IntPtr GetConsoleWindow();

	/// <summary>Starts device-code sign-in. Display the code only in trusted application UI.</summary>
	public async Task<AuthenticationResult> SignInWithDeviceCodeAsync(Action<DeviceCodeResult> showCode,
		string[]? scopes = null, CancellationToken cancellationToken = default)
	{
		if (!UseDeviceCode)
			throw new InvalidOperationException("This client is not configured for device-code sign-in.");
		await authLock.WaitAsync(cancellationToken);
		try
		{
			var result = await PublicClientApp.AcquireTokenWithDeviceCode(scopes ?? DefaultScopes, code =>
			{
				showCode(code);
				return Task.CompletedTask;
			}).ExecuteAsync(cancellationToken);
			// Single-account cache: don't silently choose another user later.
			foreach (var account in await PublicClientApp.GetAccountsAsync())
				if (account.HomeAccountId.Identifier != result.Account.HomeAccountId.Identifier)
					await PublicClientApp.RemoveAsync(account);
			authCache = result;
			return result;
		}
		finally
		{
			authLock.Release();
		}
	}

	/// <summary>Tries to sign in from the persisted token cache without any user interaction. Returns null if not possible.</summary>
	public async Task<AuthenticationResult?> TrySignInSilentAsync(string[]? scopes = null, CancellationToken cancellationToken = default)
	{
		if (!UseDeviceCode)
			return null;
		await authLock.WaitAsync(cancellationToken);
		try
		{
			var account = (await PublicClientApp.GetAccountsAsync()).FirstOrDefault();
			if (account == null)
				return null;
			authCache = await PublicClientApp.AcquireTokenSilent(scopes ?? DefaultScopes, account).ExecuteAsync(cancellationToken);
			return authCache;
		}
		catch (MsalException)
		{
			return null;
		}
		finally
		{
			authLock.Release();
		}
	}

	/// <summary>Forces a token refresh, bypassing the in-memory cache. Returns null if no account is available or refresh fails.</summary>
	public async Task<AuthenticationResult?> RefreshTokenAsync(string[]? scopes = null, CancellationToken cancellationToken = default)
	{
		await authLock.WaitAsync(cancellationToken);
		try
		{
			IAccount? account = UseDeviceCode
				? authCache?.Account ?? (await PublicClientApp.GetAccountsAsync()).FirstOrDefault()
				: PublicClientApplication.OperatingSystemAccount;
			if (account == null)
				return null;
			authCache = await PublicClientApp.AcquireTokenSilent(scopes ?? DefaultScopes, account)
				.WithForceRefresh(true)
				.ExecuteAsync(cancellationToken);
			return authCache;
		}
		catch (MsalException)
		{
			return null;
		}
		finally
		{
			authLock.Release();
		}
	}

	/// <summary>
	/// Acquires a brand-new token through a full sign-in, ignoring any cached or refreshable token.
	/// Device-code mode calls <paramref name="showCode"/> with the code to display; broker mode shows an account picker.
	/// </summary>
	public async Task<AuthenticationResult> AcquireNewTokenAsync(Action<DeviceCodeResult>? showCode = null,
		string[]? scopes = null, CancellationToken cancellationToken = default)
	{
		scopes ??= DefaultScopes;
		if (UseDeviceCode)
			return await SignInWithDeviceCodeAsync(showCode ?? (_ => { }), scopes, cancellationToken);

		await authLock.WaitAsync(cancellationToken);
		try
		{
			authCache = await PublicClientApp.AcquireTokenInteractive(scopes)
				.WithParentActivityOrWindow(hwnd)
				.WithPrompt(Prompt.SelectAccount)
				.ExecuteAsync(cancellationToken);
			return authCache;
		}
		finally
		{
			authLock.Release();
		}
	}

	/// <summary>Signs in interactively through the Windows broker (non-device-code mode).</summary>
	public Task<AuthenticationResult?> SignInAsync(string[]? scopes = null, CancellationToken cancellationToken = default) =>
		GetAuthAsync(scopes ?? DefaultScopes, cancellationToken);

	public Task<GraphResponse> GetAsync(string endpoint, string scope, CancellationToken cancellationToken) =>
		SendAsync(HttpMethod.Get, endpoint, scope, null, cancellationToken);

	public Task<GraphResponse> PostAsync(string endpoint, JsonElement body, string scope, CancellationToken cancellationToken) =>
		SendAsync(HttpMethod.Post, endpoint, scope, body, cancellationToken);

	private static bool CanReuse(AuthenticationResult? auth, string[] scopes) =>
		auth != null && auth.ExpiresOn > DateTimeOffset.UtcNow.AddMinutes(1) &&
		scopes.All(s => auth.Scopes.Contains(s, StringComparer.OrdinalIgnoreCase));

	private async Task<AuthenticationResult?> GetAuthAsync(string[] scopes, CancellationToken cancellationToken)
	{
		await authLock.WaitAsync(cancellationToken);
		try
		{
			if (CanReuse(authCache, scopes))
				return authCache;

			IAccount? account = UseDeviceCode
				? authCache?.Account ?? (await PublicClientApp.GetAccountsAsync()).SingleOrDefault()
				: PublicClientApplication.OperatingSystemAccount;
			if (account == null)
				return null; // Device-code sign-in is initiated explicitly by the application.

			try
			{
				authCache = await PublicClientApp.AcquireTokenSilent(scopes, account).ExecuteAsync(cancellationToken);
			}
			catch (MsalUiRequiredException)
			{
				if (UseDeviceCode)
					return null; // Needs the application's explicit sign-in.
				try
				{
					authCache = await PublicClientApp.AcquireTokenInteractive(scopes)
						.WithAccount(account)
						.WithParentActivityOrWindow(hwnd)
						.WithPrompt(Prompt.NoPrompt)
						.ExecuteAsync(cancellationToken);
				}
				catch (MsalException)
				{
					return null;
				}
			}
			catch (MsalException)
			{
				return null;
			}
			return authCache;
		}
		finally
		{
			authLock.Release();
		}
	}

	public async Task<GraphBinaryResponse> GetBytesAsync(string endpoint, string scope, long maxBytes, CancellationToken cancellationToken)
	{
		var uri = ValidateEndpoint(endpoint);
		var auth = await GetAuthAsync([scope], cancellationToken);
		if (auth == null)
			return new GraphBinaryResponse([], string.Empty, System.Net.HttpStatusCode.Unauthorized);

		using var request = new HttpRequestMessage(HttpMethod.Get, uri);
		request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", auth.AccessToken);
		using var response = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
		var contentType = response.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
		if (!response.IsSuccessStatusCode)
			return new GraphBinaryResponse([], contentType, response.StatusCode);
		if (response.Content.Headers.ContentLength > maxBytes)
			throw new InvalidDataException($"The content is larger than the {maxBytes} byte limit.");

		await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
		using var buffer = new MemoryStream();
		var chunk = new byte[81920];
		int read;
		while ((read = await stream.ReadAsync(chunk, cancellationToken)) > 0)
		{
			if (buffer.Length + read > maxBytes)
				throw new InvalidDataException($"The content is larger than the {maxBytes} byte limit.");
			buffer.Write(chunk, 0, read);
		}
		return new GraphBinaryResponse(buffer.ToArray(), contentType, response.StatusCode);
	}

	private static Uri ValidateEndpoint(string endpoint)
	{
		var uri = new Uri(GraphBase, endpoint);
		if (uri.Scheme != Uri.UriSchemeHttps || uri.Host != GraphBase.Host || !uri.IsDefaultPort ||
			uri.UserInfo.Length != 0 || uri.Fragment.Length != 0 ||
			!uri.AbsolutePath.StartsWith("/v1.0/", StringComparison.Ordinal))
			throw new ArgumentException("Only Microsoft Graph v1.0 endpoints are allowed.", nameof(endpoint));
		return uri;
	}

	private async Task<GraphResponse> SendAsync(HttpMethod method, string endpoint, string scope, JsonElement? body,
		CancellationToken cancellationToken)
	{
		var uri = ValidateEndpoint(endpoint);

		var auth = await GetAuthAsync([scope], cancellationToken);
		if (auth == null)
			return new GraphResponse(string.Empty, System.Net.HttpStatusCode.Unauthorized);

		using var request = new HttpRequestMessage(method, uri);
		request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", auth.AccessToken);
		if (body.HasValue)
			request.Content = new StringContent(body.Value.GetRawText(), Encoding.UTF8, "application/json");
		using var response = await Http.SendAsync(request, cancellationToken);
		var content = await response.Content.ReadAsStringAsync(cancellationToken);
		return new GraphResponse(content, response.StatusCode);
	}
}
