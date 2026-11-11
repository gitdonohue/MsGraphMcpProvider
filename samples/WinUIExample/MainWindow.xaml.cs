using Microsoft.UI.Xaml;
using MsGraphMcpProvider;
using Windows.ApplicationModel.DataTransfer;
using Windows.System;

namespace WinUIExample;

public sealed partial class MainWindow : Window
{
	private MsGraphMcpProvider.MsGraphMcpProvider? MsGraphMcpProvider;
	private MsalGraphClient? MsalGraphClient;

	public MainWindow()
	{
		InitializeComponent();
		Closed += async (_, _) => await StopAsync();
	}

	private async void ToggleButton_Click(object sender, RoutedEventArgs e)
	{
		ToggleButton.IsEnabled = false;
		try
		{
			if (MsGraphMcpProvider == null)
				await StartAsync();
			else
				await StopAsync();
		}
		catch (Exception ex)
		{
			await StopAsync();
			StatusText.Text = $"Error: {ex.Message}";
		}
		finally
		{
			ToggleButton.IsEnabled = true;
		}
	}

	private async Task StartAsync()
	{
		var clientId = ClientIdBox.Text.Trim();
		if (clientId.Length == 0)
		{
			StatusText.Text = "Enter a Client Id first.";
			return;
		}

		var useDeviceCode = DeviceCodeCheck.IsChecked == true;
		var port = (int)PortBox.Value;
		var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
		var client = new MsalGraphClient(clientId, TenantIdBox.Text.Trim(), useDeviceCode, hwnd);
		var writeAccess = new TeamsChannelWriteAccess { Enabled = WriteCheck.IsChecked == true };

		var provider = new MsGraphMcpProvider.MsGraphMcpProvider();
		await provider.StartAsync(port, client, writeAccess: writeAccess);
		MsGraphMcpProvider = provider;
		MsalGraphClient = client;

		ToggleButton.Content = "Stop MCP server";
		SignInButton.IsEnabled = useDeviceCode;
		PortBox.IsEnabled = ClientIdBox.IsEnabled = TenantIdBox.IsEnabled = DeviceCodeCheck.IsEnabled = WriteCheck.IsEnabled = false;
		StatusText.Text = $"MCP server running at http://127.0.0.1:{port}/mcp";
		if (useDeviceCode)
			_ = SignInAsync(client);
	}

	private void CopyCodeButton_Click(object sender, RoutedEventArgs e) => CopyCode();

	private void CopyCode()
	{
		var data = new DataPackage();
		data.SetText(CodeText.Text);
		Clipboard.SetContent(data);
	}

	private async void SignInButton_Click(object sender, RoutedEventArgs e)
	{
		if (MsalGraphClient is { UseDeviceCode: true } client)
			await SignInAsync(client);
	}

	private async Task SignInAsync(MsalGraphClient client)
	{
		SignInButton.IsEnabled = false;
		CodePanel.Visibility = Visibility.Collapsed;
		try
		{
			var result = await client.SignInWithDeviceCodeAsync(code => DispatcherQueue.TryEnqueue(async () =>
			{
				StatusText.Text = code.Message;
				CodeText.Text = code.UserCode;
				CodePanel.Visibility = Visibility.Visible;
				CopyCode();
				await Launcher.LaunchUriAsync(new Uri(code.VerificationUrl.ToString()));
			}));
			if (MsalGraphClient == client)
				StatusText.Text = $"Signed in as {result.Account.Username}. MCP tools are ready.";
		}
		catch (Exception ex)
		{
			if (MsalGraphClient == client)
				StatusText.Text = $"Sign-in failed: {ex.Message}";
		}
		finally
		{
			if (MsalGraphClient == client)
			{
				SignInButton.IsEnabled = true;
				CodePanel.Visibility = Visibility.Collapsed;
			}
		}
	}

	private async Task StopAsync()
	{
		var provider = MsGraphMcpProvider;
		MsGraphMcpProvider = null;
		MsalGraphClient = null;
		if (provider != null)
			await provider.DisposeAsync();

		ToggleButton.Content = "Start MCP server";
		PortBox.IsEnabled = ClientIdBox.IsEnabled = TenantIdBox.IsEnabled = DeviceCodeCheck.IsEnabled = WriteCheck.IsEnabled = true;
		SignInButton.IsEnabled = false;
		CodePanel.Visibility = Visibility.Collapsed;
		StatusText.Text = "Server stopped.";
	}
}
