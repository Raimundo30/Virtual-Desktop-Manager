using Microsoft.Toolkit.Uwp.Notifications;
using System.Diagnostics;
using System.Windows.Forms;
using Virtual_Desktop_Manager.Core;
using Virtual_Desktop_Manager.Core.Events;
using Virtual_Desktop_Manager.Core.Models;
using Virtual_Desktop_Manager.Core.Services;
using Virtual_Desktop_Manager.UI.AppWindow;
using Virtual_Desktop_Manager.UI.Notifications;
using Virtual_Desktop_Manager.UI.TrayIcon;

namespace Virtual_Desktop_Manager
{
	/// <summary>
	/// Main entry point class for the Virtual Desktop Manager application.
	/// </summary>
	internal static class Program
	{
		private static AppCore? _core;
		private static TrayIcon? _trayIcon;
		private static NotificationManager? _notification;

		/// <summary>
		///  The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			// Initialize paths
			var paths = new Paths();

			// Run installer if not running from Common folder
			// NOTE: for development purposes, you need to disable the installer
			if (InstallationManager.RunInstaller(paths))
			{
				// Installer launched the installed version, exit this process
				Debug.WriteLine("[Program] Installer executed, exiting original process");
				return;
			}
			Debug.WriteLine("[Program] Starting application from Common folder");

			// Initialize Windows Forms
			Application.SetHighDpiMode(HighDpiMode.SystemAware);
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			// Set up notification manager
			_notification = new NotificationManager();

			// Initialize and start the core application logic
			_core = new AppCore();
			_core.Notification += OnNotificationRequested;
			_core.DesktopSwitched += OnDesktopSwitched;
			_core.Start();

			// Set up the tray icon
			_trayIcon = new TrayIcon(_core);
			_trayIcon.Notification += OnNotificationRequested;

			// Run the application (keeps it alive until 'Exit' is called)
			Application.Run();

			// Cleanup when app closes
			_core.Dispose();
			_trayIcon.Dispose();
			_notification.Dispose();
		}

		/// <summary>
		/// Handles notification requests from the core and tray icon by displaying a toast notification
		/// </summary>
		/// <param name="sender">The source of the notification event.</param>
		/// <param name="e">The notification event arguments containing details about the notification.</param>
		private static void OnNotificationRequested(object? sender, NotificationEventArgs e)
		{
			_notification?.ShowToast(e);

			var level = e.Severity == NotificationSeverity.Error ? "ERROR" : e.Severity.ToString().ToUpperInvariant();
			Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] [{level}] [{e.Source}] {e.Message}");
			
			if (e.Exception != null)
			{
				Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] Exception: {e.Exception}");
			}
		}

		/// <summary>
		/// Handles the event when the desktop is switched by logging the new desktop ID.
		/// </summary>
		/// <param name="desktopId">The GUID of the newly active desktop.</param>
		private static void OnDesktopSwitched(Guid desktopId)
		{
			// Log the desktop switch event
			Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] Desktop switched to: {desktopId}");
		}
	}
}