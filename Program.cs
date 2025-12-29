using Microsoft.Toolkit.Uwp.Notifications;
using System.Diagnostics;
using System.Windows.Forms;
using Virtual_Desktop_Manager.Core;
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
		private static NotificationManager? _notify;

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

			// Initialize and start the core application logic
			_core = new AppCore();
			_core.ErrorOccurred += OnCoreError;
			_core.Message += OnMessage;
			_core.DesktopSwitched += OnDesktopSwitched;
			_core.Start();

			// Set up the tray icon
			_trayIcon = new TrayIcon(_core);
			_trayIcon.ErrorOccurred += OnCoreError;
			_trayIcon.Message += OnMessage;

			// Set up notification manager
			_notify = new NotificationManager();

			// Run the application (keeps it alive until 'Exit' is called)
			Application.Run();

			// Cleanup when app closes
			_core.Dispose();
			_trayIcon.Dispose();
			_notify.Dispose();
		}

		/// <summary>
		/// Handles errors that occur in the core logic by displaying an error message to the user.
		/// </summary>
		/// <param name="ex">The exception that represents the error encountered in the core logic. Cannot be null.</param>
		private static void OnCoreError(Exception ex)
		{
			// Show a toast notification for the error
			_notify?.ShowToast("Error", ex.Message, NotificationManager.ToastDuration.Long, "❌ ");

			// Also log to Debug output
			Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] Error: {ex.Message}");
		}

		/// <summary>
		/// Handles informational messages from the core logic by displaying them to the user.
		/// </summary>
		/// <param name="message">The informational message from the core logic. Cannot be null or empty.</param>
		private static void OnMessage(string type, string source, string message)
		{
			// Show a toast notification for the message
			_notify?.ShowToast(source, message, NotificationManager.ToastDuration.Short, type + " ");
			
			// Also log to Debug output
			Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] [{source}] {message}");
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