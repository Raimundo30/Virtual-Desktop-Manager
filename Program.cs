using Virtual_Desktop_Manager.Core;
using Virtual_Desktop_Manager.UI.AppWindow;
using System.Windows.Forms;

namespace Virtual_Desktop_Manager
{
	/// <summary>
	/// Main entry point class for the Virtual Desktop Manager application.
	/// </summary>
	internal static class Program
	{
		private static AppCore? _core;
		private static NotifyIcon? _trayIcon;

		/// <summary>
		///  The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			// Initialize and start the core application logic
			_core = new AppCore();
			_core.ErrorOccurred += OnCoreError;
			_core.DesktopSwitched += OnDesktopSwitched;
			_core.Start();

			// Set up and run the UI tray icon
			// TO DO.

			// Set up and run the UI title bar
			// TO DO.

			// Set up and run the application window
			// TO DO. 
			//Application.SetHighDpiMode(HighDpiMode.SystemAware);
			//Application.EnableVisualStyles();
			//ApplicationConfiguration.Initialize();
			//Application.Run(new MainForm());

			// Temporary debug mode
			Console.WriteLine("Virtual Desktop Manager Started");
			Console.WriteLine($"Current Desktop: {_core._monitor.GetCurrentDesktopId()}");
			Console.WriteLine("Monitoring virtual desktops. Press any key to exit...");
			Console.ReadKey();

			// Cleanup when app closes
			_core.Dispose();

			Console.WriteLine("Stopped.");
		}

		/// <summary>
		/// Handles errors that occur in the core logic by displaying an error message to the user.
		/// </summary>
		/// <remarks>This method displays a message box to inform the user of the error. It should be called when a
		/// non-recoverable error occurs in the core functionality.</remarks>
		/// <param name="ex">The exception that represents the error encountered in the core logic. Cannot be null.</param>
		private static void OnCoreError(Exception ex)
		{
			// Log to console (temporary - replace with file logging or UI notification)
			Console.WriteLine($"Core Error: {ex.Message}");

			// Uncomment when UI is ready:
			//MessageBox.Show($"Core Error: {ex.Message}", "Virtual Desktop Manager",
			//	MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		/// <summary>
		/// Handles the event when the desktop is switched by logging the new desktop ID.
		/// </summary>
		/// <param name="desktopId">The GUID of the newly active desktop.</param>
		private static void OnDesktopSwitched(Guid desktopId)
		{
			Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Desktop switched to: {desktopId}");
		}
	}
}