using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Virtual_Desktop_Manager.Core.Events;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Handles uninstallation of the Virtual Desktop Manager application.
	/// Restores the default desktop, removes startup entries, and prepares for exit.
	/// </summary>
	public class UninstallManager
	{
		private readonly AppCore _core;

		/// <summary>
		/// Occurs when a notification should be displayed to the user.
		/// </summary>
		public event EventHandler<NotificationEventArgs>? Notification;

		public UninstallManager(AppCore core)
		{
			_core = core;
		}

		/// <summary>
		/// Performs complete uninstallation by stopping monitoring, restoring the default desktop, 
		/// removing startup entry, cleaning up unused desktops, and exiting the application.
		/// Prompts the user for confirmation before proceeding.
		/// </summary>
		public void UninstallAndExit()
		{
			var result = MessageBox.Show(
				"This will:\n" +
				"• Stop desktop monitoring\n" +
				"• Restore the default Windows Desktop\n" +
				"• Remove the application from Windows startup\n" +
				"• Close the application\n\n" +
				"Your desktop folders and layouts will NOT be deleted.\n\n" +
				"Do you want to continue?",
				"Uninstall Virtual Desktop Manager",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning);

			if (result == DialogResult.Yes)
			{
				try
				{
					// Stop monitoring and restore default desktop
					Debug.WriteLine("[UninstallManager] Stopping monitoring and restoring default desktop");
					_core.Stop();

					// Disable startup
					Debug.WriteLine("[UninstallManager] Disabling startup entry");
					_core.StartupManager.DisableStartup();

					// Clean up installed executable from Common folder
					//Debug.WriteLine("[UninstallManager] Cleaning up Common folder");
					//if (Directory.Exists(_core.Paths.Common))
					//{
					//	try
					//	{
					//		Directory.Delete(_core.Paths.Common, recursive: true);
					//	}
					//	catch (Exception deleteEx)
					//	{
					//		Notification?.Invoke(this, new NotificationEventArgs(
					//			NotificationSeverity.Error,												// = Severity
					//			"Uninstall Error",														// = Source
					//			$"An error occurred while deleting Common folder: {deleteEx.Message}",  // = Message
					//			NotificationDuration.Long,                                              // = Duration
					//			deleteEx                                                                // = Exception
					//		));
					//	}
					//}

					// Perform additional cleanup if needed
					Debug.WriteLine("[UninstallManager] Performing cleanup of active desktops");
					List <Guid> activeDesktops = _core.DesktopService.ListDesktopIds();
					_core.CleanupManager.CleanupAll(activeDesktops);

					// Open Root folder for user
					Debug.WriteLine("[UninstallManager] Opening root folder for user");
					Process.Start("explorer.exe", _core.Paths.Root);
				}
				catch (Exception ex)
				{
					Notification?.Invoke(this, new NotificationEventArgs(
						NotificationSeverity.Error,                                         // = Severity
						"Uninstall Error",						                            // = Source
						$"An error occurred during uninstallation: {ex.Message}",			// = Message
						NotificationDuration.Long,                                          // = Duration
						ex                                                                  // = Exception
					));
				}

				MessageBox.Show(
					"Virtual Desktop Manager has been uninstalled.\n" +
					"The application will now close.",
					"Uninstall Complete",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information);

				Application.Exit();
			}
		}
	}
}
