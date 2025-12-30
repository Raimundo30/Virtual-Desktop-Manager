using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Virtual_Desktop_Manager.Core.Events;
using Virtual_Desktop_Manager.Core.Models;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Cleans up unused virtual desktop folders and orphan icon layout files.
	/// </summary>
	public class CleanupManager
	{
		private readonly string _desktopsFolder;
		private readonly string _layoutsFolder;
		private readonly string _binFolder;

		/// <summary>
		/// Occurs when a notification should be displayed to the user.
		/// </summary>
		public event EventHandler<NotificationEventArgs>? Notification;

		/// <summary>
		/// Initializes a new instance of the <see cref="CleanupManager"/> class.
		/// </summary>
		/// <param name="paths">Paths configuration.</param>
		public CleanupManager(Paths paths)
		{
			_desktopsFolder = paths.Root;
			_binFolder = paths.Bin;
			_layoutsFolder = paths.Layouts;
		}

		/// <summary>
		/// Deletes or moves unused virtual desktop folders.
		/// - Only Desktop_<GUID> folders are processed (system folders like Bin, Layouts, Common are ignored)
		/// - Folders corresponding to active desktops are kept
		/// - Empty folders are deleted
		/// - Non-empty folders have their contents moved to Bin, then the empty folder is deleted
		/// </summary>
		/// <param name="activeDesktopIds">List of currently active desktop IDs.</param>
		public void CleanupUnusedDesktopFolders(List<Guid> activeDesktopIds)
		{
			// Check if desktops folder exists
			if (!Directory.Exists(_desktopsFolder))
				return;

			// Ensure the Bin folder exists
			if (!Directory.Exists(_binFolder))
				Directory.CreateDirectory(_binFolder);

			var desktopFolders = Directory.GetDirectories(_desktopsFolder);

			foreach (var folder in desktopFolders)
			{
				string folderName = Path.GetFileName(folder);

				// Expecting folder names like "Desktop_<GUID>"
				if (!folderName.StartsWith("Desktop_"))
					continue;

				string guidPart = folderName.Substring("Desktop_".Length);
				if (!Guid.TryParse(guidPart, out Guid desktopId))
					continue;

				// Skip active desktops
				if (activeDesktopIds.Contains(desktopId))
					continue;

				// Check if folder is empty
				if (Directory.GetFileSystemEntries(folder).Length != 0)
				{
					// Move each item to Bin
					var entries = Directory.GetFileSystemEntries(folder);
					foreach (var entry in entries)
					{
						string destPath = Path.Combine(_binFolder, Path.GetFileName(entry));
						
						// Ensure destination does not overwrite an existing file/folder
						int counter = 1;
						string originalDest = destPath;
						while (File.Exists(destPath) || Directory.Exists(destPath))
						{
							destPath = originalDest + $"_{counter}";
							counter++;
						}

						// Move file or directory
						if (File.Exists(entry))
							File.Move(entry, destPath);
						else if (Directory.Exists(entry))
							Directory.Move(entry, destPath);
					}
				}

				// Safe to delete empty folder
				Directory.Delete(folder, false);
			}
		}

		/// <summary>
		/// Deletes icon layout files that do not correspond to any active desktop.
		/// Preserves the layout for the original desktop (Guid.Empty).
		/// </summary>
		/// <param name="activeDesktopIds">List of currently active desktop IDs.</param>
		public void CleanupUnusedLayoutFiles(List<Guid> activeDesktopIds)
		{
			if (!Directory.Exists(_layoutsFolder))
				return;

			var files = Directory.GetFiles(_layoutsFolder, "*.json");

			foreach (var file in files)
			{
				string fileName = Path.GetFileNameWithoutExtension(file);

				// Expecting file names like "Layout_<GUID>.json"
				if (!fileName.StartsWith("Layout_"))
					continue;

				string guidPart = fileName.Substring("Layout_".Length);
				if (!Guid.TryParse(guidPart, out Guid desktopId))
					continue;

				// ALWAYS preserve the original desktop layout (Guid.Empty)
				if (desktopId == Guid.Empty)
					continue;

				// Delete layout files not corresponding to active desktops
				if (!activeDesktopIds.Contains(desktopId))
					File.Delete(file);
			}
		}

		/// <summary>
		/// Runs full cleanup: folders + layout files.
		/// </summary>
		/// <param name="activeDesktopIds">List of currently active desktop IDs.</param>
		public void CleanupAll(List<Guid> activeDesktopIds)
		{
			try
			{
				CleanupUnusedDesktopFolders(activeDesktopIds);
				CleanupUnusedLayoutFiles(activeDesktopIds);
			}
			catch (Exception ex)
			{
				Notification?.Invoke(this, new NotificationEventArgs(
					NotificationSeverity.Warning,                       // = Severity
					"Cleanup Error",                                    // = Source
					$"An error occurred during cleanup: {ex.Message}",  // = Message
					NotificationDuration.Long,                          // = Duration
					ex                                                  // = Exception
				));
			}
		}
	}
}