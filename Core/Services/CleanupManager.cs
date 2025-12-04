using System;
using System.Collections.Generic;
using System.IO;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Cleans up unused virtual desktop folders and orphan icon layout files.
	/// </summary>
	public class CleanupManager
	{
		/// <summary>
		/// Occurs when an error is encountered during operation.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Root folder where all virtual desktop folders are stored.
		/// Example: %UserProfile%\Desktops_VDM
		/// </summary>
		public string DesktopRoot { get; }

		/// <summary>
		/// Root folder where icon layout files are stored.
		/// Example:  %UserProfile%\Desktops_VDM\Layouts
		/// </summary>
		public string LayoutRoot { get; }

		/// <summary>
		/// Initializes a new instance of the <see cref="CleanupManager"/> class.
		/// </summary>
		/// <param name="desktopRoot">Root folder of virtual desktops.</param>
		/// <param name="layoutRoot">Root folder of saved icon layouts.</param>
		public CleanupManager(string desktopRoot, string layoutRoot)
		{
			DesktopRoot = desktopRoot;
			LayoutRoot = layoutRoot;
		}

		/// <summary>
		/// Deletes or moves unused virtual desktop folders.
		/// - Folders corresponding to active desktops are kept
		/// - Empty folders are deleted
		/// - Non-empty folders are moved to DesktopRoot\Bin
		/// </summary>
		/// <param name="activeDesktopIds">List of currently active desktop IDs.</param>
		public void CleanupUnusedDesktopFolders(List<Guid> activeDesktopIds)
		{
			string binFolder = Path.Combine(DesktopRoot, "Bin");

			// Ensure the Bin folder exists
			if (!Directory.Exists(binFolder))
				Directory.CreateDirectory(binFolder);

			if (!Directory.Exists(DesktopRoot))
				return;

			var desktopFolders = Directory.GetDirectories(DesktopRoot);

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
				if (Directory.GetFileSystemEntries(folder).Length == 0)
				{
					// Safe to delete empty folder
					Directory.Delete(folder, false);
				}
				else
				{
					// Move non-empty folder to Bin
					string dest = Path.Combine(binFolder, folderName);

					// Ensure destination does not overwrite an existing folder
					int counter = 1;
					string originalDest = dest;
					while (Directory.Exists(dest))
					{
						dest = originalDest + $"_{counter}";
						counter++;
					}

					Directory.Move(folder, dest);
				}
			}
		}

		/// <summary>
		/// Deletes icon layout files that do not correspond to any active desktop.
		/// Preserves the layout for the original desktop (Guid.Empty).
		/// </summary>
		/// <param name="activeDesktopIds">List of currently active desktop IDs.</param>
		public void CleanupUnusedLayoutFiles(List<Guid> activeDesktopIds)
		{
			if (!Directory.Exists(LayoutRoot))
				return;

			var files = Directory.GetFiles(LayoutRoot, "*.json");

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
				ErrorOccurred?.Invoke(ex);
			}
		}
	}
}