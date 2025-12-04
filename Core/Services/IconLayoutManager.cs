using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Manages saving and loading icon layouts for both
	/// the original Windows Desktop and virtual desktops.
	/// </summary>
	public class IconLayoutManager
	{
		/// <summary>
		/// Occurs when an error is encountered during operation.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Root folder where icon layout files are stored.
		/// Example:  %UserProfile%\Desktops_VDM\Layouts
		/// </summary>
		private string LayoutRoot { get; }

		/// <summary>
		/// Reference to DesktopFolderManager for desktop path and ID retrieval.
		/// </summary>
		private readonly DesktopFolderManager _folderManager;

		/// <summary>
		/// Initializes a new instance of the <see cref="IconLayoutManager"/> class.
		/// </summary>
		/// <param name="layoutRoot">Root folder where icon layout files are stored.</param>
		/// <param name="folderManager">Reference to DesktopFolderManager for desktop path and ID retrieval.</param>
		public IconLayoutManager(string layoutRoot, DesktopFolderManager folderManager)
		{
			LayoutRoot = layoutRoot;
			_folderManager = folderManager;

			if (!Directory.Exists(LayoutRoot))
				Directory.CreateDirectory(LayoutRoot);
		}

		/// <summary>
		/// Saves the icon layout for the given desktop.
		/// The GUID is still passed, but for the original Desktop
		/// a special file name is used instead.
		/// </summary>
		public void SaveLayout()
		{
			try
			{
				string filePath = GetLayoutFilePath();

				// TODO: Capture real icon layout and serialize it

				File.WriteAllText(filePath, "{ /* icon layout placeholder */ }");
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
		}

		/// <summary>
		/// Loads the icon layout for the given desktop if it exists.
		/// The GUID is still passed, but a special file is used
		/// for the original Desktop.
		/// </summary>
		public void LoadLayout()
		{
			try
			{
				string filePath = GetLayoutFilePath();
				if (File.Exists(filePath))
				{
					string layoutJson = File.ReadAllText(filePath);
					// TODO: Implement actual desktop icon positioning.
				}
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
		}

		/// <summary>
		/// Determines the correct filename for the layout:
		/// - "Layout_OriginalDesktop.json" for the original Desktop
		/// - "Layout_<GUID>.json" otherwise
		/// </summary>
		private string GetLayoutFilePath()
		{
			string desktopPath = _folderManager.GetCurrentDesktopPath();
			Guid desktopId = _folderManager.GetDesktopId(desktopPath);

			return Path.Combine(LayoutRoot, $"Layout_{desktopId}.json");
		}
	}
}