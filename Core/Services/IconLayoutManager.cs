using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.Json;
using Virtual_Desktop_Manager.Core.Helpers;
using Virtual_Desktop_Manager.Core.Models;

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
		/// Saves the icon layout for the current desktop.
		/// Captures icon positions across all screens.
		/// </summary>
		public void SaveLayout()
		{
			try
			{
				string filePath = GetLayoutFilePath();

				// Capture current icon positions
				var iconPositions = DesktopIconHelper.GetIconPositions();
				Debug.WriteLine($"[IconLayoutManager] Captured {iconPositions.Count} icons");

				// Capture current screen configuration
				var screenConfig = DesktopIconHelper.GetScreenConfiguration();
				Debug.WriteLine($"[IconLayoutManager] Screen config: {screenConfig.Screens.Count} screen(s)");

				// Create layout object
				var layout = new DesktopLayout
				{
					Icons = iconPositions,
					ScreenConfig = screenConfig,
					SavedAt = DateTime.Now
				};

				// Serialize to JSON
				var options = new JsonSerializerOptions
				{
					WriteIndented = true
				};
				string json = JsonSerializer.Serialize(layout, options);

				// Save to file
				File.WriteAllText(filePath, json);
				Debug.WriteLine($"[IconLayoutManager] Saved layout to: {Path.GetFileName(filePath)}");
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[IconLayoutManager] SaveLayout error: {ex.Message}");
				ErrorOccurred?.Invoke(ex);
			}
		}

		/// <summary>
		/// Loads the icon layout for the current desktop if it exists.
		/// Adapts positions for current screen configuration.
		/// </summary>
		public void LoadLayout()
		{
			try
			{
				string filePath = GetLayoutFilePath();
				if (!File.Exists(filePath))
				{
					Debug.WriteLine($"[IconLayoutManager] No layout file found: {Path.GetFileName(filePath)}");
					return;
				}

				// Read and deserialize layout
				string json = File.ReadAllText(filePath);
				var layout = JsonSerializer.Deserialize<DesktopLayout>(json);

				if (layout == null || layout.Icons.Count == 0)
					return;

				// Get current screen configuration
				var currentScreenConfig = DesktopIconHelper.GetScreenConfiguration();

				// Apply icon positions with screen configuration adaptation
				DesktopIconHelper.SetIconPositions(
					layout.Icons,
					currentScreenConfig,
					layout.ScreenConfig
				);

				Debug.WriteLine($"[IconLayoutManager] Layout loaded successfully");
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[IconLayoutManager] LoadLayout error: {ex.Message}");
				ErrorOccurred?.Invoke(ex);
			}
		}

		/// <summary>
		/// Determines the correct filename for the layout.
		/// </summary>
		private string GetLayoutFilePath()
		{
			string desktopPath = _folderManager.GetCurrentDesktopPath();
			Guid desktopId = _folderManager.GetDesktopId(desktopPath);

			return Path.Combine(LayoutRoot, $"Layout_{desktopId}.json");
		}
	}
}