using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Virtual_Desktop_Manager.Core;
using Virtual_Desktop_Manager.Core.Models;
using Virtual_Desktop_Manager.Core.Services;

namespace Virtual_Desktop_Manager.UI.TrayIcon
{
	/// <summary>
	/// Manages the system tray icon and context menu for the application.
	/// </summary>
	public class TrayIcon : IDisposable
	{
		private NotifyIcon? _notifyIcon;
		private readonly AppCore _core;

		/// <summary>
		/// Occurs when an informational message is generated.
		/// </summary>
		public event Action<string, string, string>? Message;

		/// <summary>
		/// Occurs when an error is encountered during the execution of the component.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Initializes the system tray icon with context menu.
		/// </summary>
		public TrayIcon(AppCore core)
		{
			_core = core;

			_notifyIcon = new NotifyIcon
			{
				Text = "Virtual Desktop Manager",
				Visible = true,
				Icon = LoadIcon(_core.Paths.Icon)
			};

			// Build menu on right-click (when opening context menu)
			_notifyIcon.MouseUp += OnNotifyIconMouseUp;

			// Create initial empty menu
			_notifyIcon.ContextMenuStrip = new ContextMenuStrip();
		}

		/// <summary>
		/// Attempts to load an icon from the specified file path. Returns a default application icon if the file cannot be
		/// loaded.
		/// </summary>
		/// <remarks>If the specified file does not exist or cannot be loaded as an icon, the method returns <see
		/// cref="SystemIcons.Application"/>. No exception is thrown if loading fails.</remarks>
		/// <param name="iconPath">The file system path to the icon file to load. Can be a relative or absolute path.</param>
		/// <returns>An <see cref="Icon"/> loaded from the specified file if successful; otherwise, the default application icon.</returns>
		private static Icon LoadIcon(string iconPath)
		{
			try
			{
				if (File.Exists(iconPath))
				{
					return new Icon(iconPath);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[TrayIcon] Failed to load icon: {ex.Message}");
			}

			return SystemIcons.Application;
		}

		/// <summary>
		/// Handles mouse up events on the tray icon.
		/// Rebuilds and shows the menu when right-clicking.
		/// </summary>
		private void OnNotifyIconMouseUp(object? sender, MouseEventArgs e)
		{
			try
			{
				Debug.WriteLine($"[TrayIcon] Mouse up detected: {e.Button}");

				if (e.Button == MouseButtons.Right)
				{
					Debug.WriteLine("[TrayIcon] Right click - building and showing menu");

					// Build the menu
					var contextMenu = BuildContextMenu();

					// Show it at the cursor position
					contextMenu.Show(Cursor.Position);

					Debug.WriteLine("[TrayIcon] Menu shown");
				}
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
		}

		/// <summary>
		/// Builds the context menu with all options.
		/// </summary>
		/// <returns>The built context menu.</returns>
		private ContextMenuStrip BuildContextMenu()
		{
			Debug.WriteLine("[TrayIcon] === BuildContextMenu START ===");
			var contextMenu = new ContextMenuStrip();
			Debug.WriteLine("[TrayIcon] ContextMenuStrip created");

			// 1. Header
			var statusItem = new ToolStripMenuItem("Virtual Desktop Manager")
			{
				Enabled = false
			};
			Debug.WriteLine("[TrayIcon] Header item created");
			contextMenu.Items.Add(statusItem);
			Debug.WriteLine("[TrayIcon] Header item added");
			contextMenu.Items.Add(new ToolStripSeparator());
			Debug.WriteLine("[TrayIcon] Separator added");

			// 2. Switch to Desktop submenu
			var switchDesktopMenu = new ToolStripMenuItem("Switch to Desktop");
			var desktopInfo = _core.DesktopService.GetDesktopInfo();
			var currentDesktopId = _core.DesktopService.LastDesktopId;

			foreach (var kvp in desktopInfo)
			{
				var desktopItem = new ToolStripMenuItem(kvp.Value);
				desktopItem.Tag = kvp.Key;
				desktopItem.Checked = kvp.Key == currentDesktopId;
				desktopItem.Click += OnSwitchDesktop;
				switchDesktopMenu.DropDownItems.Add(desktopItem);
			}

			if (switchDesktopMenu.DropDownItems.Count == 0)
			{
				switchDesktopMenu.DropDownItems.Add(new ToolStripMenuItem("No desktops available") { Enabled = false });
			}

			contextMenu.Items.Add(switchDesktopMenu);
			contextMenu.Items.Add(new ToolStripSeparator());

			// 3. Open Folders submenu
			var openFoldersMenu = new ToolStripMenuItem("Open Folders");

			// Default desktop folder
			var defaultFolderItem = new ToolStripMenuItem("Default Desktop");
			defaultFolderItem.Click += (s, e) => OpenFolder(_core.Paths.DefaultDesktop);
			openFoldersMenu.DropDownItems.Add(defaultFolderItem);

			// Root folder
			var rootFolderItem = new ToolStripMenuItem("Root");
			rootFolderItem.Click += (s, e) => OpenFolder(_core.Paths.Root);
			openFoldersMenu.DropDownItems.Add(rootFolderItem);

			// Bin folder
			var binFolderItem = new ToolStripMenuItem("Bin");
			binFolderItem.Click += (s, e) => OpenFolder(_core.Paths.Bin);
			openFoldersMenu.DropDownItems.Add(binFolderItem);

			openFoldersMenu.DropDownItems.Add(new ToolStripSeparator());

			// Individual desktop folders
			foreach (var kvp in desktopInfo)
			{
				var folderItem = new ToolStripMenuItem(kvp.Value);
				folderItem.Tag = kvp.Key;
				folderItem.Click += OnOpenDesktopFolder;
				openFoldersMenu.DropDownItems.Add(folderItem);
			}

			contextMenu.Items.Add(openFoldersMenu);
			contextMenu.Items.Add(new ToolStripSeparator());

			// 4. Startup toggle
			bool startupEnabled = _core.StartupManager.IsStartupEnabled;
			var startupItem = new ToolStripMenuItem("Start with Windows")
			{
				Text = startupEnabled ? "Start with Windows : ON ✔️" : "Start with Windows : OFF ✖️"
			};
			startupItem.Click += OnToggleStartup;
			contextMenu.Items.Add(startupItem);

			contextMenu.Items.Add(new ToolStripSeparator());

			// 5. Uninstall
			var uninstallItem = new ToolStripMenuItem("Uninstall and Exit");
			uninstallItem.Click += OnUninstallAndExit;
			contextMenu.Items.Add(uninstallItem);

			contextMenu.Items.Add(new ToolStripSeparator());

			// 6. Exit
			Debug.WriteLine("[TrayIcon] Creating exit item");
			var exitItem = new ToolStripMenuItem("Exit");
			Debug.WriteLine("[TrayIcon] Exit item created");
			exitItem.Click += OnExit;
			Debug.WriteLine("[TrayIcon] Exit handler attached");
			contextMenu.Items.Add(exitItem);
			Debug.WriteLine("[TrayIcon] Exit item added to menu");

			Debug.WriteLine("[TrayIcon] === BuildContextMenu COMPLETE ===");

			return contextMenu;
		}

		/// <summary>
		/// Handles switching to a specific desktop.
		/// The desktop switch will be detected by monitoring and trigger the desktop folder change workflow.
		/// </summary>
		/// <param name="sender">The menu item that was clicked.</param>
		/// <param name="e">Event arguments.</param>
		private void OnSwitchDesktop(object? sender, EventArgs e)
		{
			if (sender is ToolStripMenuItem menuItem && menuItem.Tag is Guid desktopId)
			{
				_core.DesktopService.SwitchToDesktop(desktopId);
				// After switching Virtual Desktops, monitoring will detect the change and update the Desktop folder
			}
		}

		/// <summary>
		/// Handles opening a desktop folder in Windows Explorer.
		/// </summary>
		/// <param name="sender">The menu item that was clicked.</param>
		/// <param name="e">Event arguments.</param>
		private void OnOpenDesktopFolder(object? sender, EventArgs e)
		{
			if (sender is ToolStripMenuItem menuItem && menuItem.Tag is Guid desktopId)
			{
				string folderPath = _core.FolderManager.GetFolderForDesktop(desktopId);
				OpenFolder(folderPath);
			}
		}

		/// <summary>
		/// Opens a folder in Windows Explorer.
		/// </summary>
		/// <param name="folderPath">The full path to the folder to open.</param>
		private void OpenFolder(string folderPath)
		{
			try
			{
				if (Directory.Exists(folderPath))
				{
					Process.Start("explorer.exe", folderPath);
				}
				else
				{
					MessageBox.Show($"Folder does not exist:\n{folderPath}", "Folder Not Found", 
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[TrayIcon] Error opening folder: {ex.Message}");
				MessageBox.Show($"Error opening folder:\n{ex.Message}", "Error", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Handles toggling startup with Windows.
		/// </summary>
		private void OnToggleStartup(object? sender, EventArgs e)
		{
			bool newStatus = _core.StartupManager.ToggleStartup();
			string message = newStatus 
				? "Virtual Desktop Manager will now start with Windows." 
				: "Virtual Desktop Manager will no longer start with Windows.";

			Message?.Invoke("ℹ️", "Startup with windows", message);
			Debug.WriteLine("[TrayIcon] " + message);
		}

		/// <summary>
		/// Handles uninstallation and exit.
		/// Delegates to UninstallManager which handles user confirmation, cleanup, and application exit.
		/// </summary>
		/// <param name="sender">The menu item that was clicked.</param>
		/// <param name="e">Event arguments.</param>
		private void OnUninstallAndExit(object? sender, EventArgs e)
		{
			_core.UninstallManager.UninstallAndExit();
		}

		/// <summary>
		/// Handles the exit menu item click by shutting down the application.
		/// </summary>
		private void OnExit(object? sender, EventArgs e)
		{
			Application.Exit();
		}

		/// <summary>
		/// Disposes the tray icon and cleans up resources.
		/// </summary>
		public void Dispose()
		{
			if (_notifyIcon != null)
			{
				_notifyIcon.MouseClick -= OnNotifyIconMouseUp;
				_notifyIcon.Visible = false;
				_notifyIcon.Dispose();
				_notifyIcon = null;
			}
		}
	}
}
