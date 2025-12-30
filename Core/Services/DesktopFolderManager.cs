using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms.VisualStyles;
using Virtual_Desktop_Manager.Core.Events;
using Virtual_Desktop_Manager.Core.Models;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Manages the Windows Desktop *path* used by Explorer.
	/// - Reads the current Desktop path from the registry
	/// - Detects whether the current Desktop is the original Desktop or one of the Desktops_VDM folders
	/// - Creates folders for virtual desktops if necessary
	/// - Switches the Windows Desktop path (SHSetKnownFolderPath)
	/// - Refreshes the shell so Explorer reflects the change
	/// </summary>
	public class DesktopFolderManager
	{
		/// <summary>
		/// Occurs when a notification should be displayed to the user.
		/// </summary>
		public event EventHandler<NotificationEventArgs>? Notification;

		private readonly string _rootFolderPath;
		private readonly string _defaultDesktopPath;
		private readonly string _iconPath;

		/// <summary>
		/// Special GUID used by SHSetKnownFolderPath for the Desktop folder.
		/// FOLDERID_Desktop = B4BFCC3A-DB2C-424C-B029-7FE99A87C641
		/// </summary>
		private static readonly Guid FOLDERID_Desktop = new Guid("B4BFCC3A-DB2C-424C-B029-7FE99A87C641");

		/// <summary>
		/// Initializes a new instance of the <see cref="DesktopFolderManager"/> class.
		/// </summary>
		/// <param name="paths">Paths configuration.</param>
		public DesktopFolderManager(Paths paths)
		{
			_defaultDesktopPath = paths.DefaultDesktop;
			_rootFolderPath = paths.Root;
			_iconPath = paths.Icon;

			if (!Directory.Exists(_rootFolderPath))
				Directory.CreateDirectory(_rootFolderPath);

			// Set folder icon for root folder
			SetRootFolderIcon();
		}

		/// <summary>
		/// Sets the custom icon for the Desktops_VDM root folder.
		/// Uses SHGetSetFolderCustomSettings to set the icon, which forces an immediate update.
		/// </summary>
		private void SetRootFolderIcon()
		{
			try
			{
				Debug.WriteLine($"[DesktopFolderManager] === SetRootFolderIcon START ===");
				Debug.WriteLine($"[DesktopFolderManager] Root folder: {_rootFolderPath}");
				Debug.WriteLine($"[DesktopFolderManager] Icon path: {_iconPath}");
				Debug.WriteLine($"[DesktopFolderManager] Icon exists: {File.Exists(_iconPath)}");

				if (!File.Exists(_iconPath))
				{
					Debug.WriteLine($"[DesktopFolderManager] ERROR: Icon file not found at: {_iconPath}");
					return;
				}

				// Use SHFOLDERCUSTOMSETTINGS to set the icon
				SHFOLDERCUSTOMSETTINGS fcs = new SHFOLDERCUSTOMSETTINGS
				{
					dwSize = (uint)Marshal.SizeOf(typeof(SHFOLDERCUSTOMSETTINGS)),
					dwMask = FCSM_ICONFILE,
					pszIconFile = _iconPath,
					cchIconFile = 0,
					iIconIndex = 0
				};

				int hr = SHGetSetFolderCustomSettings(ref fcs, _rootFolderPath, FCS_FORCEWRITE);

				if (hr == 0)
				{
					Debug.WriteLine($"[DesktopFolderManager] SHGetSetFolderCustomSettings succeeded");
				}
				else
				{
					Debug.WriteLine($"[DesktopFolderManager] SHGetSetFolderCustomSettings failed with HRESULT: 0x{hr:X8}");
				}
			}
			catch (Exception ex)
			{
				Notification?.Invoke(this, new NotificationEventArgs(
					NotificationSeverity.Warning,                                               // = Severity
					"Icon setup error",                                                         // = Source
					$"An error occurred while setting the desktop folder icon: {ex.Message}",   // = Message
					NotificationDuration.Short,                                                 // = Duration
					ex                                                                          // = Exception
				));
			}
		}

		/// <summary>
		/// Returns the folder path corresponding to the given virtual desktop ID.
		/// Example: C:\Users\<UserName>\Desktops_VDM\Desktop_<GUID>
		/// </summary>
		/// <param name="desktopId">Virtual Desktop ID.</param>
		/// <returns>Full path to the folder for this virtual desktop.</returns>
		public string GetFolderForDesktop(Guid desktopId)
		{
			if (desktopId == Guid.Empty)
				return _defaultDesktopPath;

			string path = Path.Combine(_rootFolderPath, $"Desktop_{desktopId}");

			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);

			return path;
		}

		/// <summary>
		/// Gets the current Windows Desktop path by calling SHGetKnownFolderPath.
		/// If that fails, falls back to Environment.SpecialFolder.Desktop.
		/// </summary>
		/// <returns>The current Windows Desktop path.</returns>
		public string GetCurrentDesktopPath()
		{
			IntPtr outPathPtr = IntPtr.Zero;
			try
			{
				int hr = SHGetKnownFolderPath(FOLDERID_Desktop, 0, IntPtr.Zero, out outPathPtr);
				if (hr == 0 && outPathPtr != IntPtr.Zero)
				{
					string path = Marshal.PtrToStringUni(outPathPtr) ?? string.Empty;
					return Environment.ExpandEnvironmentVariables(path);
				}
				else
				{
					// Fallback to Environment.SpecialFolder.Desktop
					return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				}
			}
			finally
			{
				// Free COM allocated memory
				if (outPathPtr != IntPtr.Zero)
					Marshal.FreeCoTaskMem(outPathPtr);
			}
		}

		/// <summary>
		/// Detects the Desktop ID corresponding to whatever Desktop path is currently active.
		/// - If the current Desktop path is under RootFolderPath and named Desktop_<GUID>, returns that GUID.
		/// - Otherwise returns Guid.Empty which denotes the ORIGINAL desktop.
		/// </summary>
		/// <param name="desktopPath">The current Desktop path to analyze.</param>
		/// <returns>The detected desktop GUID, or Guid.Empty for the original desktop.</returns>
		public Guid GetDesktopId(string desktopPath)
		{
			// If currentPath is null, return Guid.Empty
			if (desktopPath == null)
				return Guid.Empty;

			// Normalize path for comparison
			string normalizedCurrent = Path.GetFullPath(desktopPath).TrimEnd(Path.DirectorySeparatorChar);

			// Expect a folder named Desktop_<GUID> as the last segment
			string last = Path.GetFileName(normalizedCurrent);
			if (last?.StartsWith("Desktop_", StringComparison.OrdinalIgnoreCase) == true)
			{
				string guidPart = last.Substring("Desktop_".Length);
				if (Guid.TryParse(guidPart, out var id))
					return id;
			}
			return Guid.Empty; // Default to original desktop
		}

		/// <summary>
		/// Changes the current Windows Desktop *path* to the specified folder.
		/// This calls SHSetKnownFolderPath for FOLDERID_Desktop.
		/// Note: this updates the known folder path that Explorer uses — it does NOT move files.
		/// </summary>
		/// <param name="folderPath">Target folder path for the Desktop.</param>
		public void SwitchDesktopPath(string folderPath)
		{
			if (string.IsNullOrEmpty(folderPath))
				throw new ArgumentNullException(nameof(folderPath));

			// Ensure the folder exists before switching
			if (!Directory.Exists(folderPath))
				Directory.CreateDirectory(folderPath);

			// Call SHSetKnownFolderPath to change the Desktop known folder to folderPath
			Guid folderId = FOLDERID_Desktop;
			int hr = SHSetKnownFolderPath(ref folderId, 0, IntPtr.Zero, folderPath);
			if (hr != 0)
			{
				var ex = new COMException($"SHSetKnownFolderPath failed with HRESULT 0x{hr:X8}", hr);
				Notification?.Invoke(this, new NotificationEventArgs(
					NotificationSeverity.Error,							// = Severity
					"Desktop Switch Failed",							// = Source
					$"Failed to switch desktop path: {ex.Message}",		// = Message
					NotificationDuration.Short,							// = Duration
					ex													// = Exception
				));
				throw ex;
			}
		}

		/// <summary>
		/// Refreshes the shell/Desktop so Explorer updates icons and shortcuts.
		/// Uses SHChangeNotify with SHCNE_ASSOCCHANGED which forces a refresh of the shell.
		/// </summary>
		public void RefreshDesktop()
		{
			Debug.WriteLine("[DesktopFolderManager] Refreshing desktop...");

			// Method 1: Notify the shell that associations have changed
			SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
			Debug.WriteLine("[DesktopFolderManager] SHChangeNotify sent.");

			// Method 2: Send refresh to Progman (main desktop window)
			IntPtr progmanHandle = FindWindow("Progman", "Program Manager");
			if (progmanHandle != IntPtr.Zero)
			{
				SendMessage(progmanHandle, WM_COMMAND, FCIDM_SHVIEW_REFRESH, IntPtr.Zero);
				Debug.WriteLine("[DesktopFolderManager] Refresh sent to Progman.");
			}

			// Method 3: Find and refresh all Shell_TrayWnd windows (taskbar)
			IntPtr trayHandle = FindWindow("Shell_TrayWnd", null);
			if (trayHandle != IntPtr.Zero)
			{
				SendMessage(trayHandle, WM_COMMAND, FCIDM_SHVIEW_REFRESH, IntPtr.Zero);
				Debug.WriteLine("[DesktopFolderManager] Refresh sent to Shell_TrayWnd.");
			}

			Debug.WriteLine("[DesktopFolderManager] Desktop refresh completed.");
		}

		#region Native interop
		private const uint SHCNE_ASSOCCHANGED = 0x08000000;
		private const uint SHCNE_UPDATEITEM = 0x00002000;
		private const uint SHCNF_IDLIST = 0x0000;
		private const uint WM_COMMAND = 0x0111;
		private const int FCIDM_SHVIEW_REFRESH = 0xA220;

		// Flags for SHFOLDERCUSTOMSETTINGS
		private const uint FCSM_ICONFILE = 0x00000010;
		private const uint FCS_FORCEWRITE = 0x00000002;

		// SHFOLDERCUSTOMSETTINGS structure
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		private struct SHFOLDERCUSTOMSETTINGS
		{
			public uint dwSize;
			public uint dwMask;
			public IntPtr pvid;
			[MarshalAs(UnmanagedType.LPWStr)]
			public string? pszWebViewTemplate;
			public uint cchWebViewTemplate;
			[MarshalAs(UnmanagedType.LPWStr)]
			public string? pszWebViewTemplateVersion;
			[MarshalAs(UnmanagedType.LPWStr)]
			public string? pszInfoTip;
			public uint cchInfoTip;
			public IntPtr pclsid;
			public uint dwFlags;
			[MarshalAs(UnmanagedType.LPWStr)]
			public string? pszIconFile;
			public uint cchIconFile;
			public int iIconIndex;
			[MarshalAs(UnmanagedType.LPWStr)]
			public string? pszLogo;
			public uint cchLogo;
		}

		// SHGetSetFolderCustomSettings
		[DllImport("shell32.dll", CharSet = CharSet.Unicode)]
		private static extern int SHGetSetFolderCustomSettings(ref SHFOLDERCUSTOMSETTINGS pfcs, string pszPath, uint dwReadWrite);

		// SHGetKnownFolderPath
		[DllImport("shell32.dll")]
		private static extern int SHGetKnownFolderPath(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);

		// SHSetKnownFolderPath
		[DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
		private static extern int SHSetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, string pszPath);

		// SHChangeNotify
		[DllImport("shell32.dll")]
		private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

		// FindWindow
		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

		// SendMessage
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

		#endregion
	}
}