using System;
using System.IO;
using System.Runtime.InteropServices;

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
		/// Occurs when an error is encountered during operation.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Root folder where all virtual desktop folders are stored.
		/// Example: %UserProfile%\Desktops_VDM
		/// </summary>
		private string RootFolderPath { get; }

		/// <summary>
		/// User profile path, e.g. C:\Users\<UserName>
		/// </summary>
		private readonly string _userProfilePath;

		/// <summary>
		/// Special GUID used by SHSetKnownFolderPath for the Desktop folder.
		/// FOLDERID_Desktop = B4BFCC3A-DB2C-424C-B029-7FE99A87C641
		/// </summary>
		private static readonly Guid FOLDERID_Desktop = new Guid("B4BFCC3A-DB2C-424C-B029-7FE99A87C641");

		/// <summary>
		/// Initializes a new instance of the <see cref="DesktopFolderManager"/> class.
		/// </summary>
		/// <param name="RootFolderPath">Root folder for virtual desktops.</param>
		/// <param name="UserProfilePath">User profile path.</param>
		public DesktopFolderManager(string RootFolderPath, string UserProfilePath)
		{
			_userProfilePath = UserProfilePath;
			this.RootFolderPath = RootFolderPath;

			if (!Directory.Exists(this.RootFolderPath))
				Directory.CreateDirectory(this.RootFolderPath);
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
				return Path.Combine(_userProfilePath, "Desktop");

			string path = Path.Combine(RootFolderPath, $"Desktop_{desktopId}");

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
				ErrorOccurred?.Invoke(ex);
				throw ex;
			}
		}

		/// <summary>
		/// Refreshes the shell/Desktop so Explorer updates icons and shortcuts.
		/// Uses SHChangeNotify with SHCNE_ASSOCCHANGED which forces a refresh of the shell.
		/// </summary>
		public void RefreshDesktop()
		{
			// Notify the shell that associations have changed
			SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);

			// Also send a refresh command to the desktop window
			IntPtr desktopHandle = FindWindow("Progman", "Program Manager");
			if (desktopHandle != IntPtr.Zero)
			{
				SendMessage(desktopHandle, WM_COMMAND, (int)FCIDM_SHVIEW_REFRESH, IntPtr.Zero);
			}

			// Additionally, enumerate all top-level windows and send refresh command
			EnumWindows((hWnd, _) =>
			{
				SendMessage(hWnd, WM_COMMAND, FCIDM_SHVIEW_REFRESH, IntPtr.Zero);
				return true; // continue enumeration
			}, IntPtr.Zero);
		}

		#region Native interop
		private const uint SHCNE_ASSOCCHANGED = 0x08000000;
		private const uint SHCNF_IDLIST = 0x0000;
		private const uint WM_COMMAND = 0x0111;
		private const int FCIDM_SHVIEW_REFRESH = 0xA220;

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
		[DllImport("user32.dll")]
		private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		// EnumWindows
		[DllImport("user32.dll")]
		private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

		// EnumWindows callback delegate
		private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

		// SendMessage
		[DllImport("user32.dll")]
		private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

		#endregion
	}
}