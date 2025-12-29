using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using Virtual_Desktop_Manager.Core.Models;

namespace Virtual_Desktop_Manager.Core.Helpers
{
	/// <summary>
	/// Provides methods to interact with desktop icons using Windows Shell APIs.
	/// </summary>
	internal static class DesktopIconHelper
	{
		/// <summary>
		/// Gets the current positions of all desktop icons across all screens.
		/// </summary>
		/// <returns>A list of icon positions.</returns>
		public static List<IconPosition> GetIconPositions()
		{
			var positions = new List<IconPosition>();

			try
			{
				// Get the desktop ListView handle
				IntPtr desktopHandle = GetDesktopListViewHandle();
				if (desktopHandle == IntPtr.Zero)
				{
					Debug.WriteLine("[DesktopIconHelper] Failed to get desktop ListView handle");
					return positions;
				}

				Debug.WriteLine($"[DesktopIconHelper] Desktop handle: 0x{desktopHandle:X}");

				// Get the count of icons
				IntPtr result = SendMessage(desktopHandle, LVM_GETITEMCOUNT, IntPtr.Zero, IntPtr.Zero);
				int iconCount = result.ToInt32();

				Debug.WriteLine($"[DesktopIconHelper] Icon count: {iconCount}");

				if (iconCount == 0)
				{
					Debug.WriteLine("[DesktopIconHelper] No icons found on desktop");
					return positions;
				}

				// Get screen information
				var screens = GetScreenConfiguration().Screens;

				// Iterate through each icon
				for (int i = 0; i < iconCount; i++)
				{
					try
					{
						string name = GetIconText(desktopHandle, i);
						Point position = GetIconPosition(desktopHandle, i);

						if (string.IsNullOrEmpty(name))
						{
							Debug.WriteLine($"[DesktopIconHelper] Icon {i}: Empty name, skipping");
							continue;
						}

						if (position == Point.Empty)
						{
							Debug.WriteLine($"[DesktopIconHelper] Icon {i} ({name}): Invalid position, skipping");
							continue;
						}

						// Determine which screen this icon is on
						var screenInfo = DetermineScreen(position, screens);

						positions.Add(new IconPosition
						{
							Name = name,
							X = position.X,
							Y = position.Y,
							ScreenIndex = screenInfo.Index,
							IsPrimaryScreen = screenInfo.IsPrimary
						});

						Debug.WriteLine($"[DesktopIconHelper] Icon {i}: {name} at ({position.X}, {position.Y}) on screen {screenInfo.Index}");
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"[DesktopIconHelper] Error processing icon {i}: {ex.Message}");
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[DesktopIconHelper] GetIconPositions error: {ex.Message}");
			}

			return positions;
		}

		/// <summary>
		/// Sets the positions of desktop icons.
		/// </summary>
		/// <param name="positions">The list of icon positions to apply.</param>
		/// <param name="currentScreenConfig">The current screen configuration.</param>
		/// <param name="savedScreenConfig">The screen configuration when the layout was saved.</param>
		public static void SetIconPositions(List<IconPosition> positions, ScreenConfiguration currentScreenConfig, ScreenConfiguration savedScreenConfig)
		{
			try
			{
				IntPtr desktopHandle = GetDesktopListViewHandle();
				if (desktopHandle == IntPtr.Zero)
				{
					Debug.WriteLine("[DesktopIconHelper] Failed to get desktop ListView handle for setting positions");
					return;
				}

				int iconCount = (int)SendMessage(desktopHandle, LVM_GETITEMCOUNT, IntPtr.Zero, IntPtr.Zero);
				Debug.WriteLine($"[DesktopIconHelper] Setting positions for {positions.Count} icons (desktop has {iconCount} icons)");

				foreach (var iconPos in positions)
				{
					// Find the icon by name
					int iconIndex = FindIconByName(desktopHandle, iconCount, iconPos.Name);
					if (iconIndex == -1)
					{
						Debug.WriteLine($"[DesktopIconHelper] Icon '{iconPos.Name}' not found on desktop");
						continue;
					}

					// Adapt position based on screen configuration changes
					Point newPosition = AdaptPositionForScreenConfig(
						new Point(iconPos.X, iconPos.Y),
						iconPos.ScreenIndex,
						iconPos.IsPrimaryScreen,
						currentScreenConfig,
						savedScreenConfig
					);

					// Set the icon position
					SetIconPosition(desktopHandle, iconIndex, newPosition);
					Debug.WriteLine($"[DesktopIconHelper] Moved '{iconPos.Name}' to ({newPosition.X}, {newPosition.Y})");
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[DesktopIconHelper] SetIconPositions error: {ex.Message}");
			}
		}

		/// <summary>
		/// Gets the current screen configuration.
		/// </summary>
		public static ScreenConfiguration GetScreenConfiguration()
		{
			var config = new ScreenConfiguration();
			var screens = Screen.AllScreens;

			for (int i = 0; i < screens.Length; i++)
			{
				var screen = screens[i];
				config.Screens.Add(new ScreenInfo
				{
					Index = i,
					IsPrimary = screen.Primary,
					Width = screen.Bounds.Width,
					Height = screen.Bounds.Height,
					Left = screen.Bounds.Left,
					Top = screen.Bounds.Top
				});

				if (screen.Primary)
					config.PrimaryScreenIndex = i;
			}

			return config;
		}

		#region Private Helper Methods

		private static IntPtr GetDesktopListViewHandle()
		{
			// Try method 1: Progman -> SHELLDLL_DefView -> SysListView32
			IntPtr progman = FindWindow("Progman", null);
			Debug.WriteLine($"[DesktopIconHelper] Progman handle: 0x{progman:X}");

			if (progman != IntPtr.Zero)
			{
				// Send message to spawn WorkerW window (required for some Windows 10+ systems)
				SendMessage(progman, 0x052C, new IntPtr(0xD), IntPtr.Zero);
				SendMessage(progman, 0x052C, new IntPtr(0xD), new IntPtr(1));

				IntPtr defView = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
				Debug.WriteLine($"[DesktopIconHelper] SHELLDLL_DefView handle: 0x{defView:X}");

				if (defView != IntPtr.Zero)
				{
					IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", "FolderView");
					Debug.WriteLine($"[DesktopIconHelper] SysListView32 handle: 0x{listView:X}");

					if (listView != IntPtr.Zero)
						return listView;
				}
			}

			// Try method 2: Enumerate all windows to find WorkerW -> SHELLDLL_DefView -> SysListView32
			Debug.WriteLine("[DesktopIconHelper] Trying WorkerW enumeration...");

			IntPtr shellDefViewParent = IntPtr.Zero;
			EnumWindows((topHandle, _) =>
			{
				IntPtr defView = FindWindowEx(topHandle, IntPtr.Zero, "SHELLDLL_DefView", null);
				if (defView != IntPtr.Zero)
				{
					Debug.WriteLine($"[DesktopIconHelper] Found SHELLDLL_DefView in window: 0x{topHandle:X}");
					shellDefViewParent = topHandle;
					return false; // Stop enumeration
				}
				return true;
			}, IntPtr.Zero);

			if (shellDefViewParent != IntPtr.Zero)
			{
				IntPtr defView = FindWindowEx(shellDefViewParent, IntPtr.Zero, "SHELLDLL_DefView", null);
				if (defView != IntPtr.Zero)
				{
					IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", "FolderView");
					if (listView != IntPtr.Zero)
					{
						Debug.WriteLine($"[DesktopIconHelper] Found SysListView32 via enumeration: 0x{listView:X}");
						return listView;
					}
				}
			}

			Debug.WriteLine("[DesktopIconHelper] Failed to find desktop ListView");
			return IntPtr.Zero;
		}

		private static string GetIconText(IntPtr listViewHandle, int itemIndex)
		{
			const int maxTextLength = 260;

			GetWindowThreadProcessId(listViewHandle, out IntPtr processId);

			IntPtr hProcess = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, false, processId);
			if (hProcess == IntPtr.Zero)
			{
				Debug.WriteLine($"[DesktopIconHelper] Failed to open process for icon {itemIndex}");
				return string.Empty;
			}

			try
			{
				IntPtr ptrRemoteBuffer = VirtualAllocEx(hProcess, IntPtr.Zero, new UIntPtr(maxTextLength * 2), MEM_COMMIT, PAGE_READWRITE);
				if (ptrRemoteBuffer == IntPtr.Zero)
				{
					Debug.WriteLine($"[DesktopIconHelper] Failed to allocate memory for icon {itemIndex}");
					return string.Empty;
				}

				try
				{
					// Prepare LVITEM structure
					LVITEM lvi = new LVITEM
					{
						mask = LVIF_TEXT,
						iItem = itemIndex,
						iSubItem = 0,
						cchTextMax = maxTextLength,
						pszText = ptrRemoteBuffer
					};

					IntPtr ptrLvi = VirtualAllocEx(hProcess, IntPtr.Zero, new UIntPtr((uint)Marshal.SizeOf(typeof(LVITEM))), MEM_COMMIT, PAGE_READWRITE);
					if (ptrLvi == IntPtr.Zero)
						return string.Empty;

					try
					{
						// Write LVITEM to remote process
						byte[] lviBytes = StructureToBytes(lvi);
						if (!WriteProcessMemory(hProcess, ptrLvi, lviBytes, lviBytes.Length, out _))
							return string.Empty;

						// Send message to get item text
						SendMessage(listViewHandle, LVM_GETITEMTEXT, new IntPtr(itemIndex), ptrLvi);

						// Read the text buffer
						byte[] buffer = new byte[maxTextLength * 2];
						if (!ReadProcessMemory(hProcess, ptrRemoteBuffer, buffer, buffer.Length, out _))
							return string.Empty;

						return System.Text.Encoding.Unicode.GetString(buffer).TrimEnd('\0');
					}
					finally
					{
						VirtualFreeEx(hProcess, ptrLvi, 0, MEM_RELEASE);
					}
				}
				finally
				{
					VirtualFreeEx(hProcess, ptrRemoteBuffer, 0, MEM_RELEASE);
				}
			}
			finally
			{
				CloseHandle(hProcess);
			}
		}

		private static Point GetIconPosition(IntPtr listViewHandle, int itemIndex)
		{
			GetWindowThreadProcessId(listViewHandle, out IntPtr processId);

			IntPtr hProcess = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ, false, processId);
			if (hProcess == IntPtr.Zero)
				return Point.Empty;

			try
			{
				IntPtr ptrRemoteBuffer = VirtualAllocEx(hProcess, IntPtr.Zero, new UIntPtr(16), MEM_COMMIT, PAGE_READWRITE);
				if (ptrRemoteBuffer == IntPtr.Zero)
					return Point.Empty;

				try
				{
					SendMessage(listViewHandle, LVM_GETITEMPOSITION, new IntPtr(itemIndex), ptrRemoteBuffer);

					byte[] buffer = new byte[8];
					if (!ReadProcessMemory(hProcess, ptrRemoteBuffer, buffer, 8, out _))
						return Point.Empty;

					int x = BitConverter.ToInt32(buffer, 0);
					int y = BitConverter.ToInt32(buffer, 4);

					return new Point(x, y);
				}
				finally
				{
					VirtualFreeEx(hProcess, ptrRemoteBuffer, 0, MEM_RELEASE);
				}
			}
			finally
			{
				CloseHandle(hProcess);
			}
		}

		private static void SetIconPosition(IntPtr listViewHandle, int itemIndex, Point position)
		{
			SendMessage(listViewHandle, LVM_SETITEMPOSITION, new IntPtr(itemIndex), MakeLParam(position.X, position.Y));
		}

		private static int FindIconByName(IntPtr listViewHandle, int iconCount, string name)
		{
			for (int i = 0; i < iconCount; i++)
			{
				string iconName = GetIconText(listViewHandle, i);
				if (string.Equals(iconName, name, StringComparison.OrdinalIgnoreCase))
					return i;
			}
			return -1;
		}

		private static ScreenInfo DetermineScreen(Point position, List<ScreenInfo> screens)
		{
			foreach (var screen in screens)
			{
				if (position.X >= screen.Left && position.X < screen.Left + screen.Width &&
					position.Y >= screen.Top && position.Y < screen.Top + screen.Height)
				{
					return screen;
				}
			}

			// Default to primary screen if not found
			return screens.Find(s => s.IsPrimary) ?? screens[0];
		}

		private static Point AdaptPositionForScreenConfig(
			Point originalPosition,
			int originalScreenIndex,
			bool wasOnPrimaryScreen,
			ScreenConfiguration currentConfig,
			ScreenConfiguration savedConfig)
		{
			// Strategy: If the screen configuration changed, try to place icons intelligently

			// Find the target screen in current config
			ScreenInfo targetScreen;

			if (wasOnPrimaryScreen)
			{
				// Always place on current primary screen
				targetScreen = currentConfig.Screens[currentConfig.PrimaryScreenIndex];
			}
			else
			{
				// Try to find equivalent screen, fallback to primary
				if (originalScreenIndex < currentConfig.Screens.Count)
					targetScreen = currentConfig.Screens[originalScreenIndex];
				else
					targetScreen = currentConfig.Screens[currentConfig.PrimaryScreenIndex];
			}

			// Scale position if screen resolution changed
			var originalScreen = originalScreenIndex < savedConfig.Screens.Count
				? savedConfig.Screens[originalScreenIndex]
				: savedConfig.Screens[savedConfig.PrimaryScreenIndex];

			int scaledX = (int)(originalPosition.X * ((double)targetScreen.Width / originalScreen.Width));
			int scaledY = (int)(originalPosition.Y * ((double)targetScreen.Height / originalScreen.Height));

			// Adjust for screen offset
			scaledX += targetScreen.Left;
			scaledY += targetScreen.Top;

			return new Point(scaledX, scaledY);
		}

		private static IntPtr MakeLParam(int loWord, int hiWord)
		{
			return (IntPtr)((hiWord << 16) | (loWord & 0xFFFF));
		}

		private static byte[] StructureToBytes<T>(T structure)
		{
			int size = Marshal.SizeOf(structure);
			byte[] bytes = new byte[size];
			IntPtr ptr = Marshal.AllocHGlobal(size);
			try
			{
				if (structure == null)
					throw new ArgumentNullException(nameof(structure));
				Marshal.StructureToPtr(structure, ptr, false);
				Marshal.Copy(ptr, bytes, 0, size);
			}
			finally
			{
				Marshal.FreeHGlobal(ptr);
			}
			return bytes;
		}

		#endregion

		#region Native API Declarations

		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string? className, string? windowTitle);

		[DllImport("user32.dll")]
		private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out IntPtr lpdwProcessId);

		[DllImport("kernel32.dll")]
		private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, IntPtr dwProcessId);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool CloseHandle(IntPtr hObject);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, UIntPtr dwSize, uint flAllocationType, uint flProtect);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, int dwSize, uint dwFreeType);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, int nSize, out int lpNumberOfBytesWritten);

		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, int dwSize, out int lpNumberOfBytesRead);

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

		private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

		private const uint PROCESS_VM_OPERATION = 0x0008;
		private const uint PROCESS_VM_READ = 0x0010;
		private const uint PROCESS_VM_WRITE = 0x0020;
		private const uint MEM_COMMIT = 0x1000;
		private const uint MEM_RELEASE = 0x8000;
		private const uint PAGE_READWRITE = 0x04;

		private const uint LVM_FIRST = 0x1000;
		private const uint LVM_GETITEMCOUNT = LVM_FIRST + 4;
		private const uint LVM_GETITEMTEXT = LVM_FIRST + 115;
		private const uint LVM_GETITEMPOSITION = LVM_FIRST + 16;
		private const uint LVM_SETITEMPOSITION = LVM_FIRST + 15;
		private const uint LVM_REDRAWITEMS = LVM_FIRST + 21;

		private const uint LVIF_TEXT = 0x0001;

		[StructLayout(LayoutKind.Sequential)]
		private struct LVITEM
		{
			public uint mask;
			public int iItem;
			public int iSubItem;
			public uint state;
			public uint stateMask;
			public IntPtr pszText;
			public int cchTextMax;
			public int iImage;
			public IntPtr lParam;
		}

		#endregion
	}
}