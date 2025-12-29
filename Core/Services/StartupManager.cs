using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using Virtual_Desktop_Manager.Core.Models;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Manages Windows startup registration for the application.
	/// </summary>
	public class StartupManager
	{
		private const string AppName = "VirtualDesktopManager";
		private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

		private readonly string _exePath;

		/// <summary>
		/// Initializes a new instance of the <see cref="StartupManager"/> class.
		/// </summary>
		/// <param name="paths">Paths configuration.</param>
		public StartupManager(Paths paths)
		{
			_exePath = paths.Exe;
		}

		/// <summary>
		/// Gets whether the application is set to run at Windows startup.
		/// </summary>
		public bool IsStartupEnabled
		{
			get
			{
				try
				{
					using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false);
					var registryValue = key?.GetValue(AppName) as string;

					if (string.IsNullOrEmpty(registryValue))
						return false;

					// Check if the registered path matches our startup executable path
					string cleanPath = registryValue.Trim('"');
					return string.Equals(cleanPath, _exePath, StringComparison.OrdinalIgnoreCase);
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"[StartupManager] Error checking startup status: {ex.Message}");
					return false;
				}
			}
		}

		/// <summary>
		/// Enables the application to run at Windows startup.
		/// Points to the executable in the Common folder.
		/// </summary>
		/// <returns>True if successful, false otherwise.</returns>
		public bool EnableStartup()
		{
			try
			{
				// Verify the startup executable exists in Common folder
				if (!File.Exists(_exePath))
				{
					Debug.WriteLine("[StartupManager] Startup executable not found in Common folder.");
					return false;
				}

				// Register the startup executable in registry
				using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
				key?.SetValue(AppName, $"\"{_exePath}\"");

				Debug.WriteLine($"[StartupManager] Startup enabled: {_exePath}");
				return true;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[StartupManager] Error enabling startup: {ex.Message}");
				return false;
			}
		}

		/// <summary>
		/// Disables the application from running at Windows startup.
		/// </summary>
		/// <returns>True if successful, false otherwise.</returns>
		public bool DisableStartup()
		{
			try
			{
				using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
				if (key?.GetValue(AppName) != null)
				{
					key.DeleteValue(AppName);
					Debug.WriteLine("[StartupManager] Startup disabled");
				}
				return true;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[StartupManager] Error disabling startup: {ex.Message}");
				return false;
			}
		}

		/// <summary>
		/// Toggles the startup status.
		/// </summary>
		/// <returns>The new startup status.</returns>
		public bool ToggleStartup()
		{
			if (IsStartupEnabled)
			{
				DisableStartup();
				return false;
			}
			else
			{
				EnableStartup();
				return true;
			}
		}
	}
}