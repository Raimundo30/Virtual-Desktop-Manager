using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Virtual_Desktop_Manager.Core.Models;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Handles the installation process by copying the application to the Common folder
	/// and launching it from there.
	/// </summary>
	public static class InstallationManager
	{
		/// <summary>
		/// Installs the application by copying it to the Common folder and launching it.
		/// If not running from Common folder, deletes existing files and performs a fresh installation.
		/// Returns true if the installer process was executed (calling code should exit).
		/// Returns false if no installation is needed (app should continue normally).
		/// </summary>
		/// <param name="paths">Paths configuration.</param>
		/// <returns>True if installer executed and app should exit, false if app should continue.</returns>
		public static bool RunInstaller(Paths paths)
		{
			try
			{
				// Check if already running from Common folder
				if (IsRunningFromCommonFolder(paths))
				{
					Debug.WriteLine("[InstallationManager] Already running from Common folder - no installation needed");
					return false; // Continue normal execution
				}

				Debug.WriteLine("[InstallationManager] Running from original location, performing clean installation...");

				// Ensure Root folder exists
				if (!Directory.Exists(paths.Root))
					Directory.CreateDirectory(paths.Root);

				// Clean existing Common folder for fresh installation
				if (Directory.Exists(paths.Common))
				{
					Debug.WriteLine("[InstallationManager] Deleting existing Common folder...");
					try
					{
						Directory.Delete(paths.Common, recursive: true);
						Debug.WriteLine("[InstallationManager] Existing Common folder deleted");
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"[InstallationManager] Warning: Could not fully delete Common folder: {ex.Message}");
						// Continue anyway - will try to overwrite files
					}
				}

				// Create fresh Common folder
				Directory.CreateDirectory(paths.Common);
				Debug.WriteLine("[InstallationManager] Created fresh Common folder");

				// Get source executable path
				string sourceExePath = GetCurrentExecutablePath();

				// Copy executable and dependencies
				CopyExecutableAndDependencies(sourceExePath, paths.Common);

				Debug.WriteLine($"[InstallationManager] Installation complete. Launching from: {paths.Exe}");

				// Launch the copied executable
				var startInfo = new ProcessStartInfo
				{
					FileName = paths.Exe,
					UseShellExecute = true,
					WorkingDirectory = paths.Common
				};

				Process.Start(startInfo);

				Debug.WriteLine("[InstallationManager] Launched installed version. Exiting installer.");

				// Return true to signal that the original process should exit
				return true;
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[InstallationManager] Installation failed: {ex.Message}");
				Debug.WriteLine($"[InstallationManager] Stack trace: {ex.StackTrace}");
				// On failure, continue with normal execution from original location
				return false;
			}
		}

		/// <summary>
		/// Checks if the application is running from the Common folder.
		/// </summary>
		/// <param name="paths">Paths configuration.</param>
		/// <returns>True if running from Common folder, false otherwise.</returns>
		public static bool IsRunningFromCommonFolder(Paths paths)
		{
			string currentExePath = GetCurrentExecutablePath();

			string normalizedCurrent = Path.GetFullPath(currentExePath).TrimEnd(Path.DirectorySeparatorChar);
			string normalizedCommon = Path.GetFullPath(paths.Exe).TrimEnd(Path.DirectorySeparatorChar);

			return string.Equals(normalizedCurrent, normalizedCommon, StringComparison.OrdinalIgnoreCase);
		}

		/// <summary>
		/// Copies the executable and its dependencies to the target directory.
		/// </summary>
		/// <param name="sourceExePath">Source executable path.</param>
		/// <param name="targetDirectory">Target directory path.</param>
		private static void CopyExecutableAndDependencies(string sourceExePath, string targetDirectory)
		{
			string sourceDir = Path.GetDirectoryName(sourceExePath) ?? string.Empty;
			if (string.IsNullOrEmpty(sourceDir))
				throw new InvalidOperationException("Cannot determine source directory");

			// Copy executable
			string targetExePath = Path.Combine(targetDirectory, Path.GetFileName(sourceExePath));
			CopyFileWithOverwrite(sourceExePath, targetExePath);
			Debug.WriteLine($"[InstallationManager] Copied executable: {Path.GetFileName(sourceExePath)}");

			// Copy main.ico to Common folder
			CopyIconToCommonFolder(sourceDir, targetDirectory);

			// Copy dependencies
			string[] dependencyPatterns = new[]
			{
				"*.dll",
				"*.runtimeconfig.json",
				"*.deps.json"
			};

			foreach (var pattern in dependencyPatterns)
			{
				var files = Directory.GetFiles(sourceDir, pattern, SearchOption.TopDirectoryOnly);
				foreach (var file in files)
				{
					string fileName = Path.GetFileName(file);
					string destFile = Path.Combine(targetDirectory, fileName);

					try
					{
						CopyFileWithOverwrite(file, destFile);
						Debug.WriteLine($"[InstallationManager] Copied dependency: {fileName}");
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"[InstallationManager] Warning: Could not copy {fileName}: {ex.Message}");
					}
				}
			}
		}

		/// <summary>
		/// Copies the main.ico to the Common folder for use by the application.
		/// </summary>
		/// <param name="sourceDir">Source directory containing the executable.</param>
		/// <param name="targetDirectory">Target Common directory.</param>
		private static void CopyIconToCommonFolder(string sourceDir, string targetDirectory)
		{
			try
			{
				// Look for main.ico next to the executable (already in output folder via .csproj)
				string iconSourcePath = Path.Combine(sourceDir, "main.ico");

				if (!File.Exists(iconSourcePath))
				{
					Debug.WriteLine($"[InstallationManager] Icon not found at: {iconSourcePath}");
					return;
				}

				string iconDestPath = Path.Combine(targetDirectory, "main.ico");
				File.Copy(iconSourcePath, iconDestPath, overwrite: true);

				Debug.WriteLine($"[InstallationManager] Copied icon to Common folder: {iconDestPath}");
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[InstallationManager] Warning: Could not copy icon to Common folder: {ex.Message}");
			}
		}

		/// <summary>
		/// Copies a file with retry logic for locked files.
		/// </summary>
		/// <param name="source">Source file path.</param>
		/// <param name="destination">Destination file path.</param>
		private static void CopyFileWithOverwrite(string source, string destination)
		{
			// If destination exists and is in use, try to delete it
			if (File.Exists(destination))
			{
				try
				{
					File.Delete(destination);
				}
				catch (IOException)
				{
					// File might be in use, try copy with overwrite anyway
					Debug.WriteLine($"[InstallationManager] Warning: Could not delete {Path.GetFileName(destination)}, attempting overwrite");
				}
			}

			File.Copy(source, destination, overwrite: true);
		}

		/// <summary>
		/// Gets the current executable path.
		/// </summary>
		/// <returns>Full path to the current executable.</returns>
		private static string GetCurrentExecutablePath()
		{
			string exePath = Assembly.GetExecutingAssembly().Location;

			// For .NET 6+, the location returns the DLL path, we need the executable
			if (exePath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
			{
				exePath = Path.ChangeExtension(exePath, ".exe");
			}

			// Verify the executable exists
			if (!File.Exists(exePath))
			{
				// Try using Process.GetCurrentProcess() as fallback
				exePath = Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;
			}

			return exePath;
		}
	}
}
