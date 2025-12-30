using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
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

				// Ensure there is no other instance running from Common folder
				CloseRunningInstances(paths);

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
		/// Closes any running instances of the application from the Common folder.
		/// </summary>
		/// <param name="paths">Paths configuration.</param>
		private static void CloseRunningInstances(Paths paths)
		{
			try
			{
				var currentProcess = Process.GetCurrentProcess();
				string currentProcessName = currentProcess.ProcessName;
				string commonExePath = Path.GetFullPath(paths.Exe);

				Debug.WriteLine($"[InstallationManager] Checking for running instances of {currentProcessName}...");

				// Get all processes with the same name
				var runningProcesses = Process.GetProcessesByName(currentProcessName)
					.Where(p => p.Id != currentProcess.Id) // Exclude current process
					.ToList();

				if (runningProcesses.Count == 0)
				{
					Debug.WriteLine("[InstallationManager] No other instances found");
					return;
				}

				Debug.WriteLine($"[InstallationManager] Found {runningProcesses.Count} other instance(s)");

				foreach (var process in runningProcesses)
				{
					try
					{
						string processPath = process.MainModule?.FileName ?? string.Empty;

						if (string.IsNullOrEmpty(processPath))
						{
							Debug.WriteLine($"[InstallationManager] Could not determine path for process ID {process.Id}");
							continue;
						}

						string normalizedProcessPath = Path.GetFullPath(processPath).TrimEnd(Path.DirectorySeparatorChar);

						// Check if this process is running from the Common folder
						if (string.Equals(normalizedProcessPath, commonExePath, StringComparison.OrdinalIgnoreCase))
						{
							Debug.WriteLine($"[InstallationManager] Closing instance running from Common folder (PID: {process.Id})");

							// Attempt graceful shutdown first
							if (!process.CloseMainWindow())
							{
								Debug.WriteLine($"[InstallationManager] CloseMainWindow failed for PID {process.Id}, waiting for exit...");
							}

							// Wait up to 5 seconds for graceful shutdown
							if (!process.WaitForExit(5000))
							{
								Debug.WriteLine($"[InstallationManager] Process {process.Id} did not exit gracefully, forcing termination");
								process.Kill();
								process.WaitForExit(2000); // Wait a bit after kill
							}

							Debug.WriteLine($"[InstallationManager] Process {process.Id} closed successfully");
						}
						else
						{
							Debug.WriteLine($"[InstallationManager] Ignoring instance not running from Common folder: {processPath}");
						}
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"[InstallationManager] Error closing process {process.Id}: {ex.Message}");
					}
					finally
					{
						process.Dispose();
					}
				}

				// Give the system a moment to release file handles
				Thread.Sleep(500);

				Debug.WriteLine("[InstallationManager] Finished closing running instances");
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[InstallationManager] Error in CloseRunningInstances: {ex.Message}");
				// Non-critical error, continue with installation
			}
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
