using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using Virtual_Desktop_Manager.Core.Models;
using Virtual_Desktop_Manager.Core.Services;

namespace Virtual_Desktop_Manager.Core
{
	/// <summary>
	/// Coordinates the Desktop Change workflow.
	/// - Monitors virtual desktop changes
	/// - Saves and loads icon layouts
	/// - Switches Windows Desktop path
	/// - Refreshes desktop
	/// - Cleans up unused folders and layouts
	/// </summary>
	public class AppCore : IDisposable
	{
		// Core services
		private readonly Paths _paths;
		private readonly VirtualDesktopService _desktopService;
		private readonly DesktopFolderManager _folderManager;
		private readonly IconLayoutManager _iconManager;
		private readonly CleanupManager _cleanupManager;
		private readonly StartupManager _startupManager;
		private readonly UninstallManager _uninstallManager;

		/// <summary>
		/// Internal flag to prevent re-entrant desktop switching during desktop change operations.
		/// 0 = not processing, 1 = processing
		/// </summary>
		private int _isProcessing = 0;

		/// <summary>
		/// Event fired when a desktop switch is completed.
		/// Provides the new active desktop ID.
		/// </summary>
		public event Action<Guid>? DesktopSwitched;

		/// <summary>
		/// Occurs when an informational message is generated.
		/// </summary>
		public event Action<string, string, string>? Message;

		/// <summary>
		/// Occurs when an error is encountered during the execution of the component.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Gets the virtual desktop service for managing desktop operations and monitoring.
		/// </summary>
		public VirtualDesktopService DesktopService => _desktopService;

		/// <summary>
		/// Gets the desktop folder manager for managing virtual desktop folders.
		/// </summary>
		public DesktopFolderManager FolderManager => _folderManager;

		/// <summary>
		/// Gets the paths configuration for the application.
		/// </summary>
		public Paths Paths => _paths;

		/// <summary>
		/// Gets the cleanup manager for removing unused desktop folders and layouts.
		/// </summary>
		public CleanupManager CleanupManager => _cleanupManager;

		/// <summary>
		/// Gets the startup manager for Windows startup configuration.
		/// </summary>
		public StartupManager StartupManager => _startupManager;

		/// <summary>
		/// Gets the uninstall manager for application uninstallation.
		/// </summary>
		public UninstallManager UninstallManager => _uninstallManager;

		/// <summary>
		/// Initializes a new instance of the <see cref="AppCore"/> class and sets up all core services.
		/// </summary>
		public AppCore()
		{
			_paths = new Paths();
			_desktopService = new VirtualDesktopService();
			_folderManager = new DesktopFolderManager(_paths);
			_iconManager = new IconLayoutManager(_paths.Layouts, _folderManager);
			_cleanupManager = new CleanupManager(_paths);
			_startupManager = new StartupManager(_paths);
			_uninstallManager = new UninstallManager(this);

			// Subscribe to desktop change events
			_desktopService.DesktopChanged += OnDesktopChanged;

			// Forward errors from services
			_desktopService.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_iconManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_folderManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_cleanupManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_uninstallManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
		}

		/// <summary>
		/// Starts monitoring virtual desktops.
		/// </summary>
		public void Start()
		{
			_desktopService.Start();

			OnDesktopChanged(_desktopService.LastDesktopId);
		}

		/// <summary>
		/// Stops monitoring virtual desktops.
		/// </summary>
		public void Stop()
		{
			_desktopService.Stop();

			// Optionally, switch back to the default desktop folder on stop
			OnDesktopChanged(Guid.Empty);
		}

		/// <summary>
		/// Handles a virtual desktop change event from VirtualDesktopService.
		/// </summary>
		/// <param name="desktopId">The GUID of the new active desktop.</param>
		private void OnDesktopChanged(Guid desktopId)
		{
			// Atomically check if already processing and set flag if not
			// Returns the OLD value: if it was 1 (processing), another operation is running
			if (Interlocked.CompareExchange(ref _isProcessing, 1, 0) == 1)
			{
				Debug.WriteLine($"[AppCore] Desktop switch already in progress, ignoring event for {desktopId}");
				return;
			}

			try
			{
				Debug.WriteLine($"[AppCore] Processing desktop switch to {desktopId}");

				// 1. Save icon layout for the old desktop
				_iconManager.SaveLayout();

				// 2. Ensure folder exists for the new desktop
				string newFolder = _folderManager.GetFolderForDesktop(desktopId);

				// 3. Switch Windows Desktop path to the new folder
				_folderManager.SwitchDesktopPath(newFolder);

				// 4. Refresh the desktop environment
				_folderManager.RefreshDesktop();

				// 4.1 Wait a moment to ensure Explorer has updated (minimum 200ms)
				WaitForDesktopRefresh(newFolder);

				// 5. Load icon layout for the new desktop
				_iconManager.LoadLayout();

				// 6. Run cleanup for unused folders and layouts
				List<Guid> activeDesktops = _desktopService.ListDesktopIds();
				_cleanupManager.CleanupAll(activeDesktops);

				// 7. Notify UI or other listeners
				DesktopSwitched?.Invoke(desktopId);
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
			finally
			{
				// Atomically reset the flag to allow next operation
				Interlocked.Exchange(ref _isProcessing, 0);

				// Update last desktop ID
				_desktopService.LastDesktopId = desktopId;

				Debug.WriteLine($"[AppCore] Desktop switch completed, flag reset");
			}
		}

		/// <summary>
		/// Waits for the desktop to refresh by verifying the desktop path has changed.
		/// </summary>
		/// <param name="expectedFolder">The folder path that should be active on the desktop.</param>
		private void WaitForDesktopRefresh(string expectedFolder)
		{
			const int delayMs = 100;
			const int maxAttempts = 20; // Maximum 2 seconds (20 * 100ms)

			for (int i = 0; i < maxAttempts; i++)
			{
				// Verify the desktop path has actually changed in the registry/system
				string currentPath = _folderManager.GetCurrentDesktopPath();
				string normalizedCurrent = Path.GetFullPath(currentPath).TrimEnd(Path.DirectorySeparatorChar);
				string normalizedExpected = Path.GetFullPath(expectedFolder).TrimEnd(Path.DirectorySeparatorChar);

				if (string.Equals(normalizedCurrent, normalizedExpected, StringComparison.OrdinalIgnoreCase))
				{
					Debug.WriteLine($"[AppCore] Desktop path verified after {i * delayMs}ms");
					// Give Explorer a bit more time to actually render icons
					Thread.Sleep(200);
					return;
				}

				Thread.Sleep(delayMs);
			}

			Debug.WriteLine($"[AppCore] Warning: Desktop path verification timeout after {maxAttempts * delayMs}ms");
		}

		/// <summary>
		/// Releases all resources used by the current instance.
		/// </summary>
		public void Dispose()
		{
			Stop();
			_desktopService.Dispose();
		}
	}
}