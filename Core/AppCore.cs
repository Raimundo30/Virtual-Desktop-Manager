using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
		public readonly VirtualDesktopMonitor _monitor;
		private readonly DesktopFolderManager _folderManager;
		private readonly IconLayoutManager _iconManager;
		private readonly CleanupManager _cleanupManager;

		/// <summary>
		/// Event fired when a desktop switch is completed.
		/// Provides the new active desktop ID.
		/// </summary>
		public event Action<Guid>? DesktopSwitched;

		/// <summary>
		/// Occurs when an error is encountered during the execution of the component.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Initializes a new instance of the <see cref="AppCore"/> class and sets up all core services.
		/// </summary>
		public AppCore()
		{
			string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
			string rootFolder = Path.Combine(userProfile, "Desktops_VDM");
			string layoutFolder = Path.Combine(rootFolder, "Layouts");

			_monitor = new VirtualDesktopMonitor();
			_folderManager = new DesktopFolderManager(rootFolder, userProfile);
			_iconManager = new IconLayoutManager(layoutFolder, _folderManager);
			_cleanupManager = new CleanupManager(rootFolder, layoutFolder);

			// Subscribe to desktop change events
			_monitor.DesktopChanged += OnDesktopChanged;
			_monitor.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_iconManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_folderManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
			_cleanupManager.ErrorOccurred += (ex) => ErrorOccurred?.Invoke(ex);
		}

		/// <summary>
		/// Starts monitoring virtual desktops.
		/// </summary>
		public void Start()
		{
			_monitor.Start();
			OnDesktopChanged(_monitor.LastDesktopId);
		}

		/// <summary>
		/// Stops monitoring virtual desktops.
		/// </summary>
		public void Stop()
		{
			_monitor.Stop();

			// Optionally, switch back to the default desktop folder on stop
			OnDesktopChanged(Guid.Empty);
		}

		/// <summary>
		/// Handles a virtual desktop change event from VirtualDesktopMonitor.
		/// </summary>
		/// <param name="desktopId">The GUID of the new active desktop.</param>
		private void OnDesktopChanged(Guid desktopId)
		{
			_monitor._isPaused = true;

			try
			{
				// 1. Save icon layout for the old desktop
				_iconManager.SaveLayout();

				// 2. Ensure folder exists for the new desktop
				string newFolder = _folderManager.GetFolderForDesktop(desktopId);

				// 3. Switch Windows Desktop path to the new folder
				_folderManager.SwitchDesktopPath(newFolder);

				// 4. Load icon layout for the new desktop
				_iconManager.LoadLayout();

				// 5. Refresh the desktop environment
				_folderManager.RefreshDesktop();

				// 6. Run cleanup for unused folders and layouts
				List<Guid> activeDesktops = _monitor.ListDesktopIds();
				_cleanupManager.CleanupAll(activeDesktops);

				// 7. Notify UI or other listeners
				DesktopSwitched?.Invoke(desktopId);
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}

			_monitor._isPaused = false;
		}

		/// <summary>
		/// Releases all resources used by the current instance.
		/// </summary>
		public void Dispose()
		{
			Stop();
			_monitor.Dispose();
		}
	}
}