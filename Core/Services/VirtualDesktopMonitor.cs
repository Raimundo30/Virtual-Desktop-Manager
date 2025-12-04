using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using Virtual_Desktop_Manager.Core.Helpers;
using static Virtual_Desktop_Manager.Core.Helpers.VirtualDesktopInterop;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Monitors virtual desktop changes by periodically checking the current desktop ID.
	/// Raises an event whenever a change is detected.
	/// </summary>
	public class VirtualDesktopMonitor : IDisposable
	{
		/// <summary>
		/// Occurs when an error is encountered during operation.
		/// </summary>
		public event Action<Exception>? ErrorOccurred;

		/// <summary>
		/// Occurs when a change to the current desktop is detected.
		/// </summary>
		public event Action<Guid>? DesktopChanged;

		/// <summary>
		/// Internal flag to prevent re-entrant desktop switching during desktop change operations.
		/// </summary>
		public bool _isPaused = false;

		/// <summary>
		/// Gets the ID of the last known virtual desktop.
		/// </summary>
		public Guid LastDesktopId { get; private set; }

		/// <summary>
		/// Gets or sets the polling interval in milliseconds for checking desktop changes.
		/// Default is 100ms.
		/// </summary>
		public int PollingIntervalMs { get; set; } = 100;

		// Timer for periodic polling
		private readonly System.Timers.Timer _timer;

		// Internal COM interface for virtual desktop management
		private VirtualDesktopInterop.IVirtualDesktopManagerInternal? _vdmInternal;

		/// <summary>
		/// Initializes a new instance of the <see cref="VirtualDesktopMonitor"/> class.
		/// </summary>
		public VirtualDesktopMonitor()
		{
			_timer = new System.Timers.Timer(PollingIntervalMs);
			_timer.Elapsed += TimerElapsed;
			_timer.AutoReset = true;

			// Initialize COM interface
			InitializeCOM();

			// Initialize LastDesktopId to the current desktop
			LastDesktopId = GetCurrentDesktopId();
		}

		/// <summary>
		/// Initializes the COM interface required for virtual desktop management operations.
		/// </summary>
		/// <remarks>If initialization fails, the method triggers the <see cref="ErrorOccurred"/> event with the
		/// encountered exception. This method should be called before performing any actions that depend on the virtual
		/// desktop manager COM interface.</remarks>
		private void InitializeCOM()
		{
			try
			{
				// Step 1: Create ImmersiveShell (this CLSID is stable across Windows versions)
				var shellType = Type.GetTypeFromCLSID(VirtualDesktopInterop.CLSID_ImmersiveShell);
				if (shellType == null)
				{
					ErrorOccurred?.Invoke(new COMException("Failed to get ImmersiveShell COM type"));
					return;
				}

				var shell = Activator.CreateInstance(shellType) as VirtualDesktopInterop.IServiceProvider10;

				// Step 2: Try to get IVirtualDesktopManagerInternal via QueryService with different IIDs
				bool success = false;

				// Try cached IID first if available
				if (shell != null && VirtualDesktopInterop.WorkingManagerIID.HasValue)
				{
					success = TryQueryService(shell, VirtualDesktopInterop.WorkingManagerIID.Value);
				}

				// If cached attempt failed, try all known IIDs
				if (!success && shell != null)
				{
					foreach (var iid in VirtualDesktopInterop.IID_ManagerInternal_Candidates)
					{
						if (TryQueryService(shell, iid))
						{
							VirtualDesktopInterop.WorkingManagerIID = iid;
							success = true;
							break;
						}
					}
				}

				// Cleanup shell COM object
				if (shell != null && Marshal.IsComObject(shell))
					Marshal.ReleaseComObject(shell);

				if (!success)
				{
					ErrorOccurred?.Invoke(new COMException("Failed to initialize Virtual Desktop COM interface with any known IID"));
				}
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
		}

		/// <summary>
		/// Attempts to query IVirtualDesktopManagerInternal service with a specific IID.
		/// </summary>
		/// <param name="shell">The IServiceProvider10 instance to query.</param>
		/// <param name="iid">The interface identifier to attempt.</param>
		/// <returns>True if successful, false otherwise.</returns>
		private bool TryQueryService(VirtualDesktopInterop.IServiceProvider10 shell, Guid iid)
		{
			try
			{
				var clsid = VirtualDesktopInterop.CLSID_VirtualDesktopManagerInternal;
				var service = shell.QueryService(ref clsid, ref iid);

				if (service is VirtualDesktopInterop.IVirtualDesktopManagerInternal vdmInternal)
				{
					_vdmInternal = vdmInternal;
					return true;
				}

				return false;
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
				return false;
			}
		}

		/// <summary>
		/// Gets the current virtual desktop ID.
		/// If retrieval fails, returns Guid.Empty and triggers the <see cref="ErrorOccurred"/> event.
		/// </summary>
		/// <returns>Guid representing the current virtual desktop.</returns>
		public Guid GetCurrentDesktopId()
		{
			if (_vdmInternal == null)
				return Guid.Empty;

			try
			{
				// Retrieve current desktop and its ID
				var desktop = _vdmInternal.GetCurrentDesktop();
				if (desktop == null)
					return Guid.Empty;

				Guid id = desktop.GetId();

				// Release COM object immediately
				if (Marshal.IsComObject(desktop))
					Marshal.ReleaseComObject(desktop);

				return id;
			}
			catch (COMException ex)
			{
				ErrorOccurred?.Invoke(ex);
				return Guid.Empty;
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
				return Guid.Empty;
			}
		}

		/// <summary>
		/// Starts monitoring virtual desktops.
		/// </summary>
		public void Start()
		{
			_timer.Interval = PollingIntervalMs;
			_timer.Start();
		}

		/// <summary>
		/// Stops monitoring virtual desktops.
		/// </summary>
		public void Stop()
		{
			_timer.Stop();
		}

		/// <summary>
		/// Timer callback: checks for desktop changes.
		/// </summary>
		private void TimerElapsed(object? sender, ElapsedEventArgs e)
		{
			// Prevent re-entrant calls
			if (_isPaused == true)
				return;

			Guid currentId = GetCurrentDesktopId();

			if (currentId != LastDesktopId)
			{
				// Update LastDesktopId
				LastDesktopId = currentId;

				// Raise event with old and new desktop IDs
				DesktopChanged?.Invoke(currentId);

			}
		}

		/// <summary>
		/// Gets all virtual desktop IDs currently available in Windows.
		/// </summary>
		/// <returns>List of all desktop GUIDs. Returns empty list on error.</returns>
		public List<Guid> ListDesktopIds()
		{
			var desktopIds = new List<Guid>();

			if (_vdmInternal == null)
				return desktopIds;

			try
			{
				_vdmInternal.GetDesktops(out IObjectArray desktops);
				desktops.GetCount(out int count);

				var iid = typeof(VirtualDesktopInterop.IVirtualDesktop).GUID;

				for (int i = 0; i < count; i++)
				{
					desktops.GetAt(i, ref iid, out object obj);
					if (obj is VirtualDesktopInterop.IVirtualDesktop desktop)
					{
						desktopIds.Add(desktop.GetId());

						// Release COM object immediately
						if (Marshal.IsComObject(desktop))
							Marshal.ReleaseComObject(desktop);
					}
				}

				// Release the IObjectArray
				if (Marshal.IsComObject(desktops))
					Marshal.ReleaseComObject(desktops);
			}
			catch (COMException ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}

			return desktopIds;
		}

		/// <summary>
		/// Releases all resources used by the current instance.
		/// </summary>
		public void Dispose()
		{
			_timer.Stop();
			_timer.Dispose();

			// Release COM object
			if (_vdmInternal != null && Marshal.IsComObject(_vdmInternal))
			{
				Marshal.ReleaseComObject(_vdmInternal);
				_vdmInternal = null;
			}
		}
	}
}