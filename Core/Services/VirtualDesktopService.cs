using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using Virtual_Desktop_Manager.Core.Helpers;
using static Virtual_Desktop_Manager.Core.Helpers.VirtualDesktopInterop;

namespace Virtual_Desktop_Manager.Core.Services
{
	/// <summary>
	/// Provides virtual desktop services including monitoring desktop changes and performing desktop operations.
	/// Monitors changes by periodically checking the current desktop ID and raises events when changes are detected.
	/// Also provides methods to switch desktops and retrieve desktop information.
	/// </summary>
	public class VirtualDesktopService : IDisposable
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
		/// Gets the ID of the last known virtual desktop.
		/// </summary>
		public Guid LastDesktopId { get; set; }

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
		/// Initializes a new instance of the <see cref="VirtualDesktopService"/> class.
		/// </summary>
		public VirtualDesktopService()
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
			Guid currentId = GetCurrentDesktopId();

			if (currentId != LastDesktopId)
			{
				// Raise event with old and new desktop IDs
				DesktopChanged?.Invoke(currentId);

				// LastDesktopId will be updated after desktop switch is fully processed

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
		/// Gets information about all virtual desktops (ID and Name).
		/// </summary>
		/// <returns>Dictionary mapping desktop GUIDs to their names.</returns>
		public Dictionary<Guid, string> GetDesktopInfo()
		{
			var desktopInfo = new Dictionary<Guid, string>();

			if (_vdmInternal == null)
				return desktopInfo;

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
						Guid id = desktop.GetId();
						Debug.WriteLine($"[VirtualDesktopService] Desktop {i}: ID = {id}");

						string name = string.Empty;
						try
						{
							Debug.WriteLine($"[VirtualDesktopService] Calling GetName() for desktop {i}");
							IntPtr namePtr = desktop.GetName();
							Debug.WriteLine($"[VirtualDesktopService] GetName() returned IntPtr: 0x{namePtr:X}");

							if (namePtr != IntPtr.Zero)
							{
								// Use WindowsGetStringRawBuffer instead of PtrToStringUni for HSTRING
								uint length = 0;
								IntPtr buffer = WindowsGetStringRawBuffer(namePtr, out length);

								if (buffer != IntPtr.Zero && length > 0)
								{
									name = Marshal.PtrToStringUni(buffer, (int)length) ?? string.Empty;
									Debug.WriteLine($"[VirtualDesktopService] Converted to string: '{name}' (length: {length})");
								}
								else
								{
									Debug.WriteLine($"[VirtualDesktopService] HSTRING buffer is null or empty");
								}

								// Free the HSTRING memory
								try
								{
									WindowsDeleteString(namePtr);
									Debug.WriteLine($"[VirtualDesktopService] Successfully freed HSTRING");
								}
								catch (Exception freeEx)
								{
									Debug.WriteLine($"[VirtualDesktopService] Failed to free HSTRING: {freeEx.Message}");
								}
							}
							else
							{
								Debug.WriteLine($"[VirtualDesktopService] GetName() returned null pointer");
							}
						}
						catch (Exception ex)
						{
							// If GetName() fails, we'll use the default name below
							Debug.WriteLine($"[VirtualDesktopService] Exception in GetName(): {ex.GetType().Name}: {ex.Message}");
							name = string.Empty;
						}

						// If name is empty, use a default name with index
						if (string.IsNullOrWhiteSpace(name))
						{
							name = $"Desktop {i + 1}";
							Debug.WriteLine($"[VirtualDesktopService] Using default name: '{name}'");
						}
						else
						{
							Debug.WriteLine($"[VirtualDesktopService] Using retrieved name: '{name}'");
						}

						desktopInfo[id] = name;

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
				Debug.WriteLine($"[VirtualDesktopService] COM Exception: {ex.Message}");
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
				Debug.WriteLine($"[VirtualDesktopService] Exception: {ex.Message}");
			}

			return desktopInfo;
		}

		/// <summary>
		/// Switches to a specific virtual desktop by its ID.
		/// </summary>
		/// <param name="desktopId">The ID of the desktop to switch to.</param>
		/// <returns>True if successful, false otherwise.</returns>
		public bool SwitchToDesktop(Guid desktopId)
		{
			if (_vdmInternal == null)
				return false;

			try
			{
				var desktop = _vdmInternal.FindDesktop(ref desktopId);
				if (desktop != null)
				{
					_vdmInternal.SwitchDesktop(desktop);

					// Release COM object
					if (Marshal.IsComObject(desktop))
						Marshal.ReleaseComObject(desktop);

					return true;
				}
			}
			catch (COMException ex)
			{
				ErrorOccurred?.Invoke(ex);
			}
			catch (Exception ex)
			{
				ErrorOccurred?.Invoke(ex);
			}

			return false;
		}

		[DllImport("api-ms-win-core-winrt-string-l1-1-0.dll", CallingConvention = CallingConvention.StdCall)]
		private static extern IntPtr WindowsGetStringRawBuffer(IntPtr hstring, out uint length);

		[DllImport("api-ms-win-core-winrt-string-l1-1-0.dll", CallingConvention = CallingConvention.StdCall)]
		private static extern int WindowsDeleteString(IntPtr hstring);

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