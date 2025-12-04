using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Virtual_Desktop_Manager.Core.Helpers
{
	/// <summary>
	/// Provides COM interop interfaces and GUIDs for accessing Windows Virtual Desktop APIs.
	/// Contains interface definitions and CLSIDs required to interact with the undocumented
	/// Virtual Desktop Manager functionality in Windows 10 and Windows 11.
	/// </summary>
	/// <remarks>
	/// Based on Markus Scholtes' implementation. The interface GUIDs vary across Windows versions,
	/// so multiple candidates are provided for compatibility.
	/// </remarks>
	internal static class VirtualDesktopInterop
	{
		/// <summary>
		/// CLSID for the ImmersiveShell COM object used to access Virtual Desktop services.
		/// </summary>
		public static readonly Guid CLSID_ImmersiveShell = new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239");

		/// <summary>
		/// CLSID for the VirtualDesktopManagerInternal service.
		/// </summary>
		public static readonly Guid CLSID_VirtualDesktopManagerInternal = new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");

		/// <summary>
		/// CLSID for the public VirtualDesktopManager API.
		/// </summary>
		public static readonly Guid CLSID_VirtualDesktopManager = new Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A");

		/// <summary>
		/// Array of known IID candidates for the IVirtualDesktopManagerInternal interface across different Windows versions.
		/// The correct IID must be discovered at runtime by trying each candidate.
		/// </summary>
		public static readonly Guid[] IID_ManagerInternal_Candidates = new[]
		{
			new Guid("53F5CA0B-158F-4124-900C-057158060B27"), // Windows 11 24H2
			new Guid("AF8DA486-95BB-4460-B3B7-6E7A6B2962B5"), // Windows 11 22H2
			new Guid("094AFE11-44F2-4BA0-976F-29A97E263EE0"), // Windows 10 21H2  
			new Guid("F31574D6-B682-4CDC-BD56-1827860ABEC6"), // Windows 10 2004-20H2
			new Guid("B2F925B9-5A0F-4D2E-9F4D-2B1507593C10")  // Windows 10 1903-1909
		};

		/// <summary>
		/// Gets or sets the working IID for IVirtualDesktopManagerInternal that was successfully discovered at runtime.
		/// Caching this value avoids repeated discovery attempts.
		/// </summary>
		public static Guid? WorkingManagerIID { get; set; }

		/// <summary>
		/// COM interface for querying services from the ImmersiveShell.
		/// Used to obtain the IVirtualDesktopManagerInternal service.
		/// </summary>
		[ComImport]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		[Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
		public interface IServiceProvider10
		{
			[return: MarshalAs(UnmanagedType.IUnknown)]
			object QueryService(ref Guid service, ref Guid riid);
		}

		/// <summary>
		/// Internal COM interface for managing Virtual Desktops.
		/// Provides methods to create, switch, move, and remove virtual desktops.
		/// </summary>
		/// <remarks>
		/// This is an undocumented Windows API. The interface GUID varies by Windows version.
		/// </remarks>
		[ComImport]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		[Guid("53F5CA0B-158F-4124-900C-057158060B27")]
		public interface IVirtualDesktopManagerInternal
		{
			int GetCount();
			void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
			bool CanViewMoveDesktops(IApplicationView view);
			IVirtualDesktop GetCurrentDesktop();
			void GetDesktops(out IObjectArray desktops);
			int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
			void SwitchDesktop(IVirtualDesktop desktop);
			void SwitchDesktopAndMoveForegroundView(IVirtualDesktop desktop);
			IVirtualDesktop CreateDesktop();
			void MoveDesktop(IVirtualDesktop desktop, int nIndex);
			void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
			IVirtualDesktop FindDesktop(ref Guid desktopid);
		}

		/// <summary>
		/// COM interface representing a single virtual desktop.
		/// Provides methods to query desktop properties and visibility.
		/// </summary>
		[ComImport]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		[Guid("3F07F4BE-B107-441A-AF0F-39D82529072C")]
		public interface IVirtualDesktop
		{
			bool IsViewVisible(IApplicationView view);
			Guid GetId();
			[return: MarshalAs(UnmanagedType.HString)]
			string GetName();
		}

		/// <summary>
		/// COM interface representing an application view (window).
		/// This interface is required by other Virtual Desktop APIs but is not directly used for desktop ID retrieval.
		/// </summary>
		[ComImport]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		[Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
		public interface IApplicationView { }

		/// <summary>
		/// COM interface representing an array of objects.
		/// Used to enumerate collections of virtual desktops.
		/// </summary>
		[ComImport]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		[Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
		public interface IObjectArray
		{
			void GetCount(out int count);
			void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object obj);
		}
	}
}