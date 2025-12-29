using System;

namespace Virtual_Desktop_Manager.Core.Models
{
    /// <summary>
    /// Contains path constants for the Virtual Desktop Manager application.
    /// </summary>
    public class Paths
    {
		/// <summary>
		/// User profile path (e.g., C:\Users\Username).
		/// </summary>
		public readonly string UserProfile;

		/// <summary>
		/// Default Windows Desktop folder (%UserProfile%\Desktop).
		/// </summary>
		public readonly string DefaultDesktop;

		/// <summary>
		/// Root folder for all virtual desktop data (%UserProfile%\Desktops_VDM).
		/// </summary>
		public readonly string Root;

		/// <summary>
		/// Bin folder for deleted/moved desktop items (%UserProfile%\Desktops_VDM\Bin).
		/// </summary>
		public readonly string Bin;

		/// <summary>
		/// Path to the application Common folder (%UserProfile%\Desktops_VDM\Common).
		/// </summary>
		public readonly string Common;

		/// <summary>
		/// Folder for icon layout files (%UserProfile%\Desktops_VDM\Layouts).
		/// </summary>
		public readonly string Layouts;

		/// <summary>
		/// Gets the full path to the executable in the Common folder (%UserProfile%\Desktops_VDM\Common\VirtualDesktopManager.exe)
		/// </summary>
		public readonly string Exe;

		/// <summary>
		/// Gets the full path to the main icon in the Common folder (%UserProfile%\Desktops_VDM\Common\main.ico)
		/// </summary>
		public readonly string Icon;

		/// <summary>
		/// Initializes a new instance of the <see cref="Paths"/> class and sets up all application paths.
		/// </summary>
		public Paths()
		{
			UserProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
			DefaultDesktop = Path.Combine(UserProfile, "Desktop");
			Root = Path.Combine(UserProfile, "Desktops_VDM");
			Bin = Path.Combine(Root, "Bin");
			Common = Path.Combine(Root, "Common");
			Layouts = Path.Combine(Root, "Layouts");
			Exe = Path.Combine(Common, "Virtual-Desktop-Manager.exe");
			Icon = Path.Combine(Common, "main.ico");
		}
    }
}
