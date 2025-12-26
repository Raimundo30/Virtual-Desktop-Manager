using System;
using System.Collections.Generic;

namespace Virtual_Desktop_Manager.Core.Models
{
	/// <summary>
	/// Represents the position and metadata of a desktop icon.
	/// </summary>
	public class IconPosition
	{
		/// <summary>
		/// Gets or sets the name of the icon/file.
		/// </summary>
		public string Name { get; set; } = string.Empty;

		/// <summary>
		/// Gets or sets the X coordinate relative to the screen.
		/// </summary>
		public int X { get; set; }

		/// <summary>
		/// Gets or sets the Y coordinate relative to the screen.
		/// </summary>
		public int Y { get; set; }

		/// <summary>
		/// Gets or sets the screen index this icon belongs to.
		/// </summary>
		public int ScreenIndex { get; set; }

		/// <summary>
		/// Gets or sets whether this icon is on the primary screen.
		/// </summary>
		public bool IsPrimaryScreen { get; set; }
	}

	/// <summary>
	/// Represents the complete layout of desktop icons across all screens.
	/// </summary>
	public class DesktopLayout
	{
		/// <summary>
		/// Gets or sets the list of icon positions.
		/// </summary>
		public List<IconPosition> Icons { get; set; } = new();

		/// <summary>
		/// Gets or sets the screen configuration at the time of capture.
		/// </summary>
		public ScreenConfiguration ScreenConfig { get; set; } = new();

		/// <summary>
		/// Gets or sets the timestamp when the layout was saved.
		/// </summary>
		public DateTime SavedAt { get; set; }
	}

	/// <summary>
	/// Represents the screen configuration at a point in time.
	/// </summary>
	public class ScreenConfiguration
	{
		/// <summary>
		/// Gets or sets the list of screens and their properties.
		/// </summary>
		public List<ScreenInfo> Screens { get; set; } = new();

		/// <summary>
		/// Gets or sets the index of the primary screen.
		/// </summary>
		public int PrimaryScreenIndex { get; set; }
	}

	/// <summary>
	/// Represents information about a single screen.
	/// </summary>
	public class ScreenInfo
	{
		/// <summary>
		/// Gets or sets the screen index.
		/// </summary>
		public int Index { get; set; }

		/// <summary>
		/// Gets or sets whether this is the primary screen.
		/// </summary>
		public bool IsPrimary { get; set; }

		/// <summary>
		/// Gets or sets the screen width in pixels.
		/// </summary>
		public int Width { get; set; }

		/// <summary>
		/// Gets or sets the screen height in pixels.
		/// </summary>
		public int Height { get; set; }

		/// <summary>
		/// Gets or sets the X position of the screen in the virtual desktop.
		/// </summary>
		public int Left { get; set; }

		/// <summary>
		/// Gets or sets the Y position of the screen in the virtual desktop.
		/// </summary>
		public int Top { get; set; }
	}
}