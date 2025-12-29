using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Virtual_Desktop_Manager.UI.Notifications
{
	/// <summary>
	/// Manages application notifications through the Windows notification system.
	/// Notifications appear in the Windows Action Center and respect system notification settings.
	/// </summary>
	public class NotificationManager
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="NotificationManager"/> class.
		/// </summary>
		public NotificationManager()
		{
			// Register the app for toast notifications
			ToastNotificationManagerCompat.OnActivated += OnToastActivated;
		}

		/// <summary>
		/// Shows a Windows toast notification.
		/// </summary>
		/// <param name="title">The notification title.</param>
		/// <param name="message">The notification message.</param>
		/// <param name="duration">The duration the notification should be displayed.</param>
		/// <param name="iconPrefix">Optional emoji or text prefix for the title.</param>
		public void ShowToast(string title, string message, ToastDuration duration, string iconPrefix = "")
		{
			try
			{
				new ToastContentBuilder()
					.AddText(iconPrefix + title)
					.AddText(message)
					.SetToastDuration(duration)
					.Show();
			}
			catch (Exception ex)
			{
				// Fallback to Debug output logging if toast fails
				Debug.WriteLine($"[Notification] {title}: {message}");
				Debug.WriteLine($"[Notification Error] {ex.Message}");
			}
		}

		/// <summary>
		/// Handles toast notification activation (when user clicks the notification).
		/// </summary>
		private void OnToastActivated(ToastNotificationActivatedEventArgsCompat e)
		{
			// Handle notification clicks here if needed
			// For example, could open a settings window
		}

		/// <summary>
		/// Cleans up notification resources.
		/// </summary>
		public void Dispose()
		{
			ToastNotificationManagerCompat.OnActivated -= OnToastActivated;
			ToastNotificationManagerCompat.Uninstall();
		}

		/// <summary>
		/// Toast duration options.
		/// </summary>
		public enum ToastDuration
		{
			Short,  // Default duration (system-controlled, typically ~7 seconds)
			Long    // Extended duration using Reminder scenario (remains visible until dismissed)
		}
	}

	/// <summary>
	/// Extends ToastContentBuilder with custom duration handling for notifications.
	/// </summary>
	internal static class ToastExtensions
	{
		/// <summary>
		/// Sets the toast duration based on the specified duration type.
		/// </summary>
		/// <param name="builder">The ToastContentBuilder to configure.</param>
		/// <param name="duration">The duration type to apply.</param>
		/// <returns>The configured ToastContentBuilder for method chaining.</returns>
		public static ToastContentBuilder SetToastDuration(this ToastContentBuilder builder, NotificationManager.ToastDuration duration)
		{
			// Long duration uses Reminder scenario for extended visibility
			if (duration == NotificationManager.ToastDuration.Long)
			{
				builder.SetToastScenario(ToastScenario.Reminder);
			}
			return builder;
		}
	}
}
