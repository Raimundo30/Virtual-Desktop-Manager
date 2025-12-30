using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Virtual_Desktop_Manager.Core.Events
{
	/// <summary>
	/// Provides data for notification events.
	/// </summary>
	public class NotificationEventArgs : EventArgs
	{
		/// <summary>
		/// Gets the severity level of the notification.
		/// </summary>
		public NotificationSeverity Severity { get; init; }

		/// <summary>
		/// Gets the source component that generated the notification.
		/// </summary>
		public string Source { get; init; }

		/// <summary>
		/// Gets the notification message content.
		/// </summary>
		public string Message { get; init; }

		/// <summary>
		/// Gets the duration the notification should be displayed.
		/// </summary>
		public NotificationDuration Duration { get; init; }

		/// <summary>
		/// Gets the exception associated with this notification (if any).
		/// </summary>
		public Exception? Exception { get; init; }

		/// <summary>
		/// Initializes a new instance of the <see cref="NotificationEventArgs"/> class.
		/// </summary>
		public NotificationEventArgs(
			NotificationSeverity severity,
			string source,
			string message,
			NotificationDuration duration = NotificationDuration.Short,
			Exception? exception = null)
		{
			Severity = severity;
			Source = source;
			Message = message;
			Duration = duration;
			Exception = exception;
		}

		/// <summary>
		/// Gets the icon/emoji prefix for the notification based on severity.
		/// </summary>
		public string Icon => Severity switch
		{
			NotificationSeverity.Info => "ℹ️",
			NotificationSeverity.Success => "✅",
			NotificationSeverity.Warning => "⚠️",
			NotificationSeverity.Error => "❌",
			_ => ""
		};
	}

	/// <summary>
	/// Specifies the severity level of a notification.
	/// </summary>
	public enum NotificationSeverity
	{
		/// <summary>
		/// Informational message.
		/// </summary>
		Info,

		/// <summary>
		/// Success message.
		/// </summary>
		Success,

		/// <summary>
		/// Warning message.
		/// </summary>
		Warning,

		/// <summary>
		/// Error message.
		/// </summary>
		Error
	}

	/// <summary>
	/// Specifies the duration for displaying notifications.
	/// </summary>
	public enum NotificationDuration
	{
		/// <summary>
		/// Short duration (system default, typically ~7 seconds).
		/// </summary>
		Short,

		/// <summary>
		/// Long duration (remains visible until dismissed).
		/// </summary>
		Long
	}
}
