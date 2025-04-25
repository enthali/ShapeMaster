using System;
using System.Diagnostics;
using System.Text;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Provides centralized error handling functionality for the ShapeMaster add-in
    /// </summary>
    public class ErrorHandlingService
    {
        // Reference to notification service for displaying errors to the user
        private readonly NotificationService _notificationService;
        
        // Control whether to include detailed stacktrace info in logs
        private readonly bool _includeStackTraceInLogs;
        
        /// <summary>
        /// Initializes a new instance of the ErrorHandlingService class
        /// </summary>
        /// <param name="notificationService">Notification service for displaying errors</param>
        /// <param name="includeStackTraceInLogs">Whether to include stacktrace info in logs (default: true)</param>
        public ErrorHandlingService(NotificationService notificationService, bool includeStackTraceInLogs = true)
        {
            _notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            _includeStackTraceInLogs = includeStackTraceInLogs;
        }
        
        /// <summary>
        /// Handles an exception by logging it and optionally showing a notification to the user
        /// </summary>
        /// <param name="ex">The exception to handle</param>
        /// <param name="userMessage">Optional custom message to show to the user</param>
        /// <param name="showNotification">Whether to show a notification to the user</param>
        /// <param name="operationName">Name of the operation that caused the exception</param>
        public void HandleException(Exception ex, string userMessage = null, bool showNotification = true, string operationName = null)
        {
            if (ex == null) return;
            
            // Log the exception
            LogException(ex, operationName);
            
            // Show notification if requested
            if (showNotification && _notificationService != null)
            {
                string message = userMessage ?? GetUserFriendlyErrorMessage(ex);
                _notificationService.ShowNotification(message, true);
            }
        }
        
        /// <summary>
        /// Executes an action with exception handling
        /// </summary>
        /// <param name="action">The action to execute</param>
        /// <param name="userErrorMessage">Optional custom error message to display on failure</param>
        /// <param name="showNotification">Whether to show a notification on error</param>
        /// <param name="operationName">Name of the operation for logging</param>
        /// <returns>True if the action executed successfully, false otherwise</returns>
        public bool TryExecute(Action action, string userErrorMessage = null, bool showNotification = true, string operationName = null)
        {
            if (action == null) return false;
            
            try
            {
                action();
                return true;
            }
            catch (Exception ex)
            {
                HandleException(ex, userErrorMessage, showNotification, operationName);
                return false;
            }
        }
        
        /// <summary>
        /// Executes a function with exception handling
        /// </summary>
        /// <typeparam name="T">The return type of the function</typeparam>
        /// <param name="func">The function to execute</param>
        /// <param name="defaultValue">Default value to return on error</param>
        /// <param name="userErrorMessage">Optional custom error message to display on failure</param>
        /// <param name="showNotification">Whether to show a notification on error</param>
        /// <param name="operationName">Name of the operation for logging</param>
        /// <returns>The function result or the default value if an exception occurred</returns>
        public T TryExecute<T>(Func<T> func, T defaultValue, string userErrorMessage = null, bool showNotification = true, string operationName = null)
        {
            if (func == null) return defaultValue;
            
            try
            {
                return func();
            }
            catch (Exception ex)
            {
                HandleException(ex, userErrorMessage, showNotification, operationName);
                return defaultValue;
            }
        }
        
        /// <summary>
        /// Logs an exception to the debug output
        /// </summary>
        /// <param name="ex">The exception to log</param>
        /// <param name="operationName">Optional name of the operation</param>
        private void LogException(Exception ex, string operationName)
        {
            if (ex == null) return;
            
            StringBuilder logMessage = new StringBuilder();
            
            logMessage.AppendLine("=== ShapeMaster Error ===");
            
            if (!string.IsNullOrEmpty(operationName))
            {
                logMessage.AppendLine($"Operation: {operationName}");
            }
            
            logMessage.AppendLine($"Exception: {ex.GetType().Name}");
            logMessage.AppendLine($"Message: {ex.Message}");
            
            if (ex.InnerException != null)
            {
                logMessage.AppendLine($"Inner Exception: {ex.InnerException.GetType().Name}");
                logMessage.AppendLine($"Inner Message: {ex.InnerException.Message}");
            }
            
            if (_includeStackTraceInLogs && !string.IsNullOrEmpty(ex.StackTrace))
            {
                logMessage.AppendLine("Stack Trace:");
                logMessage.AppendLine(ex.StackTrace);
            }
            
            logMessage.AppendLine("======================");
            
            Debug.WriteLine(logMessage.ToString());
        }
        
        /// <summary>
        /// Gets a user-friendly error message from an exception
        /// </summary>
        /// <param name="ex">The exception</param>
        /// <returns>A user-friendly error message</returns>
        private string GetUserFriendlyErrorMessage(Exception ex)
        {
            if (ex == null) return "An unknown error occurred.";
            
            // Start with the main exception message
            string message = ex.Message;
            
            // If it's a specific known exception type, we could customize the message
            if (ex is System.IO.FileNotFoundException)
            {
                message = "A required file could not be found: " + ex.Message;
            }
            else if (ex is System.UnauthorizedAccessException)
            {
                message = "Access denied: " + ex.Message;
            }
            else if (ex is System.Runtime.InteropServices.COMException)
            {
                message = "An error occurred when communicating with PowerPoint: " + ex.Message;
            }
            else if (ex is ArgumentException)
            {
                message = "Invalid parameter: " + ex.Message;
            }
            else if (ex is NullReferenceException)
            {
                message = "An operation failed because a required object was not available.";
            }
            
            // Limit the message length for the notification
            const int maxLength = 150;
            if (message.Length > maxLength)
            {
                message = message.Substring(0, maxLength) + "...";
            }
            
            return message;
        }
    }
}