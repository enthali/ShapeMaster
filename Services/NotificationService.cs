using System.Drawing;
using System.Windows.Forms;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing user notifications
    /// </summary>
    public class NotificationService
    {
        // ToolTip instance for showing non-modal notifications
        private readonly ToolTip _notificationTip = new ToolTip();

        /// <summary>
        /// Initializes a new instance of the NotificationService class
        /// </summary>
        public NotificationService()
        {
            // Configure the tooltip for notifications
            _notificationTip.IsBalloon = true;
            _notificationTip.AutoPopDelay = 2500; // 2.5 seconds
            _notificationTip.InitialDelay = 0;
            _notificationTip.ReshowDelay = 0;
            _notificationTip.UseAnimation = true;
            _notificationTip.UseFading = true;
            _notificationTip.ToolTipTitle = "ShapeMaster";
        }

        /// <summary>
        /// Shows a non-modal notification
        /// </summary>
        /// <param name="message">The message to display</param>
        /// <param name="isError">Whether this is an error message</param>
        public void ShowNotification(string message, bool isError = false)
        {
            try
            {
                // For critical errors, still use MessageBox
                if (isError)
                {
                    MessageBox.Show(message, "ShapeMaster", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Get cursor position
                Point cursorPos = Cursor.Position;

                // Create a dummy control to show the tooltip near
                Control dummyControl = new Control();
                dummyControl.Location = new Point(cursorPos.X - 20, cursorPos.Y - 20);
                dummyControl.Size = new Size(1, 1);

                // Add the dummy control to a dummy form temporarily
                Form dummyForm = new Form();
                dummyForm.StartPosition = FormStartPosition.Manual;
                dummyForm.Location = new Point(cursorPos.X - 25, cursorPos.Y - 25);
                dummyForm.Size = new Size(10, 10);
                dummyForm.FormBorderStyle = FormBorderStyle.None;
                dummyForm.ShowInTaskbar = false;
                dummyForm.TopMost = true;
                dummyForm.Opacity = 0;

                dummyForm.Controls.Add(dummyControl);
                dummyForm.Show();

                // Show the tooltip
                _notificationTip.Show(message, dummyControl, 20, 0, 3000);

                // Dispose of the form after the tooltip is shown
                Timer cleanupTimer = new Timer();
                cleanupTimer.Interval = 3100;  // Slightly longer than tooltip display time
                cleanupTimer.Tick += (s, e) =>
                {
                    dummyForm.Close();
                    dummyForm.Dispose();
                    cleanupTimer.Stop();
                    cleanupTimer.Dispose();
                };
                cleanupTimer.Start();
            }
            catch
            {
                // Fall back to MessageBox if tooltip fails
                MessageBox.Show(message, "ShapeMaster", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}