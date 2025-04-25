using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing ribbon UI operations and callbacks
    /// </summary>
    public class RibbonUIService
    {
        private readonly PowerPoint.Application _application;
        private readonly Action<string, bool> _notificationCallback;
        private readonly TextFormattingService _textFormattingService;
        private readonly ComObjectManager _comObjectManager;
        private Office.IRibbonUI _ribbonUI;

        /// <summary>
        /// Initializes a new instance of the RibbonUIService class
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        /// <param name="notificationCallback">Callback for displaying notifications</param>
        /// <param name="textFormattingService">Text formatting service for color operations</param>
        /// <param name="comObjectManager">COM object manager for handling COM object releases</param>
        public RibbonUIService(
            PowerPoint.Application application,
            Action<string, bool> notificationCallback,
            TextFormattingService textFormattingService,
            ComObjectManager comObjectManager = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
            _textFormattingService = textFormattingService ?? throw new ArgumentNullException(nameof(textFormattingService));
            _comObjectManager = comObjectManager; // Can be null for backward compatibility
        }

        /// <summary>
        /// Sets the active ribbon UI instance
        /// </summary>
        public void SetRibbonUI(Office.IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
        }

        /// <summary>
        /// Gets the active ribbon UI instance
        /// </summary>
        public Office.IRibbonUI GetRibbonUI()
        {
            return _ribbonUI;
        }

        /// <summary>
        /// Refreshes the ribbon UI to update all theme-dependent elements
        /// </summary>
        public void RefreshRibbonUI()
        {
            try
            {
                if (_ribbonUI != null)
                {
                    // Invalidate the main button image
                    _ribbonUI.InvalidateControl("ColorBoldTextMainButton");

                    // Invalidate each color menu item
                    for (int i = 1; i <= 10; i++)
                    {
                        _ribbonUI.InvalidateControl($"theme_color_{i}");
                    }

                    // Invalidate the menu item container
                    _ribbonUI.InvalidateControl("ColorBoldTextMenu");

                    // Invalidate the split button
                    _ribbonUI.InvalidateControl("ColorBoldTextSplitButton");
                }
            }
            catch (Exception ex)
            {
                // Log this, but don't show notification to avoid disrupting the user
                System.Diagnostics.Debug.WriteLine($"Error refreshing ribbon: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets an image for the main Color Bold Text button that shows the current theme color
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <returns>A bitmap representing the current color</returns>
        public System.Drawing.Bitmap GetColorBoldTextImage(Office.IRibbonControl control)
        {
            try
            {
                // Create a bitmap for the button (32x32 for large buttons)
                Bitmap bmp = new Bitmap(32, 32);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    // Get the current theme color index
                    Office.MsoThemeColorIndex themeColorIndex = _textFormattingService.GetThemeColor();

                    // Get the live theme color from PowerPoint
                    Color sampleColor = GetLiveThemeColor((int)themeColorIndex);

                    // Draw the text formatting icon background in white/light gray
                    g.FillRectangle(new SolidBrush(Color.WhiteSmoke), 0, 0, 32, 32);

                    // Draw a colored rectangle at the top to represent the current color
                    using (SolidBrush brush = new SolidBrush(sampleColor))
                    {
                        g.FillRectangle(brush, 0, 0, 32, 8);
                    }

                    // Draw a border around the colored rectangle
                    using (Pen pen = new Pen(Color.DarkGray, 1))
                    {
                        g.DrawRectangle(pen, 0, 0, 31, 7);
                    }

                    // Draw stylized "B" to represent bold text
                    using (Font boldFont = new Font("Arial", 14, FontStyle.Bold))
                    using (SolidBrush textBrush = new SolidBrush(Color.Black))
                    {
                        g.DrawString("B", boldFont, textBrush, 8, 10);
                    }

                    // Draw a line/underline to represent text
                    using (Pen pen = new Pen(Color.Black, 2))
                    {
                        g.DrawLine(pen, 8, 26, 24, 26);
                    }
                }

                return bmp;
            }
            catch (Exception)
            {
                // Return a default icon if there's an error
                Bitmap bmp = new Bitmap(32, 32);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.FillRectangle(new SolidBrush(Color.WhiteSmoke), 0, 0, 32, 32);
                    using (Font boldFont = new Font("Arial", 14, FontStyle.Bold))
                    using (SolidBrush textBrush = new SolidBrush(Color.Black))
                    {
                        g.DrawString("B", boldFont, textBrush, 8, 8);
                    }
                }
                return bmp;
            }
        }

        /// <summary>
        /// Gets a color swatch image for theme color buttons in the ribbon
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <returns>A bitmap with the color sample and text</returns>
        public System.Drawing.Bitmap GetThemeColorImage(Office.IRibbonControl control)
        {
            try
            {
                // Create a bitmap for the color sample
                Bitmap bmp = new Bitmap(16, 16);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    // Parse the color index from the control's tag
                    if (int.TryParse(control.Tag, out int colorIndexValue) &&
                        Enum.IsDefined(typeof(Office.MsoThemeColorIndex), colorIndexValue))
                    {
                        // Get the live theme color from PowerPoint
                        Color sampleColor = GetLiveThemeColor(colorIndexValue);

                        // Fill the background with the sample color
                        using (SolidBrush brush = new SolidBrush(sampleColor))
                        {
                            g.FillRectangle(brush, 0, 0, 16, 16);
                        }

                        // Add a border
                        using (Pen pen = new Pen(Color.DarkGray, 1))
                        {
                            g.DrawRectangle(pen, 0, 0, 15, 15);
                        }
                    }
                    else
                    {
                        // Draw a grey square if the color index is invalid
                        using (SolidBrush brush = new SolidBrush(Color.LightGray))
                        {
                            g.FillRectangle(brush, 0, 0, 16, 16);
                        }
                        using (Pen pen = new Pen(Color.DarkGray, 1))
                        {
                            g.DrawRectangle(pen, 0, 0, 15, 15);
                        }
                    }
                }

                return bmp;
            }
            catch (Exception)
            {
                // Return a default gray color sample if there's an error
                Bitmap bmp = new Bitmap(16, 16);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.FillRectangle(new SolidBrush(Color.LightGray), 0, 0, 16, 16);
                    g.DrawRectangle(new Pen(Color.DarkGray, 1), 0, 0, 15, 15);
                }
                return bmp;
            }
        }

        /// <summary>
        /// Gets an image for the swap positions button
        /// </summary>
        /// <param name="control">The ribbon control</param>
        /// <returns>A bitmap representing the swap positions action</returns>
        public System.Drawing.Bitmap GetSwapPositionsImage(Office.IRibbonControl control)
        {
            // Method 1: Load from embedded resource
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream imageStream = asm.GetManifestResourceStream("ShapeMaster.Images.SwapPositions.png");
            if (imageStream != null)
            {
                return new System.Drawing.Bitmap(imageStream);
            }

            // Method 2: Create a simple icon programmatically as fallback
            Bitmap bmp = new Bitmap(32, 32);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                // Draw two boxes with arrows to represent swapping
                using (Pen pen = new Pen(Color.Black, 1))
                {
                    // Draw first box
                    g.DrawRectangle(pen, 2, 2, 12, 12);
                    // Draw second box
                    g.DrawRectangle(pen, 18, 18, 12, 12);

                    pen.Color = Color.Blue;

                    // Draw arrow down
                    g.DrawLine(pen, 16, 8, 24, 8);
                    g.DrawLine(pen, 24, 8, 24, 15);
                    g.DrawLine(pen, 24, 15, 26, 13);
                    g.DrawLine(pen, 24, 15, 22, 13);
                    // Draw arrow up
                    g.DrawLine(pen, 16, 24, 8, 24);
                    g.DrawLine(pen, 8, 24, 8, 16);
                    g.DrawLine(pen, 8, 16, 10, 18);
                    g.DrawLine(pen, 8, 16, 6, 18);
                }
            }

            return bmp;
        }

        /// <summary>
        /// Gets a user-friendly name for a theme color
        /// </summary>
        public string GetThemeColorName(Office.MsoThemeColorIndex colorIndex)
        {
            switch (colorIndex)
            {
                case Office.MsoThemeColorIndex.msoThemeColorAccent1:
                    return "Accent 1";
                case Office.MsoThemeColorIndex.msoThemeColorAccent2:
                    return "Accent 2";
                case Office.MsoThemeColorIndex.msoThemeColorAccent3:
                    return "Accent 3";
                case Office.MsoThemeColorIndex.msoThemeColorAccent4:
                    return "Accent 4";
                case Office.MsoThemeColorIndex.msoThemeColorAccent5:
                    return "Accent 5";
                case Office.MsoThemeColorIndex.msoThemeColorAccent6:
                    return "Accent 6";
                case Office.MsoThemeColorIndex.msoThemeColorBackground1:
                    return "Background 1";
                case Office.MsoThemeColorIndex.msoThemeColorBackground2:
                    return "Background 2";
                case Office.MsoThemeColorIndex.msoThemeColorText1:
                    return "Text 1";
                case Office.MsoThemeColorIndex.msoThemeColorText2:
                    return "Text 2";
                case Office.MsoThemeColorIndex.msoThemeColorHyperlink:
                    return "Hyperlink";
                case Office.MsoThemeColorIndex.msoThemeColorFollowedHyperlink:
                    return "Followed Hyperlink";
                default:
                    return "Theme Color";
            }
        }

        /// <summary>
        /// Gets the live theme color from the current PowerPoint theme
        /// </summary>
        /// <param name="colorIndex">The theme color index</param>
        /// <returns>A Color object representing the current theme color</returns>
        private Color GetLiveThemeColor(int colorIndex)
        {
            // Convert the int to the theme color enum
            Office.MsoThemeColorIndex themeColorIndex = (Office.MsoThemeColorIndex)colorIndex;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tempShape = null;
            bool tempSlideCreated = false;

            try
            {
                // Check if there's an active presentation
                if (_application.ActivePresentation != null)
                {
                    // Try to get the active slide, or create a temporary one if needed
                    try
                    {
                        if (_application.ActiveWindow != null &&
                            _application.ActiveWindow.View != null)
                        {
                            slide = _application.ActiveWindow.View.Slide;
                        }
                    }
                    catch
                    {
                        // Ignore error and will create a temp slide below
                    }

                    if (slide == null)
                    {
                        // If no slide is active, create a temporary one
                        slide = _application.ActivePresentation.Slides.Add(
                            _application.ActivePresentation.Slides.Count + 1,
                            PowerPoint.PpSlideLayout.ppLayoutBlank);
                        tempSlideCreated = true;
                    }

                    // Create a temporary shape to sample the color
                    tempShape = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        0, 0, 1, 1);

                    // Set the shape's color to the theme color we want to sample
                    tempShape.TextFrame.TextRange.Font.Color.ObjectThemeColor = themeColorIndex;

                    // Get the RGB values
                    int red = tempShape.TextFrame.TextRange.Font.Color.RGB & 0xFF;
                    int green = (tempShape.TextFrame.TextRange.Font.Color.RGB & 0xFF00) >> 8;
                    int blue = (tempShape.TextFrame.TextRange.Font.Color.RGB & 0xFF0000) >> 16;

                    // Return the actual theme color as an RGB value
                    return Color.FromArgb(red, green, blue);
                }
            }
            catch
            {
                // Fall back to default colors if there's any error
            }
            finally
            {
                // Clean up COM objects
                if (tempShape != null)
                {
                    try
                    {
                        tempShape.Delete();
                    }
                    catch
                    {
                        // Ignore errors during deletion
                    }

                    // Release the COM object
                    if (_comObjectManager != null)
                    {
                        _comObjectManager.ReleaseComObject(tempShape, "Temporary Shape");
                    }
                }

                // Delete and release temporary slide if we created one
                if (tempSlideCreated && slide != null)
                {
                    try
                    {
                        slide.Delete();
                    }
                    catch
                    {
                        // Ignore errors during deletion
                    }
                }

                // Release slide COM object
                if (slide != null && _comObjectManager != null)
                {
                    _comObjectManager.ReleaseComObject(slide, "Slide");
                }
            }

            // Fallback to hard-coded values if we couldn't get the actual theme color
            // this hasn't happen in real life yet, we might consider removing this safetynet
            switch (themeColorIndex)
            {
                case Office.MsoThemeColorIndex.msoThemeColorAccent1:
                    return Color.FromArgb(79, 129, 189); // Blue
                case Office.MsoThemeColorIndex.msoThemeColorAccent2:
                    return Color.FromArgb(192, 80, 77);  // Red
                case Office.MsoThemeColorIndex.msoThemeColorAccent3:
                    return Color.FromArgb(155, 187, 89); // Green
                case Office.MsoThemeColorIndex.msoThemeColorAccent4:
                    return Color.FromArgb(128, 100, 162); // Purple
                case Office.MsoThemeColorIndex.msoThemeColorAccent5:
                    return Color.FromArgb(75, 172, 198); // Cyan
                case Office.MsoThemeColorIndex.msoThemeColorAccent6:
                    return Color.FromArgb(247, 150, 70); // Orange
                case Office.MsoThemeColorIndex.msoThemeColorBackground1:
                    return Color.White;
                case Office.MsoThemeColorIndex.msoThemeColorBackground2:
                    return Color.FromArgb(242, 242, 242); // Light Gray
                case Office.MsoThemeColorIndex.msoThemeColorText1:
                    return Color.Black;
                case Office.MsoThemeColorIndex.msoThemeColorText2:
                    return Color.FromArgb(89, 89, 89);   // Dark Gray
                case Office.MsoThemeColorIndex.msoThemeColorHyperlink:
                    return Color.FromArgb(0, 102, 204);  // Blue
                case Office.MsoThemeColorIndex.msoThemeColorFollowedHyperlink:
                    return Color.FromArgb(149, 79, 114); // Purple
                default:
                    return Color.Gray;
            }
        }

        public System.Drawing.Bitmap GetNoteImage(Office.IRibbonControl control)
        {
            // Default color and label
            Color fillColor = Color.LightYellow;
            string label = "";
            try
            {
                if (control != null && !string.IsNullOrEmpty(control.Tag))
                {
                    var parts = control.Tag.Split('|');
                    if (parts.Length > 0)
                    {
                        string colorString = parts[0];
                        if (colorString.StartsWith("#"))
                            colorString = colorString.Substring(1);
                        if (int.TryParse(colorString, System.Globalization.NumberStyles.HexNumber, null, out int hex))
                        {
                            // Use BGR order as defined in the XML
                            fillColor = Color.FromArgb(hex & 0xFF, (hex >> 8) & 0xFF, (hex >> 16) & 0xFF);
                        }
                    }
                    if (parts.Length > 1)
                    {
                        label = parts[1].Trim();
                    }
                }
            }
            catch { }

            Bitmap bmp = new Bitmap(32, 32);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                // Draw snipped rectangle (sticky note style)
                Point[] points = new Point[] {
                    new Point(0, 0), new Point(31, 0), new Point(31, 24), new Point(24, 31), new Point(0, 31)
                };
                using (SolidBrush brush = new SolidBrush(fillColor))
                {
                    g.FillPolygon(brush, points);
                }
                using (Pen pen = new Pen(Color.DarkGray, 1))
                {
                    g.DrawPolygon(pen, points);
                }
                // Optionally overlay a letter for the note type
                string overlay = "";
                if (label.StartsWith("TODO", StringComparison.OrdinalIgnoreCase)) overlay = "T";
                else if (label.StartsWith("Review", StringComparison.OrdinalIgnoreCase)) overlay = "R";
                else if (label.StartsWith("Comment", StringComparison.OrdinalIgnoreCase)) overlay = "C";
                if (!string.IsNullOrEmpty(overlay))
                {
                    using (Font font = new Font("Arial", 14, FontStyle.Bold))
                    using (SolidBrush textBrush = new SolidBrush(Color.Black))
                    {
                        g.DrawString(overlay, font, textBrush, 8, 6);
                    }
                }
            }
            return bmp;
        }
    }
}