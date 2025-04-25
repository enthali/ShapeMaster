using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ShapeMaster.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace ShapeMaster
{
    [ComVisible(true)]
    public class ShapeMasterRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        // Reference to service manager
        private ServiceManager _serviceManager;

        public ShapeMasterRibbon()
        {
            // Service manager will be initialized when Ribbon_Load is called
            // This ensures we have access to the PowerPoint Application instance
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ShapeMaster.ShapeMasterRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            try
            {
                // Get the PowerPoint application instance first
                PowerPoint.Application application = Globals.ThisAddIn.Application;

                // Get the service manager instance, passing the application if it needs to be initialized
                _serviceManager = ServiceManager.Instance(application);

                // Set the ribbon UI in the service manager
                _serviceManager.SetRibbonUI(ribbonUI);

                // For backward compatibility - while we transition to full ServiceManager usage
                Globals.ThisAddIn.SetRibbonUI(ribbonUI);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Ribbon_Load: {ex.Message}");
                // Cannot show notification here as services might not be initialized yet
            }
        }

        public void OnResizeButtonClick(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnResizeButtonClick");
                return;
            }

            try
            {
                _serviceManager.ShapeResizingService.ResizeSelectedShapesToMatch();
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error resizing shapes: {ex.Message}";
                _serviceManager.NotificationService.ShowNotification(errorMsg, true);
            }
        }

        public void OnResizeWidthButtonClick(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnResizeWidthButtonClick");
                return;
            }

            try
            {
                _serviceManager.ShapeResizingService.ResizeSelectedShapesToMatchWidth();
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error resizing shapes: {ex.Message}";
                _serviceManager.NotificationService.ShowNotification(errorMsg, true);
            }
        }

        public void OnResizeHeightButtonClick(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnResizeHeightButtonClick");
                return;
            }

            try
            {
                _serviceManager.ShapeResizingService.ResizeSelectedShapesToMatchHeight();
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error resizing shapes: {ex.Message}";
                _serviceManager.NotificationService.ShowNotification(errorMsg, true);
            }
        }

        public void OnSwapPositionsButtonClick(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnSwapPositionsButtonClick");
                return;
            }

            PowerPoint.ShapeRange shapes = null;

            try
            {
                // Get valid shapes (exactly 2) using the service
                shapes = _serviceManager.ShapePositioningService.GetTwoSelectedShapes();
                if (shapes != null)
                {
                    // Swap positions using the service
                    _serviceManager.ShapePositioningService.SwapShapePositions(shapes);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error swapping positions: {ex.Message}";
                _serviceManager.NotificationService.ShowNotification(errorMsg, true);
            }
            finally
            {
                // Always release the shapes collection when done
                if (shapes != null && _serviceManager.ComObjectManager != null)
                {
                    _serviceManager.ComObjectManager.ReleaseComObject(shapes, "ShapeRange in OnSwapPositionsButtonClick");
                }
            }
        }

        public void OnColorBoldTextClick(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnColorBoldTextClick");
                return;
            }

            PowerPoint.ShapeRange shapes = null;

            try
            {
                // Check for active presentation first
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow == null ||
                    app.ActiveWindow.View == null ||
                    app.ActiveWindow.View.Slide == null)
                {
                    _serviceManager.NotificationService.ShowNotification("Please open a presentation and navigate to a slide first.", true);
                    return;
                }

                // Get shapes with text
                shapes = _serviceManager.TextFormattingService.GetSelectedShapesWithText();
                if (shapes != null)
                {
                    // Apply coloring to bold text using the current theme color
                    _serviceManager.TextFormattingService.ColorBoldText(shapes);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error: {ex.Message}";
                _serviceManager.NotificationService.ShowNotification(errorMsg, true);
            }
            finally
            {
                // Always release the shapes collection when done
                if (shapes != null && _serviceManager.ComObjectManager != null)
                {
                    _serviceManager.ComObjectManager.ReleaseComObject(shapes, "ShapeRange in OnColorBoldTextClick");
                }
            }
        }

        /// <summary>
        /// Simple theme color selection handler - direct approach
        /// </summary>
        public void OnSimpleThemeColorSelected(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnSimpleThemeColorSelected");
                return;
            }

            PowerPoint.ShapeRange shapes = null;

            try
            {
                // Check for active presentation first
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow == null ||
                    app.ActiveWindow.View == null ||
                    app.ActiveWindow.View.Slide == null)
                {
                    _serviceManager.NotificationService.ShowNotification("Please open a presentation and navigate to a slide first.", true);
                    return;
                }

                // Parse the color index from the tag
                if (control != null && control.Tag != null && int.TryParse(control.Tag, out int colorIndexValue))
                {
                    // Make sure it's a valid enum value
                    if (Enum.IsDefined(typeof(Office.MsoThemeColorIndex), colorIndexValue))
                    {
                        // Convert to MsoThemeColorIndex
                        Office.MsoThemeColorIndex themeColor = (Office.MsoThemeColorIndex)colorIndexValue;

                        // Set the theme color in the service
                        _serviceManager.TextFormattingService.SetThemeColor(themeColor);

                        // Apply the color to selected shapes
                        shapes = _serviceManager.TextFormattingService.GetSelectedShapesWithText();
                        if (shapes != null)
                        {
                            _serviceManager.TextFormattingService.ColorBoldText(shapes);
                        }

                        // Notify the user
                        string colorName = _serviceManager.RibbonUIService.GetThemeColorName(themeColor);
                        _serviceManager.NotificationService.ShowNotification($"Applied {colorName} to bold text", false);
                    }
                    else
                    {
                        _serviceManager.NotificationService.ShowNotification($"Invalid color index: {colorIndexValue}", true);
                    }
                }
                else
                {
                    _serviceManager.NotificationService.ShowNotification("Invalid or missing color selection", true);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Error applying color: {ex.Message}";
                _serviceManager.NotificationService.ShowNotification(errorMsg, true);
            }
            finally
            {
                // Always release the shapes collection when done
                if (shapes != null && _serviceManager.ComObjectManager != null)
                {
                    _serviceManager.ComObjectManager.ReleaseComObject(shapes, "ShapeRange in OnSimpleThemeColorSelected");
                }
            }
        }

        /// <summary>
        /// Gets a user-friendly name for a theme color
        /// </summary>
        private string GetSimpleThemeColorName(Office.MsoThemeColorIndex colorIndex)
        {
            return _serviceManager.RibbonUIService.GetThemeColorName(colorIndex);
        }

        // Generic handler for all note buttons
        public void OnNoteButtonClick(Office.IRibbonControl control)
        {
            if (_serviceManager == null)
            {
                System.Diagnostics.Debug.WriteLine("Error: ServiceManager not initialized in OnNoteButtonClick");
                return;
            }

            try
            {
                // Expecting control.Tag in the format "#RRGGBB|label text"
                string color = null;
                string label = null;
                if (!string.IsNullOrEmpty(control.Tag) && control.Tag.Contains("|"))
                {
                    var parts = control.Tag.Split('|');
                    color = parts[0];
                    label = parts.Length > 1 ? parts[1] : "note";
                }
                else
                {
                    color = control.Tag ?? "#FFFF00";
                    label = "note";
                }
                _serviceManager.NoteService.CreateNoteShape(color, label);
            }
            catch (Exception ex)
            {
                _serviceManager.NotificationService.ShowNotification($"Error inserting note: {ex.Message}", true);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public System.Drawing.Bitmap GetSwapPositionsImage(Office.IRibbonControl control)
        {
            if (_serviceManager == null || _serviceManager.RibbonUIService == null)
            {
                // Return a default image if the service manager isn't initialized
                return CreateDefaultImage();
            }
            return _serviceManager.RibbonUIService.GetSwapPositionsImage(control);
        }

        public System.Drawing.Bitmap GetColorBoldTextImage(Office.IRibbonControl control)
        {
            if (_serviceManager == null || _serviceManager.RibbonUIService == null)
            {
                // Return a default image if the service manager isn't initialized
                return CreateDefaultImage();
            }
            return _serviceManager.RibbonUIService.GetColorBoldTextImage(control);
        }

        public System.Drawing.Bitmap GetThemeColorImage(Office.IRibbonControl control)
        {
            if (_serviceManager == null || _serviceManager.RibbonUIService == null)
            {
                // Return a default image if the service manager isn't initialized
                return CreateDefaultImage();
            }
            return _serviceManager.RibbonUIService.GetThemeColorImage(control);
        }

        public System.Drawing.Bitmap GetNoteImage(Office.IRibbonControl control)
        {
            if (_serviceManager == null || _serviceManager.RibbonUIService == null)
            {
                // Return a default image if the service manager isn't initialized
                return CreateDefaultImage();
            }
            return _serviceManager.RibbonUIService.GetNoteImage(control);
        }

        /// <summary>
        /// Creates a simple default image for ribbon buttons when services aren't initialized
        /// </summary>
        private System.Drawing.Bitmap CreateDefaultImage()
        {
            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(32, 32);
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
            {
                g.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.LightGray), 0, 0, 32, 32);
                g.DrawRectangle(new System.Drawing.Pen(System.Drawing.Color.DarkGray, 1), 0, 0, 31, 31);
            }
            return bmp;
        }

        #endregion

    }
}