using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ShapeMaster
{
    public partial class ThisAddIn
    {
        // Service manager instance
        private Services.ServiceManager _serviceManager;
        // Reference to the ribbon UI (kept for backward compatibility)
        private Office.IRibbonUI _ribbonUI;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Initialize the service manager, which will handle all service initialization
                _serviceManager = Services.ServiceManager.Instance(Application);

                // We won't set the default theme color here to avoid errors when no document is open
                // The color will be initialized when first needed
            }
            catch (Exception ex)
            {
                // Use MessageBox for startup errors as our tooltip system might not be ready
                System.Windows.Forms.MessageBox.Show($"Error during startup: {ex.Message}", "ShapeMaster Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Shutdown all services via the service manager
                if (_serviceManager != null)
                {
                    _serviceManager.Shutdown();
                }
            }
            catch
            {
                // Ignore errors during shutdown
            }
        }

        /// <summary>
        /// Swaps the positions of exactly two selected shapes.
        /// </summary>
        public void SwapSelectedShapePositions()
        {
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
            finally
            {
                // Always release the shape range when done
                if (shapes != null)
                {
                    _serviceManager.ComObjectManager.ReleaseComObject(shapes, "ShapeRange in SwapSelectedShapePositions");
                }
            }
        }


        /// <summary>
        /// Colors bold text in the selected shapes using the current theme color.
        /// </summary>
        public void ColorBoldTextInSelectedShapes()
        {
            PowerPoint.ShapeRange shapes = null;
            try
            {
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
                ShowNotification($"Error applying color: {ex.Message}", true);
            }
            finally
            {
                // Always release the shape range when done
                if (shapes != null)
                {
                    _serviceManager.ComObjectManager.ReleaseComObject(shapes, "ShapeRange in ColorBoldTextInSelectedShapes");
                }
            }
        }

        /// <summary>
        /// Gets the notification service instance
        /// </summary>
        public Services.NotificationService NotificationService => _serviceManager.NotificationService;

        /// <summary>
        /// Shows a notification using the notification service
        /// This method exists for backward compatibility
        /// </summary>
        public void ShowNotification(string message, bool isError = false)
        {
            _serviceManager.NotificationService.ShowNotification(message, isError);
        }

        /// <summary>
        /// Gets the event handling service instance
        /// </summary>
        public Services.EventHandlingService EventHandlingService => _serviceManager.EventHandlingService;

        /// <summary>
        /// Gets the ribbon UI service instance
        /// </summary>
        public Services.RibbonUIService RibbonUIService => _serviceManager.RibbonUIService;

        /// <summary>
        /// Gets the shape resizing service instance
        /// </summary>
        public Services.ShapeResizingService ShapeResizingService => _serviceManager.ShapeResizingService;

        /// <summary>
        /// Gets the text formatting service instance
        /// </summary>
        public Services.TextFormattingService TextFormattingService => _serviceManager.TextFormattingService;

        /// <summary>
        /// Resizes all selected shapes to match the dimensions of the first selected shape.
        /// </summary>
        public void ResizeSelectedShapesToMatch()
        {
            try
            {
                _serviceManager.ShapeResizingService.ResizeSelectedShapesToMatch();
            }
            catch (Exception ex)
            {
                ShowNotification($"Error resizing shapes: {ex.Message}", true);
            }
        }

        /// <summary>
        /// Resizes width of all selected shapes to match the width of the first selected shape.
        /// </summary>
        public void ResizeSelectedShapesToMatchWidth()
        {
            try
            {
                _serviceManager.ShapeResizingService.ResizeSelectedShapesToMatchWidth();
            }
            catch (Exception ex)
            {
                ShowNotification($"Error resizing shapes: {ex.Message}", true);
            }
        }

        /// <summary>
        /// Resizes height of all selected shapes to match the height of the first selected shape.
        /// </summary>
        public void ResizeSelectedShapesToMatchHeight()
        {
            try
            {
                _serviceManager.ShapeResizingService.ResizeSelectedShapesToMatchHeight();
            }
            catch (Exception ex)
            {
                ShowNotification($"Error resizing shapes: {ex.Message}", true);
            }
        }

        /// <summary>
        /// Gets the active ribbon UI instance
        /// </summary>
        public Office.IRibbonUI GetActiveRibbon()
        {
            return _ribbonUI;
        }

        /// <summary>
        /// Sets the active ribbon UI instance
        /// </summary>
        public void SetRibbonUI(Office.IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI; // Keep for backward compatibility for now

            // The ServiceManager should already be initialized
            if (_serviceManager != null)
            {
                _serviceManager.SetRibbonUI(ribbonUI);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Warning: ServiceManager not initialized when SetRibbonUI was called.");
            }
        }

        /// <summary>
        /// Refreshes the ribbon UI to update all theme-dependent elements
        /// </summary>
        public void RefreshRibbonUI()
        {
            // The ServiceManager should already be initialized
            if (_serviceManager != null && _serviceManager.RibbonUIService != null)
            {
                _serviceManager.RibbonUIService.RefreshRibbonUI();
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Warning: ServiceManager or RibbonUIService not initialized when RefreshRibbonUI was called.");

                // Fallback to direct implementation if service is not available
                try
                {
                    if (_ribbonUI != null)
                    {
                        _ribbonUI.InvalidateControl("ColorBoldTextMainButton");

                        for (int i = 1; i <= 10; i++)
                        {
                            _ribbonUI.InvalidateControl($"theme_color_{i}");
                        }

                        _ribbonUI.InvalidateControl("ColorBoldTextMenu");
                        _ribbonUI.InvalidateControl("ColorBoldTextSplitButton");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error refreshing ribbon: {ex.Message}");
                }
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ShapeMasterRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}