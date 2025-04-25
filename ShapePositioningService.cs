using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing shape positioning operations
    /// </summary>
    public class ShapePositioningService
    {
        private readonly PowerPoint.Application _application;
        private readonly Action<string, bool> _notificationCallback;

        /// <summary>
        /// Initializes a new instance of the ShapePositioningService class
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        /// <param name="notificationCallback">Callback for displaying notifications</param>
        public ShapePositioningService(PowerPoint.Application application, Action<string, bool> notificationCallback)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
        }

        /// <summary>
        /// Validates and gets two shapes from the current selection
        /// </summary>
        /// <returns>ShapeRange containing exactly two shapes, or null if invalid</returns>
        public PowerPoint.ShapeRange GetTwoSelectedShapes()
        {
            try
            {
                // Check if we're on a slide
                PowerPoint.Slide currentSlide = _application.ActiveWindow.View.Slide;
                if (currentSlide == null)
                {
                    _notificationCallback("Please navigate to a slide first.", false);
                    return null;
                }

                // Check if shapes are selected
                PowerPoint.Selection selection = _application.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _notificationCallback("Please select exactly two shapes.", false);
                    return null;
                }

                // Get selected shapes
                PowerPoint.ShapeRange shapes = selection.ShapeRange;

                // Check if we have exactly two shapes
                if (shapes.Count != 2)
                {
                    _notificationCallback("Please select exactly two shapes.", false);
                    return null;
                }

                return shapes;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error accessing selected shapes: {ex.Message}", true);
                return null;
            }
        }

        /// <summary>
        /// Swaps the positions of two shapes
        /// </summary>
        /// <param name="shapes">The ShapeRange containing exactly two shapes</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool SwapShapePositions(PowerPoint.ShapeRange shapes)
        {
            try
            {
                // Validate we have exactly two shapes
                if (shapes.Count != 2)
                {
                    _notificationCallback("Please select exactly two shapes.", false);
                    return false;
                }
                
                // Get the two shapes
                var shape1 = shapes[1];
                var shape2 = shapes[2];

                // Store original positions
                float shape1Left = shape1.Left;
                float shape1Top = shape1.Top;
                float shape2Left = shape2.Left;
                float shape2Top = shape2.Top;
                
                // Swap positions
                shape1.Left = shape2Left;
                shape1.Top = shape2Top;
                shape2.Left = shape1Left;
                shape2.Top = shape1Top;
                
                _notificationCallback("Positions swapped successfully.", false);
                return true;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error swapping positions: {ex.Message}", true);
                return false;
            }
        }
    }
}
