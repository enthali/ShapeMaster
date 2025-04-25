using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing shape positioning operations
    /// </summary>
    public class ShapePositioningService
    {
        private readonly PowerPoint.Application _application;
        private readonly Action<string, bool> _notificationCallback;
        private readonly ComObjectManager _comObjectManager;

        /// <summary>
        /// Initializes a new instance of the ShapePositioningService class
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        /// <param name="notificationCallback">Callback for displaying notifications</param>
        /// <param name="comObjectManager">COM object manager for handling COM object releases</param>
        public ShapePositioningService(
            PowerPoint.Application application,
            Action<string, bool> notificationCallback,
            ComObjectManager comObjectManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
            _comObjectManager = comObjectManager ?? throw new ArgumentNullException(nameof(comObjectManager));
        }

        /// <summary>
        /// Validates and gets two shapes from the current selection
        /// </summary>
        /// <returns>ShapeRange containing exactly two shapes, or null if invalid</returns>
        public PowerPoint.ShapeRange GetTwoSelectedShapes()
        {
            PowerPoint.Slide currentSlide = null;
            PowerPoint.Selection selection = null;
            PowerPoint.ShapeRange shapes = null;

            try
            {
                // Check if we're on a slide
                currentSlide = _application.ActiveWindow.View.Slide;
                if (currentSlide == null)
                {
                    _notificationCallback("Please navigate to a slide first.", false);
                    return null;
                }

                // Check if shapes are selected
                selection = _application.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _notificationCallback("Please select exactly two shapes.", false);
                    return null;
                }

                // Get selected shapes
                shapes = selection.ShapeRange;

                // Check if we have exactly two shapes
                if (shapes.Count != 2)
                {
                    _notificationCallback("Please select exactly two shapes.", false);

                    // Release the COM object before returning null
                    _comObjectManager.ReleaseComObject(shapes, "ShapeRange");
                    shapes = null;
                    return null;
                }

                // Return the shapes without releasing them - the caller is responsible for releasing
                return shapes;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error accessing selected shapes: {ex.Message}", true);

                // Clean up any COM objects we created
                if (shapes != null)
                {
                    _comObjectManager.ReleaseComObject(shapes, "ShapeRange");
                }
                return null;
            }
            finally
            {
                // Always release these temporary COM objects
                if (selection != null) _comObjectManager.ReleaseComObject(selection, "Selection");
                if (currentSlide != null) _comObjectManager.ReleaseComObject(currentSlide, "Slide");
            }
        }

        /// <summary>
        /// Swaps the positions of two shapes
        /// </summary>
        /// <param name="shapes">The ShapeRange containing exactly two shapes</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool SwapShapePositions(PowerPoint.ShapeRange shapes)
        {
            PowerPoint.Shape shape1 = null;
            PowerPoint.Shape shape2 = null;

            try
            {
                // Validate we have exactly two shapes
                if (shapes.Count != 2)
                {
                    _notificationCallback("Please select exactly two shapes.", false);
                    return false;
                }

                // Get the two shapes
                shape1 = shapes[1];
                shape2 = shapes[2];

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
            finally
            {
                // Release the individual shape objects
                if (shape1 != null) _comObjectManager.ReleaseComObject(shape1, "Shape1");
                if (shape2 != null) _comObjectManager.ReleaseComObject(shape2, "Shape2");

                // Note: We don't release the shapes collection here as it was passed in
                // and should be released by the caller
            }
        }

        /// <summary>
        /// Helper method to release a ShapeRange if it's no longer needed
        /// </summary>
        /// <param name="shapeRange">The ShapeRange to release</param>
        public void ReleaseShapeRange(PowerPoint.ShapeRange shapeRange)
        {
            if (shapeRange != null)
            {
                _comObjectManager.ReleaseComObject(shapeRange, "ShapeRange");
            }
        }
    }
}
