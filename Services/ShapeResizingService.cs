using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing shape resizing operations
    /// </summary>
    public class ShapeResizingService
    {
        private readonly PowerPoint.Application _application;
        private readonly Action<string, bool> _notificationCallback;
        private readonly ErrorHandlingService _errorHandlingService;
        private readonly ComObjectManager _comObjectManager;

        /// <summary>
        /// Initializes a new instance of the ShapeResizingService class
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        /// <param name="notificationCallback">Callback for displaying notifications</param>
        /// <param name="errorHandlingService">Error handling service</param>
        /// <param name="comObjectManager">COM object manager for handling COM object releases</param>
        public ShapeResizingService(
            PowerPoint.Application application,
            Action<string, bool> notificationCallback,
            ErrorHandlingService errorHandlingService = null,
            ComObjectManager comObjectManager = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
            _errorHandlingService = errorHandlingService; // Can be null for backward compatibility
            _comObjectManager = comObjectManager ?? throw new ArgumentNullException(nameof(comObjectManager));
        }

        /// <summary>
        /// Gets the currently selected shapes if they meet the requirements for resizing operations
        /// </summary>
        /// <returns>Selected shapes if valid, null otherwise</returns>
        public PowerPoint.ShapeRange GetValidSelectedShapes()
        {
            return _errorHandlingService != null
                ? _errorHandlingService.TryExecute(GetValidSelectedShapesCore, (PowerPoint.ShapeRange)null,
                    "Unable to access selected shapes.", true, "GetValidSelectedShapes")
                : GetValidSelectedShapesLegacy();
        }

        /// <summary>
        /// Legacy implementation of GetValidSelectedShapes for backward compatibility
        /// </summary>
        private PowerPoint.ShapeRange GetValidSelectedShapesLegacy()
        {
            try
            {
                return GetValidSelectedShapesCore();
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error accessing selected shapes: {ex.Message}", true);
                return null;
            }
        }

        /// <summary>
        /// Core implementation of GetValidSelectedShapes
        /// </summary>
        private PowerPoint.ShapeRange GetValidSelectedShapesCore()
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
                    _notificationCallback("Please select at least two shapes.", false);
                    return null;
                }

                // Get selected shapes
                shapes = selection.ShapeRange;

                // Check if we have at least two shapes
                if (shapes.Count < 2)
                {
                    _notificationCallback("Please select at least two shapes.", false);

                    // Release shapes before returning null
                    _comObjectManager.ReleaseComObject(shapes, "Invalid ShapeRange");
                    shapes = null;

                    return null;
                }

                // Return the shapes without releasing them - caller is responsible for release
                return shapes;
            }
            catch (Exception ex)
            {
                // In case of error, release any COM objects we created
                if (shapes != null)
                {
                    _comObjectManager.ReleaseComObject(shapes, "ShapeRange in error");
                }

                throw; // Rethrow so the error handling service can handle it
            }
            finally
            {
                // Always clean up temporary COM objects
                if (selection != null) _comObjectManager.ReleaseComObject(selection, "Selection");
                if (currentSlide != null) _comObjectManager.ReleaseComObject(currentSlide, "Slide");
            }
        }

        /// <summary>
        /// Generic method to resize shapes based on the specified resize action
        /// </summary>
        private void ResizeShapesHelper(PowerPoint.ShapeRange shapes, Action<PowerPoint.ShapeRange, float, float> resizeAction, string successMessage)
        {
            if (_errorHandlingService != null)
            {
                _errorHandlingService.TryExecute(
                    () => ResizeShapesHelperCore(shapes, resizeAction, successMessage),
                    $"Error resizing shapes.", true, "ResizeShapesHelper");
            }
            else
            {
                ResizeShapesHelperLegacy(shapes, resizeAction, successMessage);
            }

            // Always release the shapes collection after resizing
            if (shapes != null)
            {
                _comObjectManager.ReleaseComObject(shapes, "ShapeRange after resizing");
            }
        }

        /// <summary>
        /// Legacy implementation of ResizeShapesHelper for backward compatibility
        /// </summary>
        private void ResizeShapesHelperLegacy(PowerPoint.ShapeRange shapes, Action<PowerPoint.ShapeRange, float, float> resizeAction, string successMessage)
        {
            try
            {
                ResizeShapesHelperCore(shapes, resizeAction, successMessage);
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error resizing shapes: {ex.Message}", true);
            }
            finally
            {
                // Always release the shapes collection after resizing
                if (shapes != null)
                {
                    _comObjectManager.ReleaseComObject(shapes, "ShapeRange after resizing");
                }
            }
        }

        /// <summary>
        /// Core implementation of ResizeShapesHelper
        /// </summary>
        private void ResizeShapesHelperCore(PowerPoint.ShapeRange shapes, Action<PowerPoint.ShapeRange, float, float> resizeAction, string successMessage)
        {
            // Get dimensions of first shape (the reference shape)
            float referenceWidth = shapes[1].Width;
            float referenceHeight = shapes[1].Height;

            // Apply the resize action
            resizeAction(shapes, referenceWidth, referenceHeight);

            // Show success notification with count
            int count = shapes.Count - 1;
            string countText = count == 1 ? "1 shape" : $"{count} shapes";
            _notificationCallback(successMessage.Replace("{count}", countText), false);
        }

        /// <summary>
        /// Resizes all selected shapes to match the dimensions of the first selected shape.
        /// </summary>
        public void ResizeSelectedShapesToMatch()
        {
            // Get valid shapes
            PowerPoint.ShapeRange shapes = GetValidSelectedShapes();
            if (shapes == null) return;

            ResizeShapesHelper(shapes, (shapesToResize, refWidth, refHeight) =>
            {
                // Resize all other shapes to match first shape
                for (int i = 2; i <= shapesToResize.Count; i++)
                {
                    shapesToResize[i].Width = refWidth;
                    shapesToResize[i].Height = refHeight;
                }
            }, "{count} resized to match the first selected shape.");
        }

        /// <summary>
        /// Resizes width of all selected shapes to match the width of the first selected shape.
        /// </summary>
        public void ResizeSelectedShapesToMatchWidth()
        {
            // Get valid shapes
            PowerPoint.ShapeRange shapes = GetValidSelectedShapes();
            if (shapes == null) return;

            ResizeShapesHelper(shapes, (shapesToResize, refWidth, refHeight) =>
            {
                // Resize width of all other shapes to match first shape
                for (int i = 2; i <= shapesToResize.Count; i++)
                {
                    shapesToResize[i].Width = refWidth;
                }
            }, "{count} resized to match the width of the first selected shape.");
        }

        /// <summary>
        /// Resizes height of all selected shapes to match the height of the first selected shape.
        /// </summary>
        public void ResizeSelectedShapesToMatchHeight()
        {
            // Get valid shapes
            PowerPoint.ShapeRange shapes = GetValidSelectedShapes();
            if (shapes == null) return;

            ResizeShapesHelper(shapes, (shapesToResize, refWidth, refHeight) =>
            {
                // Resize height of all other shapes to match first shape
                for (int i = 2; i <= shapesToResize.Count; i++)
                {
                    shapesToResize[i].Height = refHeight;
                }
            }, "{count} resized to match the height of the first selected shape.");
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