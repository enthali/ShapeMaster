using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service class for managing text formatting operations
    /// </summary>
    public class TextFormattingService
    {
        private readonly PowerPoint.Application _application;
        private readonly Action<string, bool> _notificationCallback;
        private readonly ComObjectManager _comObjectManager;

        // Default theme color to use
        private Office.MsoThemeColorIndex _themeColorIndex = Office.MsoThemeColorIndex.msoThemeColorAccent1;
        // This is our changeToColor that we'll use for operations
        private PowerPoint.ColorFormat _changeToColor = null;

        /// <summary>
        /// Initializes a new instance of the TextFormattingService class
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        /// <param name="notificationCallback">Callback for displaying notifications</param>
        /// <param name="comObjectManager">COM object manager for handling COM object releases</param>
        public TextFormattingService(
            PowerPoint.Application application,
            Action<string, bool> notificationCallback,
            ComObjectManager comObjectManager)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
            _comObjectManager = comObjectManager ?? throw new ArgumentNullException(nameof(comObjectManager));
        }

        /// <summary>
        /// Gets the current theme color index for text formatting
        /// </summary>
        public Office.MsoThemeColorIndex GetThemeColor()
        {
            return _themeColorIndex;
        }

        /// <summary>
        /// Sets the theme color to use for text formatting
        /// </summary>
        /// <param name="themeColor">The theme color index to use</param>
        public void SetThemeColor(Office.MsoThemeColorIndex themeColor)
        {
            try
            {
                _themeColorIndex = themeColor;

                // Clean up any previous ColorFormat object
                if (_changeToColor != null)
                {
                    _comObjectManager.ReleaseComObject(_changeToColor, "Previous ColorFormat");
                    _changeToColor = null;
                }

                // Try to update the changeToColor if an active document exists
                PowerPoint.ColorFormat colorFormat = GetColorFromTheme(themeColor);
                if (colorFormat != null)
                {
                    _changeToColor = colorFormat;
                    _notificationCallback($"Theme color set to {GetThemeColorName(themeColor)}", false);

                    // Refresh the entire ribbon UI to update all theme-dependent elements
                    try
                    {
                        // Use the centralized ribbon refresh method
                        Globals.ThisAddIn.RefreshRibbonUI();
                    }
                    catch
                    {
                        // Ignore errors with ribbon updating
                    }
                }
                else
                {
                    // Just save the theme color index without creating a ColorFormat object
                    // We'll create the actual color when needed and a document is available
                    // Don't show notification here since GetColorFromTheme already showed one
                }
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error setting theme color: {ex.Message}", true);
            }
        }

        /// <summary>
        /// Resets the cached color format objects, forcing them to be recreated with current theme colors
        /// </summary>
        public void ResetColorCache()
        {
            try
            {
                // Clear the cached color (with proper COM cleanup)
                if (_changeToColor != null)
                {
                    _comObjectManager.ReleaseComObject(_changeToColor, "Cached ColorFormat");
                    _changeToColor = null;
                }
            }
            catch
            {
                // Silently fail to avoid disrupting PowerPoint
                // We'll recreate the color when needed
            }
        }

        /// <summary>
        /// Gets a user-friendly name for a theme color
        /// </summary>
        private string GetThemeColorName(Office.MsoThemeColorIndex themeColor)
        {
            switch (themeColor)
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
        /// Sets the theme color to the secondary theme color (Accent2)
        /// </summary>
        public void SetToSecondaryThemeColor()
        {
            try
            {
                // Check for active presentation first
                if (_application.ActiveWindow == null ||
                    _application.ActiveWindow.View == null ||
                    _application.ActiveWindow.View.Slide == null)
                {
                    _notificationCallback("Please open a presentation and navigate to a slide first.", true);
                    return;
                }

                SetThemeColor(Office.MsoThemeColorIndex.msoThemeColorAccent2);
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error setting secondary theme color: {ex.Message}", true);
            }
        }

        /// <summary>
        /// Gets a ColorFormat object for the specified theme color
        /// </summary>
        private PowerPoint.ColorFormat GetColorFromTheme(Office.MsoThemeColorIndex themeColor)
        {
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tempShape = null;
            PowerPoint.ColorFormat colorFormat = null;

            try
            {
                // Check if there's an active window and slide
                if (_application.ActiveWindow == null)
                {
                    _notificationCallback("No active presentation window.", false);
                    return null;
                }

                if (_application.ActiveWindow.View.Slide == null)
                {
                    _notificationCallback("No active slide.", false);
                    return null;
                }

                // Create a temporary shape to get a ColorFormat object
                slide = _application.ActiveWindow.View.Slide;
                tempShape = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    0, 0, 1, 1);

                // Set the color to specified theme color
                tempShape.TextFrame.TextRange.Font.Color.ObjectThemeColor = themeColor;

                // Get the ColorFormat - create a new reference that we'll return
                colorFormat = tempShape.TextFrame.TextRange.Font.Color;

                // Delete the temporary shape
                tempShape.Delete();
                _comObjectManager.ReleaseComObject(tempShape, "Temporary Shape");
                tempShape = null;

                return colorFormat;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error creating color: {ex.Message}", true);

                // Clean up created objects if there was an error
                if (colorFormat != null)
                {
                    _comObjectManager.ReleaseComObject(colorFormat, "Error ColorFormat");
                }

                return null;
            }
            finally
            {
                // Always clean up temporary objects
                if (tempShape != null)
                {
                    try { tempShape.Delete(); } catch { /* ignore deletion errors */ }
                    _comObjectManager.ReleaseComObject(tempShape, "Temp Shape");
                }

                if (slide != null)
                {
                    _comObjectManager.ReleaseComObject(slide, "Slide");
                }
            }
        }

        /// <summary>
        /// Gets default theme color as a PowerPoint ColorFormat (for API compatibility)
        /// </summary>
        /// <returns>ColorFormat object with the current theme color</returns>
        public PowerPoint.ColorFormat GetDefaultColor()
        {
            try
            {
                // If we already have a cached color, return it
                if (_changeToColor != null)
                {
                    return _changeToColor;
                }

                // Check if there's an active presentation
                if (_application.ActiveWindow == null)
                {
                    _notificationCallback("No active presentation window.", false);
                    return null;
                }

                if (_application.ActiveWindow.View.Slide == null)
                {
                    _notificationCallback("No active slide.", false);
                    return null;
                }

                // Create a new color from theme
                _changeToColor = GetColorFromTheme(_themeColorIndex);
                return _changeToColor;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error creating color: {ex.Message}", true);
                return null;
            }
        }

        /// <summary>
        /// Validates and gets shapes containing text from the current selection
        /// </summary>
        /// <returns>ShapeRange containing shapes with text, or null if invalid</returns>
        public PowerPoint.ShapeRange GetSelectedShapesWithText()
        {
            PowerPoint.Slide currentSlide = null;
            PowerPoint.Selection selection = null;
            PowerPoint.ShapeRange shapes = null;
            List<PowerPoint.Shape> shapesWithText = null;

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
                    _notificationCallback("Please select at least one shape containing text.", false);
                    return null;
                }

                // Get selected shapes
                shapes = selection.ShapeRange;

                // Check if we have at least one shape
                if (shapes.Count < 1)
                {
                    _notificationCallback("Please select at least one shape containing text.", false);
                    return null;
                }

                // Filter to shapes that have text frames
                shapesWithText = new List<PowerPoint.Shape>();
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    try
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                            shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            shapesWithText.Add(shape);
                        }

                        // Don't release individual shapes from the ShapeRange as we're 
                        // returning the entire ShapeRange to the caller
                    }
                    catch
                    {
                        // Skip shapes that throw errors when checking text properties
                    }
                }

                if (shapesWithText.Count == 0)
                {
                    _notificationCallback("None of the selected shapes contain text.", false);

                    // Release the shapes collection
                    if (shapes != null)
                    {
                        _comObjectManager.ReleaseComObject(shapes, "Empty ShapeRange");
                    }

                    return null;
                }

                // Return the original shape range - the caller is responsible for releasing it
                return shapes;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error accessing selected shapes: {ex.Message}", true);

                // Release shapes in case of error
                if (shapes != null)
                {
                    _comObjectManager.ReleaseComObject(shapes, "ShapeRange in Error");
                }

                return null;
            }
            finally
            {
                // Clean up temporary objects
                if (selection != null) _comObjectManager.ReleaseComObject(selection, "Selection");
                if (currentSlide != null) _comObjectManager.ReleaseComObject(currentSlide, "Slide");

                // Clear the list of shapes with text (we don't release these as they're part of the returned ShapeRange)
                if (shapesWithText != null)
                {
                    shapesWithText.Clear();
                }
            }
        }

        /// <summary>
        /// Applies the current color to all bold text in the selected shapes
        /// </summary>
        /// <param name="shapes">ShapeRange containing shapes to process</param>
        /// <returns>Number of text ranges that were colored</returns>
        public int ColorBoldText(PowerPoint.ShapeRange shapes)
        {
            try
            {
                // Use the current theme color
                Office.MsoThemeColorIndex themeColorToApply = _themeColorIndex;
                int totalTextRangesColored = 0;
                int shapesProcessed = 0;

                // Process each shape
                for (int shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                {
                    PowerPoint.Shape shape = shapes[shapeIndex];
                    PowerPoint.TextRange textRange = null;

                    try
                    {
                        // Skip shapes without text
                        if (shape.HasTextFrame != Office.MsoTriState.msoTrue ||
                            shape.TextFrame.HasText != Office.MsoTriState.msoTrue)
                        {
                            continue;
                        }

                        shapesProcessed++;
                        textRange = shape.TextFrame.TextRange;

                        // Process each character in the text
                        for (int i = 1; i <= textRange.Length; i++)
                        {
                            PowerPoint.TextRange charRange = null;

                            try
                            {
                                charRange = textRange.Characters(i, 1);

                                // Check if this character is bold
                                if (charRange.Font.Bold == Office.MsoTriState.msoTrue)
                                {
                                    // Apply theme color directly - no need for a temporary shape or ColorFormat
                                    charRange.Font.Color.ObjectThemeColor = themeColorToApply;
                                    totalTextRangesColored++;
                                }
                            }
                            finally
                            {
                                // Release each character range
                                if (charRange != null)
                                {
                                    _comObjectManager.ReleaseComObject(charRange, "CharRange");
                                }
                            }
                        }
                    }
                    finally
                    {
                        // Release the text range for this shape
                        if (textRange != null)
                        {
                            _comObjectManager.ReleaseComObject(textRange, "TextRange");
                        }

                        // Note: We don't release the shape itself as it's part of the ShapeRange
                        // that was passed in and should be managed by the caller
                    }
                }

                // Display appropriate message based on results
                if (totalTextRangesColored > 0)
                {
                    _notificationCallback($"Applied color to {totalTextRangesColored} bold text segments in {shapesProcessed} shapes.", false);
                }
                else if (shapesProcessed > 0)
                {
                    _notificationCallback("No bold text found in the selected shapes.", false);
                }
                else
                {
                    _notificationCallback("No text-containing shapes were selected.", false);
                }

                return totalTextRangesColored;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error coloring bold text: {ex.Message}", true);
                return 0;
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

        /// <summary>
        /// Properly releases all cached COM objects when the service is being shut down
        /// </summary>
        public void Cleanup()
        {
            // Release cached color format
            if (_changeToColor != null)
            {
                _comObjectManager.ReleaseComObject(_changeToColor, "Cached ColorFormat");
                _changeToColor = null;
            }
        }
    }
}
