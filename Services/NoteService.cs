using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ShapeMaster.Services
{
    public class NoteService
    {
        private readonly Application _application;
        private readonly Action<string, bool> _notificationCallback;

        public NoteService(Application application, Action<string, bool> notificationCallback)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notificationCallback = notificationCallback ?? throw new ArgumentNullException(nameof(notificationCallback));
        }

        /// <summary>
        /// Creates a note shape on the active slide with the specified color and label.
        /// </summary>
        public void CreateNoteShape(string colorString, string label)
        {
            try
            {
                if (_application.ActiveWindow == null || _application.ActiveWindow.View == null || _application.ActiveWindow.View.Slide == null)
                {
                    _notificationCallback("Please open a presentation and navigate to a slide first.", true);
                    return;
                }

                var slide = _application.ActiveWindow.View.Slide;
                float slideWidth = _application.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = _application.ActivePresentation.PageSetup.SlideHeight;
                float left = slideWidth * 0.10f;
                float top = slideHeight * 0.10f;
                float width = slideWidth * 0.25f;
                float fontSize = 10f;
                float paddingTop = 2f;
                float paddingBottom = 2f;
                float height = fontSize + paddingTop + paddingBottom;

                var shape = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeSnip1Rectangle,
                    left, top, width, height);

                shape.Name = label;
                shape.TextFrame.TextRange.Text = label;
                shape.TextFrame.TextRange.Font.Name = "Arial";
                shape.TextFrame.TextRange.Font.Size = fontSize;
                shape.TextFrame.TextRange.Font.Color.RGB = 0x000000;
                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Font.Underline = Office.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Font.Shadow = Office.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.Font.Emboss = Office.MsoTriState.msoFalse;

                // Use TextFrame2 to force solid black and clear possible WordArt effects
                shape.TextFrame2.TextRange.Font.Fill.Visible = Office.MsoTriState.msoTrue;
                shape.TextFrame2.TextRange.Font.Fill.Solid();
                shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x000000;
                try { shape.TextFrame2.TextRange.Font.Glow.Radius = 0; } catch { }
                try { shape.TextFrame2.TextRange.Font.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone; } catch { }

                // Parse color string (hex or name)
                int fillColor = ParseColorString(colorString);
                shape.Fill.ForeColor.RGB = fillColor;

                shape.Line.Visible = Office.MsoTriState.msoFalse;
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                shape.TextFrame2.MarginLeft = 2;
                shape.TextFrame2.MarginTop = paddingTop;
                shape.TextFrame2.MarginRight = 2;
                shape.TextFrame2.MarginBottom = paddingBottom;
                shape.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            }
            catch (Exception ex)
            {
                _notificationCallback($"Error inserting note: {ex.Message}", true);
            }
        }

        private int ParseColorString(string colorString)
        {
            if (string.IsNullOrWhiteSpace(colorString))
                return 0xC0C0C0; // Default: gray
            colorString = colorString.Trim();
            if (colorString.StartsWith("#"))
                colorString = colorString.Substring(1);
            if (int.TryParse(colorString, System.Globalization.NumberStyles.HexNumber, null, out int hex))
                return hex;
            // Add more color name parsing if needed
            return 0xC0C0C0; // Default: gray
        }
    }
}
