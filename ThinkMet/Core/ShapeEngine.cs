using System;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ThinkMet.Core
{
    public static class ShapeEngine
    {
        public static void CreateStackedBarChart(PowerPoint.Slide slide, float startX, float startY, float barWidth, float[] values)
        {
            if (slide == null || values == null || values.Length == 0) return;

            float currentY = startY;
            float scaleFactor = 15f;

            for (int i = 0; i < values.Length; i++)
            {
                float val = values[i];
                float height = val * scaleFactor;
                currentY -= height;

                // 1. Crear rectángulo (Usando MsoAutoShapeType)
                PowerPoint.Shape rect = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    startX,
                    currentY,
                    barWidth,
                    height
                );

                // Relleno y borde (Usando rect.Line.ForeColor)
                rect.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(
                    i % 2 == 0 ? System.Drawing.Color.DodgerBlue : System.Drawing.Color.LightSkyBlue
                );
                rect.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                // 2. Agregar etiqueta de texto (Usando MsoTextOrientation)
                PowerPoint.Shape label = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    startX,
                    currentY + (height / 2f) - 10f,
                    barWidth,
                    20f
                );

                label.TextFrame.TextRange.Text = val.ToString("N0");
                label.TextFrame.TextRange.Font.Size = 10;
                label.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                label.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            }
        }
    }
}