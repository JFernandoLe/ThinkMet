using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ThinkMet.Core;

namespace ThinkMet
{
    public partial class MainRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

                if (pptApp.ActiveWindow != null && pptApp.ActiveWindow.View.Slide != null)
                {
                    PowerPoint.Slide activeSlide = (PowerPoint.Slide)pptApp.ActiveWindow.View.Slide;
                    float[] datosEjemplo = new float[] { 10f, 15f, 8f, 20f };

                    ShapeEngine.CreateStackedBarChart(activeSlide, 200f, 400f, 80f, datosEjemplo);
                }
                else
                {
                    MessageBox.Show("Asegúrate de estar en la vista de diapositiva activa.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al generar el gráfico: " + ex.Message);
            }
        }
    }
}