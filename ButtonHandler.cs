using CubeWithHoleAddin;
using Inventor;
using System;
using System.Windows.Forms;

namespace InvAddIn
{
    internal class ButtonHandler
    {
        public static void OnExecute(NameValueMap context)
        {
            try
            {
                MessageBox.Show("Creating cube with hole");

                var documentHandler = new DocumentHandler(StandardAddInServer.m_inventorApplication);
                var sketchHandler = new SketchHandler(StandardAddInServer.m_inventorApplication);

                PartDocument partDoc = documentHandler.CreateNewPartDocument();
                PlanarSketch sketch = sketchHandler.CreateSketch(partDoc);

                sketchHandler.DrawRectangle(sketch);
                sketchHandler.DrawCircle(sketch);
                documentHandler.ExtrudeSketch(partDoc);
                documentHandler.SavePartDocument(partDoc);

                DocumentHandler.ShowOneViewAndExportAsJpg(partDoc);

                Console.WriteLine("Cube with hole created successfully!");
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(ex);
            }
        }
        
    }
}
