using CubeWithHoleAddin;
using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvAddIn
{
    internal class DocumentHandler
    {
        private Inventor.Application _inventorApplication;

        public DocumentHandler(Inventor.Application inventorApplication)
        {
            _inventorApplication = inventorApplication;
        }

        public PartDocument CreateNewPartDocument()
        {
            try
            {
                return _inventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, _inventorApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)) as PartDocument;
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(new Exception($"Error creating a new part document: {ex.Message}"));
                return null;
            }
        }

        public void ExtrudeSketch(PartDocument partDoc)
        {
            try
            {
                PartComponentDefinition compDef = partDoc.ComponentDefinition;
                PlanarSketch sketch = compDef.Sketches[1];
                Profile profile = sketch.Profiles.AddForSolid();

                ExtrudeDefinition extrudeDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(profile, PartFeatureOperationEnum.kNewBodyOperation);
                extrudeDef.SetDistanceExtent(10, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

                compDef.Features.ExtrudeFeatures.Add(extrudeDef);
                partDoc.Update();
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(new Exception($"Error extruding the sketch: {ex.Message}"));
            }
        }

        public void SavePartDocument(PartDocument partDoc)
        {
            try
            {
                partDoc.SaveAs(@"D:\\Mayur_Workspace\\Incubation\\InventorAddIn1\\output\\NewPart.ipt", false);
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(new Exception($"Error saving the part document: {ex.Message}"));
            }
        }

        public static void ShowOneViewAndExportAsJpg(PartDocument partDoc)
        {
            // Assuming you have a part document loaded or created
            if (partDoc != null)
            {
                // Create a drawing document
                DrawingDocument drawingDoc = StandardAddInServer.m_inventorApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject) as DrawingDocument;

                // Create a new sheet in the drawing
                Sheet sheet = drawingDoc.Sheets.Add();

                // Create a base view of the part on the sheet
                DrawingView baseView = sheet.DrawingViews.AddBaseView(
                    partDoc as _Document,
                    StandardAddInServer.m_inventorApplication.TransientGeometry.CreatePoint2d(10, 10),
                    2.0,
                    ViewOrientationTypeEnum.kIsoTopRightViewOrientation,
                    DrawingViewStyleEnum.kHiddenLineDrawingViewStyle,//gives perfect drawing
                    "", // Provide the name of the view, or leave it empty
                    null, // Use default scale settings
                    null // Use default position settings
                );


                // Adjust the view scale and position if needed
                baseView.Scale = 1.0;
                baseView.Position = StandardAddInServer.m_inventorApplication.TransientGeometry.CreatePoint2d(10, 10);

                Inventor.View view = StandardAddInServer.m_inventorApplication.ActiveView;

                view.SaveAsBitmap(@"D:\Mayur_Workspace\Incubation\DwgOutput\DrawingImage.jpg", 0, 0);

                // Export the drawing view as an image (JPG)
                //ExportOptionsBitmap bmpOptions = StandardAddInServer.m_inventorApplication.TransientObjects.CreateExportOptionsBitmap();
                //bmpOptions.PixelFormat = PixelFormatEnum.k24Bit;
                //bmpOptions.Transparency = true;
                //bmpOptions.Resolution = PrintResolutionEnum.kScreenResolution;

                //string jpgFilePath = @"D:\Mayur_Workspace\Incubation\InventorAddIn1\output\DrawingImage.jpg";
                //baseView.ExportBitmap(jpgFilePath, bmpOptions);

                // Optional: Close the drawing document if not needed
                //  drawingDoc.Close(true);
            }
        }
    }
}
