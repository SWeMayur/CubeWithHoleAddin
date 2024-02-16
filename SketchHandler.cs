using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvAddIn
{
    internal class SketchHandler
    {
        private Application _inventorApplication;

        public SketchHandler(Inventor.Application inventorApplication)
        {
            _inventorApplication = inventorApplication;
        }

        public PlanarSketch CreateSketch(PartDocument partDoc)
        {
            try
            {
                PartComponentDefinition compDef = partDoc.ComponentDefinition;
                return compDef.Sketches.Add(compDef.WorkPlanes[3]);
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(new Exception($"Error creating a sketch: {ex.Message}"));
                return null;
            }
        }

        public void DrawRectangle(PlanarSketch sketch)
        {
            try
            {
                Point2d startPoint = _inventorApplication.TransientGeometry.CreatePoint2d(0, 0);
                Point2d endPoint = _inventorApplication.TransientGeometry.CreatePoint2d(10, 10);
                sketch.SketchLines.AddAsTwoPointRectangle(startPoint, endPoint);
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(new Exception($"Error drawing a rectangle: {ex.Message}"));
            }
        }

        public void DrawCircle(PlanarSketch sketch)
        {
            try
            {
                Point2d centerPoint = _inventorApplication.TransientGeometry.CreatePoint2d(5, 5);
                sketch.SketchCircles.AddByCenterRadius(centerPoint, 2);
            }
            catch (Exception ex)
            {
                ExceptionHandler.HandleException(new Exception($"Error drawing a circle: {ex.Message}"));
            }
        }

    }
}
