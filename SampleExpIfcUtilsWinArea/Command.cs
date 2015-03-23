#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using I = Autodesk.Revit.DB.IFC;
#endregion

namespace SampleExpIfcUtilsWinArea
{
  [Transaction( TransactionMode.ReadOnly )]
  public class Command : IExternalCommand
  {
    const double _square_feet_to_square_metres
      = 0.09290304;

    /// <summary>
    /// Return surface area of given family instance
    /// in square metres.
    /// </summary>
    static double GetInstanceSurfaceAreaMetric(
      FamilyInstance familyInstance )
    {
      double area_sq_ft = 0;

      Wall wall = familyInstance.Host as Wall;

      if( null != wall )
      {
        if( wall.WallType.Kind == WallKind.Curtain )
        {
          area_sq_ft = familyInstance.get_Parameter(
            BuiltInParameter.HOST_AREA_COMPUTED )
              .AsDouble();
        }
        else
        {
          Document doc = familyInstance.Document;
          XYZ basisY = XYZ.BasisY;

          // using I = Autodesk.Revit.DB.IFC;

          CurveLoop curveLoop = I.ExporterIFCUtils
            .GetInstanceCutoutFromWall( doc, wall,
              familyInstance, out basisY );

          IList<CurveLoop> loops
            = new List<CurveLoop>( 1 );

          loops.Add( curveLoop );

          area_sq_ft = I.ExporterIFCUtils
            .ComputeAreaOfCurveLoops( loops );
        }
      }
      else
      {
        double width
          = familyInstance.Symbol.get_Parameter(
            BuiltInParameter.FAMILY_WIDTH_PARAM )
              .AsDouble();

        double height
          = familyInstance.Symbol.get_Parameter(
            BuiltInParameter.FAMILY_HEIGHT_PARAM )
              .AsDouble();

        area_sq_ft = width * height;
      }
      return _square_feet_to_square_metres * area_sq_ft;
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      Selection sel = uidoc.Selection;

      StringBuilder sb = new StringBuilder();
      double areaTotal = 0;

      IEnumerable<ElementId> elementIds = sel.GetElementIds();

      foreach( ElementId elementId in elementIds )
      {
        FamilyInstance fi = doc.GetElement( elementId )
          as FamilyInstance;

        if( null != fi )
        {
          double areaMetric =
              GetInstanceSurfaceAreaMetric( fi );
          areaTotal += areaMetric;

          double areaRound = Math.Round( areaMetric, 2 );

          sb.AppendLine();
          sb.Append( "ElementId: " + fi.Id.IntegerValue );
          sb.Append( "  Name: " + fi.Name );
          sb.AppendLine( "  Area: " + areaRound + " m2" );
        }
      }
      int count = elementIds.Count<ElementId>();

      double areaPrintFriendly = Math.Round( areaTotal, 2 );

      sb.AppendLine( "\nTotal area: "
        + areaPrintFriendly + " m2" );

      TaskDialog taskDialog = new TaskDialog(
        "Selection Area" );

      taskDialog.MainInstruction = "Elements selected: "
        + count;

      taskDialog.MainContent = sb.ToString();

      taskDialog.Show();

      return Result.Succeeded;
    }
  }
}
