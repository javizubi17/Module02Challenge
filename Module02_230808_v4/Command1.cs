#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace Module02_230808_v4
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        internal WallType GetWallTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc);
            wallCollector.OfClass(typeof(WallType));

            foreach (WallType wallType in wallCollector)
            {
                if (wallType.Name == typeName)
                {
                    return wallType;
                }
            }
            return null;
        }
        internal DuctType GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector ductCollector = new FilteredElementCollector(doc);
            ductCollector.OfClass(typeof(DuctType));

            foreach (DuctType ductType in ductCollector)
            {
                if (ductType.Name == typeName)
                {
                    return ductType;
                }
            }
            return null;
        }
        internal Element GetMEPSystemByName(Document doc, string typeName)
        {
            FilteredElementCollector ductCollector = new FilteredElementCollector(doc);
            ductCollector.OfClass(typeof(MEPSystemType));

            foreach (Element ductType in ductCollector)
            {
                if (ductType.Name == typeName)
                {
                    return ductType;
                }
            }
            return null;
        }
        internal PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector pipeCollector = new FilteredElementCollector(doc);
            pipeCollector.OfClass(typeof(PipeType));

            foreach (PipeType pipeType in pipeCollector)
            {
                if (pipeType.Name == typeName)
                {
                    return pipeType;
                }
            }
            return null;
        }
        internal Level GetLevelByName(Document doc, string name)
        {
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            levelCollector.OfClass(typeof(Level));
            levelCollector.WhereElementIsNotElementType();

            foreach (Level levelType in levelCollector)
            {
                if (levelType.Name == name)
                {
                    return levelType;
                }
            }
            return null;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. pick elements and filter them into list
            UIDocument uidoc = uiapp.ActiveUIDocument;
            TaskDialog.Show("Select Lines", "Select some line to convert to Revit Elements.");
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select Elements");

            //2b. filter selected elements for model curves
            List<CurveElement> modelCurves = new List<CurveElement>();
            foreach (Element elem in pickList)
            {
                if (elem is CurveElement)
                {
                    CurveElement curvElem = elem as CurveElement;
                    if (curvElem.CurveElementType == CurveElementType.ModelCurve)
                    {
                        modelCurves.Add(curvElem);
                    }
                }
            }

            //TaskDialog.Show("Curves", "You selected" + modelCurves.Count).ToString() + "lines.");
            TaskDialog.Show("Curves", $"You selected {modelCurves.Count} lines.");

            // Get level and various types
            Level myLevel = GetLevelByName(doc, "Level 1");
            
            // Get Types
            Element wallType1 = GetWallTypeByName(doc, "Storefront");
            Element wallType2 = GetWallTypeByName(doc, "Generic - 8\" Masonry");

            Element ductSystemType = GetMEPSystemByName(doc, "Supply Air");
            Element ductType = GetDuctTypeByName(doc, "Default");

            Element pipeSystemType = GetMEPSystemByName(doc, "Domestic Hot Water");
            Element pipeType = GetPipeTypeByName(doc, "Default");

            List<ElementId> linesToHide = new List<ElementId>();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Elements");

                foreach (CurveElement curveCurveElem in modelCurves)
                {
                    Curve elementCurve = curveCurveElem.GeometryCurve;
                    GraphicsStyle curveCurveElemGS = curveCurveElem.LineStyle as GraphicsStyle;
                        
                    if (elementCurve.IsBound == false)
                        continue; //adding continue here means if it's false, it will skip over it as what we want is curves where IsBound == True

                    XYZ curve1Start = elementCurve.GetEndPoint(0);
                    XYZ curve1End = elementCurve.GetEndPoint(1);

                    switch (curveCurveElemGS.Name)
                    {
                        case "A-GLAZ":
                            Wall currentWall = Wall.Create(doc, elementCurve, wallType1.Id, myLevel.Id, 20, 0, false, false);
                            break;
                        case "A-WALL":
                            Wall currentWall2 = Wall.Create(doc, elementCurve, wallType2.Id, myLevel.Id, 20, 0, false, false);
                            break;
                        case "M-DUCT":
                            Duct currentDuct = Duct.Create(doc, ductSystemType.Id, ductType.Id, myLevel.Id, curve1Start, curve1End);
                            break;
                        case "P-PIPE":
                            Pipe currentPipe = Pipe.Create(doc, pipeSystemType.Id, pipeType.Id, myLevel.Id, curve1Start, curve1End);
                            break;
                        default:
                            linesToHide.Add(curveCurveElem.Id);
                            break;

                    }
                }
                t.Commit();
            }
            using (Transaction t2 = new Transaction(doc))
            {
                List<ElementId> linesToHide2 = new List<ElementId>();
                foreach (Element linesH in pickList)
                {
                    ElementId linesHId = linesH.Id;
                    linesToHide2.Add(linesHId);
                }
                    
                t2.Start("Hide Lines in View");
                View active_View = doc.ActiveView;
                active_View.HideElements(linesToHide2);
                t2.Commit();
            }
            return Result.Succeeded;
        }
            internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
