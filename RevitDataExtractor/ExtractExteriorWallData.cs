using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using CsvHelper;
using System.Globalization;
using Autodesk.Revit.DB.Analysis;
using System.IO;

namespace RevitDataExtractor
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class ExtractExteriorWallData : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UI Document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get document
            Document doc = uidoc.Document;

            //Create BuildingEnvelopeAnalyzer
            BuildingEnvelopeAnalyzerOptions options = new BuildingEnvelopeAnalyzerOptions();
            options.AnalyzeEnclosedSpaceVolumes = false;
            options.OptimizeGridCellSize = true;
            BuildingEnvelopeAnalyzer analyzer = BuildingEnvelopeAnalyzer.Create(doc, options);

            //Get exterior walls from analyzer
            var walls = new List<Element>();
            var exteriorElementIds = analyzer.GetBoundingElements();
            foreach (var id in exteriorElementIds )
            {
                var wall = doc.GetElement(id.HostElementId);
                //Only gather if element is a wall type and thicker than 4". Thickness test removes finish walls from selection.
                //Get element type to extract type parameters from
                var elementTypeName = wall.GetType().Name;
                
                if (elementTypeName == "Wall")
                {
                    var wallType = wall as Wall;
                    var wallThickness = wallType.Width;
                    if (wallThickness > 0.25)
                    {
                        walls.Add(wall);
                    }
                }   
            }

            //Collect distinct wall family names
            var typeNames = new List<string>();
            foreach (var wallInstance in walls)
            {
                typeNames.Add(wallInstance.Name);
            }
            var distinctTypeNames = typeNames.Distinct();

            //Gather data for each unique family type
            List<WallCollector> wallsData = new List<WallCollector>();
            foreach (var typeName in distinctTypeNames)
            {
                //Get length by isolating all instances of current type and summing alength
                IEnumerable<Element> wallsOfType = walls.Where(x => x.Name == typeName);

                double length = 0;
                foreach (var wallEle in wallsOfType)
                {
                    length += wallEle.LookupParameter("Length").AsDouble();
                }
                var wall = wallsOfType.First();

                //Get lement type to extract type parameters from
                Element eleType = doc.GetElement(wall.GetTypeId());

                //Get parameter values from familySymbol or wall object
                string family = wall.LookupParameter("Family").AsValueString();
                string typeMark = eleType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsValueString();
                string description = eleType.LookupParameter("Description").AsValueString();
                string type = wall.LookupParameter("Type").AsValueString();

                //Create new wall object
                WallCollector wallObject = new WallCollector { Family = family, TypeMark = typeMark, Type = type, Length = length, Description = description };

                //Add object to array
                wallsData.Add(wallObject);
            }

            //Export Data to CSV
            var targetLocation = @"C:\Users\lelandcurtis\Documents";
            var fileName = "ExteriorWalls.csv";
            using (var writer = new StreamWriter(targetLocation + "\\" + fileName))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(wallsData);
            }

            //Print results to dialog box
            TaskDialog.Show("Exterior Walls", string.Format("{0} wall types exported to {1}", distinctTypeNames.Count(), (targetLocation + "\\" + fileName)));

            //Return success
            return Result.Succeeded;
        }
    }
}
