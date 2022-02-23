using System;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using CsvHelper;
using System.Globalization;

namespace RevitDataExtractor
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class ExtractFloorData :IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UI Document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get document
            Document doc = uidoc.Document;

            //Create filtered element collectors
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);

            //Create floor fixture filter
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Floors);

            //Apply filters to get floor element instances
            ICollection<Element> floors = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            //Collect distinct floor family names
            var typeNames = new List<string>();
            foreach (var floorInstance in floors)
            {
                typeNames.Add(floorInstance.Name);
            }
            var distinctTypeNames = typeNames.Distinct();

            //Gather data for each unique family type
            List<FloorCollector> floorsData = new List<FloorCollector>();
            foreach (var familyName in distinctTypeNames)
            {
                //Get area by isolating all instances of current family type and summing area
                IEnumerable<Element> floorsOfType = floors.Where(x => x.Name == familyName);
                
                double area = 0;
                foreach (var floorEle in floorsOfType)
                {
                    area += floorEle.LookupParameter("Area").AsDouble();
                }
                var floor = floorsOfType.First();

                //Get lement type to extract type parameters from
                Element eleType = doc.GetElement(floor.GetTypeId());

                //Get parameter values from familySymbol or floor object
                string family = floor.LookupParameter("Family").AsValueString();
                string typeMark = eleType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsValueString();
                string description = eleType.LookupParameter("Description").AsValueString();
                string type = floor.LookupParameter("Type").AsValueString();

                //Create new floor object
                FloorCollector floorObject = new FloorCollector { Family = family, TypeMark = typeMark, Type = type, Area = area , Description = description};

                //Add object to array
                floorsData.Add(floorObject);
            }

            //Export Data to CSV
            var targetLocation = @"C:\Users\lelandcurtis\Documents";
            var fileName = "Floors.csv";
            using (var writer = new StreamWriter(targetLocation + "\\" + fileName))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(floorsData);
            }

            //Print results to dialog box
            TaskDialog.Show("Floors", string.Format("{0} floor types exported to {1}", distinctTypeNames.Count(), (targetLocation + "\\" + fileName)));

            //Return success
            return Result.Succeeded;
        }
    }
}
