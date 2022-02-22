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
    public class ExtractPlumbingFixtureData : IExternalCommand
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

            //Create plumbing fixture filter
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_PlumbingFixtures);

            //Apply filters to get plumbing fixture element instances
            ICollection<Element> plumbingFixtures = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            //Filter Collector to get plumbing fixtures family symbols
            var families = collector2.WherePasses(filter).WhereElementIsElementType().OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();

            //Collect distinct plumbing fixture family names
            var familyNames = new List<string>();
            foreach (var family in families)
            {
                familyNames.Add(family.FamilyName);
            }
            var distinctFamilyNames = familyNames.Distinct();

            //Gather data for each unique family type
            List<PlumbingFixture> plumbingFixturesData = new List<PlumbingFixture>();
            foreach (var familyName in distinctFamilyNames)
            {
                //Get count by isolating all instances of current family type
                IEnumerable<Element> fixtures = plumbingFixtures.Where(x => x.LookupParameter("Family").AsValueString() == familyName);
                var fixture = fixtures.First();
                int count = fixtures.Count();

                //Get familySymbol object that matches unique family name
                var familySymbols = families.Where(x => x.FamilyName == familyName);
                var familySymbol = familySymbols.First();

                //Get lement type to extract type parameters from
                Element eleType = doc.GetElement(fixture.GetTypeId());
                
                //Get parameter values from familySymbol
                string family = familySymbol.FamilyName;
                string typeMark = eleType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsValueString();
                string description = eleType.LookupParameter("Description").AsValueString();
                string type = fixture.LookupParameter("Type").AsValueString();
                string manufacturer = eleType.LookupParameter("Manufacturer").AsValueString();

                //Create new Plumbing Fixture object
                PlumbingFixture plumbingFixture = new PlumbingFixture { Family = family, TypeMark = typeMark, Description = description, Type = type, Manufacturer = manufacturer, Count = count };

                //Add hash to array
                plumbingFixturesData.Add(plumbingFixture);

            }

            //Export Data to CSV
            var targetLocation = @"C:\Users\lelandcurtis\Documents";
            var fileName = "Plumbing_Fixtures.csv";
            using (var writer = new StreamWriter(targetLocation + "\\" + fileName))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(plumbingFixturesData);
            }

            //Print results to dialog box
            TaskDialog.Show("Plumbing Fixtures", string.Format("{0} plumbing fixture families exported to {1}", distinctFamilyNames.Count(), (targetLocation + "\\" + fileName)));

            //Return success
            return Result.Succeeded;
            
        }
    }
}
