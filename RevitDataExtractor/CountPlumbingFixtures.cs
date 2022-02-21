using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace RevitDataExtractor
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CountPlumbingFixtures : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                //Get UI Document
                UIDocument uidoc = commandData.Application.ActiveUIDocument;

                //Get document
                Document doc = uidoc.Document;

                //Create filtered element collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                //Create filter
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_PlumbingFixtures);

                //Apply filter to get elements
                IList<Element> plumbingFixtures = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

                //Print results to dialog box
                TaskDialog.Show("Plumbing Fixtures", string.Format("There are {0} plumbing fixtures in the model", plumbingFixtures.Count));

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
        }
    }
}
