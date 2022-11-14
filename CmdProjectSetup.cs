#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using static Autodesk.Revit.DB.SpecTypeId;

#endregion

namespace RevitAddin_Challenge02_LevelsAndSheet
{
    [Transaction(TransactionMode.Manual)]
    public class CmdProjectSetup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //code starts
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            //variables
            string filepathLevels = ("C:\\Users\\GARY\\Downloads\\RAB_Session_02_Challenge_Levels.csv");
            string filePathSheets = ("C:\\Users\\GARY\\Downloads\\RAB_Session_02_Challenge_Sheets.csv");

            //pick titleblock
            ElementId tblockId = collector.FirstElementId();

            Transaction t = new Transaction(doc);
            t.Start("Create level and sheet");

            //read data

            string[] fileLevelArray = System.IO.File.ReadAllLines(filepathLevels);
            foreach(string rowString in fileLevelArray)
            {
                string[] cellString = rowString.Split(',');
                string LevelName = cellString[0];
                string elevation = cellString[1];

                double elevationDouble = 0;
                bool DidItParse = double.TryParse(elevation,out elevationDouble);

                //create levels
                double LevelHeight = elevationDouble;
                Level mylevel = Level.Create(doc, LevelHeight);
                mylevel.Name = LevelName;
            }
            
            string[] fileSheetArray = System.IO.File.ReadAllLines(filePathSheets);
            foreach (string rowString in fileSheetArray)
            {
                string[] cellString = rowString.Split(',');
                string SheetNumber = cellString[0];
                string SheetName = cellString[1];
                
                //create sheets

                ViewSheet MySheet = ViewSheet.Create(doc, tblockId);
                MySheet.Name = SheetName;
                MySheet.SheetNumber = SheetNumber;
            }

            var levelcount = 0;
            levelcount = fileLevelArray.Count();

            var Sheetcount = 0;
            Sheetcount = fileSheetArray.Count();

            TaskDialog.Show("app", levelcount + " levels & "+ Sheetcount +" sheets created");
            t.Commit();
            t.Dispose();



            return Result.Succeeded;
        }
    }
}
