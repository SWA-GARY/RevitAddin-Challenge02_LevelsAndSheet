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
using Forms = System.Windows.Forms;
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
            string filepathLevels = ("");
            string filePathSheets = ("");
            XYZ insertPoint = new XYZ(2, 1, 0);
            XYZ secondInsertPoint = new XYZ(0, 1, 0);


            //pick titleblock
            ElementId tblockId = collector.FirstElementId();

            Transaction t = new Transaction(doc);
            t.Start("Create level and sheet");

            //dialog - setup
            TaskDialog.Show("app", "load levels file");

            Forms.OpenFileDialog selectlevelsfile = new Forms.OpenFileDialog();
            selectlevelsfile.InitialDirectory = "c:\\";
            selectlevelsfile.Filter = "CSV FILES|*.CSV|ALL FILES|*.*";
            selectlevelsfile.Multiselect = false;

            //open levels file
           
            
            if (selectlevelsfile.ShowDialog() == Forms.DialogResult.OK)
            {
                filepathLevels = selectlevelsfile.FileName;

            }

            //dialog - setup
            TaskDialog.Show("app", "load sheets file");

            Forms.OpenFileDialog selectsheetsfile = new Forms.OpenFileDialog();
            selectsheetsfile.InitialDirectory = "c:\\";
            selectsheetsfile.Filter = "CSV FILES|*.CSV|ALL FILES|*.*";
            selectsheetsfile.Multiselect = false;

            /*open sheets file
            if (selectsheetsfile.ShowDialog() == Forms.DialogResult.OK)
            {
                filePathSheets = selectsheetsfile.FileName;

            }
            */
            //read data
            List <string[]> LevelList= new List<string[]>();
            List<string[]> SheetsList = new List<string[]>();

            string[] fileLevelArray = System.IO.File.ReadAllLines(filepathLevels);
            
            foreach(string rowString in fileLevelArray)
            {
               //split cells
                string[] cellString = rowString.Split(',');
                //add to list
                LevelList.Add(cellString); 
            }
            //remove headers
            LevelList.RemoveAt(0);
            

            //create rcp and plan views
            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType planvft = null;
            ViewFamilyType rcpvft = null;

            foreach (ViewFamilyType vft in vftCollector)
            {
                if (vft.ViewFamily == ViewFamily.FloorPlan)
                    planvft = vft;

                if (vft.ViewFamily == ViewFamily.CeilingPlan)
                    rcpvft = vft;
            }

            //create levels
            int FLOORNUMBER = 0;
            int plannumber = 100;
            int RCPnumber = 200;


            foreach (string[] rowString in LevelList)
            {
                string LevelName = rowString[0];
                string elevation = rowString[1];

                double elevationDouble = 0;
                bool DidItParse = double.TryParse(elevation, out elevationDouble);

                double LevelHeight = elevationDouble;
                Level mylevel = Level.Create(doc, LevelHeight);
                mylevel.Name = LevelName;

                ViewPlan planview = ViewPlan.Create(doc, planvft.Id, mylevel.Id);
                FLOORNUMBER++;
                planview.Name = (FLOORNUMBER.ToString() + " FLOOR PLAN");

                ViewSheet newSheet = ViewSheet.Create(doc, tblockId);
                newSheet.Name = planview.Name;
                //plannumber++;
                newSheet.SheetNumber = ("A"+plannumber++.ToString());
               
                Viewport newViewPort = Viewport.Create(doc, newSheet.Id, planview.Id, insertPoint);

                ViewPlan RCPview = ViewPlan.Create(doc, rcpvft.Id,mylevel.Id);
                RCPview.Name= "RCP_"+LevelName;

                ViewSheet newCeilingSheet = ViewSheet.Create(doc, tblockId);
                newCeilingSheet.Name = "RCP_" + LevelName;
                newCeilingSheet.SheetNumber = "A"+RCPnumber++.ToString();

                Viewport newrcpPort = Viewport.Create(doc, newCeilingSheet.Id, RCPview.Id, insertPoint);






            }
                
            /*setup sheets list
            string[] fileSheetArray = System.IO.File.ReadAllLines(filePathSheets);
            foreach (string rowString in fileSheetArray)
            {
                string[] cellString = rowString.Split(',');
                SheetsList.Add(cellString);
            }

            //remove headers
            SheetsList.RemoveAt(0);

            foreach (string[] rowstring in SheetsList)
            {
                string SheetNumber = rowstring[0];
                string SheetName = rowstring[1];

                //create sheets

                ViewSheet MySheet = ViewSheet.Create(doc, tblockId);
                MySheet.Name = SheetName;
                MySheet.SheetNumber = SheetNumber;
                
            }
            */
            
            
            

            var levelcount = 0;
            levelcount = fileLevelArray.Count()*2;

           // var Sheetcount = 0;
           // Sheetcount = fileSheetArray.Count();

            TaskDialog.Show("app", levelcount + " levels & sheets created");
            t.Commit();
            t.Dispose();



            return Result.Succeeded;
        }
    }
}
