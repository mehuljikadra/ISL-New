using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using excel = Microsoft.Office.Interop.Excel;

namespace ISL
{


    class CleanSheetStats
    {
        private IWebDriver driver;
        private object missing;

        public CleanSheetStats(IWebDriver driver)
        {
            this.driver = driver;
        }
        public IWebElement GetCleanSheetData()
        {



            var DropDown = driver.FindElement(By.CssSelector(".si-stats-dropdown-container"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(DropDown);
            action1.Perform();
            DropDown = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.CssSelector(".si-stats-dropdown-container")));


            var players = new List<players>
            {
                Firstplayer()
            };
            var allDivs = driver.FindElements(By.CssSelector(".si-tRow ")).Skip(2);
            foreach (var div in allDivs)
            {
                var name = div.FindElement(By.CssSelector(".si-fullName ")).Text;
                var id = div.GetAttribute("data-playerid");

                var cleansheet1 = div.FindElement(By.CssSelector(".si-plyStats-gamplyd span")).Text;

                players.Add(new players { Name = name, Id = id, CleanSheets = cleansheet1 });
                TestContext.Out.WriteLine($"NAME: { name }  | Id:{id}| CleanSheets: { cleansheet1 }");
            }


            CreateXlSheet(players);
            return null;
        }

        public void CreateXlSheet(List<players> players)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel._Worksheet dSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                //oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 1] = "Name";
                oSheet.Cells[1, 2] = "Id";

                oSheet.Cells[1, 3] = "CleanSheets";

                for (int i = 0; i < players.Count; i++)
                {
                    int row = i + 1;
                    oSheet.Cells[row, 1] = players[i].Name;
                    oSheet.Cells[row, 2] = players[i].Id;

                    oSheet.Cells[row, 3] = players[i].CleanSheets;
                }

                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\Cleansheets.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;

                string srcPath = (@"C:\Users\aditya.bhosle\Desktop\ISL\Cleansheet.xls");

                oWB = (excel._Workbook)(oXL.Workbooks.Open(srcPath));
                oSheet = oWB.Worksheets.get_Item(1);

                string destPath = (@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\Cleansheets.xlsx");
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(destPath));
                dSheet = oWB.Worksheets.Add();

                excel.Range from = oSheet.Range["A:A,B:B,C:C"];
                excel.Range torange = dSheet.Range["A1:B1:C1"];

                from.Copy(torange);


                oXL.ActiveSheet.Range["D2:D1000"] = "=VLOOKUP(B2,Sheet1!B:C,1,False)";
                oXL.ActiveSheet.Range["E2:E1000"] = "=VLOOKUP(B2,Sheet1!B:C,2,FALSE)";
                oXL.ActiveSheet.Range["F2:F1000"] = "=EXACT(D:D,B:B)";
                oXL.ActiveSheet.Range["G2:G1000"] = "=EXACT(E:E,C:C)";


                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\Cleansheets.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;
            }
            catch (Exception) { }
        }
        public players Firstplayer()

        {
           

            var name = driver.FindElement(By.CssSelector(".si-awdPlyrName")).Text;
            var id = driver.FindElement(By.CssSelector(".si-statHeadRow")).GetAttribute("data-playerid");

            var goals = driver.FindElement(By.CssSelector(".si-points span")).Text;

            return new players
            {

                Name = name,
                Id = id,
                Goals = goals

            };

        }
    }
}
