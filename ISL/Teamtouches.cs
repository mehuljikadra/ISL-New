using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using excel = Microsoft.Office.Interop.Excel;

namespace ISL
{
    class Teamtouches
    {
        private IWebDriver driver;
        private object missing;

        public Teamtouches(IWebDriver driver)
        {
            this.driver = driver;
        }
        public IWebElement GetTeamtouchesData()
        {


            var Team = new List<Team>();

            Thread.Sleep(2000);
            var allDivs = driver.FindElements(By.CssSelector(".si-team-data"));
            foreach (var div in allDivs)
            {
                var teamname = div.FindElement(By.CssSelector(".si-fullName")).Text;

                var teamTouches = div.FindElement(By.CssSelector(".si-goals")).Text;

                Team.Add(new Team { TeamName = teamname, TeamTouches = teamTouches });
                TestContext.Out.WriteLine($"TeamName: { teamname }| TeamTouches: { teamTouches } ");


            }
            CreateXlSheet(Team);
            return null;
        }

        public void CreateXlSheet(List<Team> Team)
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


                oSheet.Cells[1, 1] = "TeamName";

                oSheet.Cells[1, 2] = "TeamTouches";

                for (int i = 0; i < Team.Count; i++)
                {
                    int row = i + 1;
                    oSheet.Cells[row, 1] = Team[i].TeamName;

                    oSheet.Cells[row, 2] = Team[i].TeamTouches;
                }

                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\TeamTouches.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;

                string srcPath = (@"C:\Users\aditya.bhosle\Desktop\ISL\FTeamtouches.xls");

                oWB = (excel._Workbook)(oXL.Workbooks.Open(srcPath));
                oSheet = oWB.Worksheets.get_Item(1);

                string destPath = (@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\TeamTouches.xlsx");
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(destPath));
                dSheet = oWB.Worksheets.Add();

                excel.Range from = oSheet.Range["A:A,B:B"];
                excel.Range torange = dSheet.Range["A1:B1"];

                from.Copy(torange);


                oXL.ActiveSheet.Range["D2:D1000"] = "=VLOOKUP(A2,Sheet1!A:B,1,False)";
                oXL.ActiveSheet.Range["E2:E1000"] = "=VLOOKUP(A2,Sheet1!A:B,2,FALSE)";
                oXL.ActiveSheet.Range["F2:F1000"] = "=EXACT(D:D,A:A)";
                oXL.ActiveSheet.Range["G2:G1000"] = "=EXACT(E:E,B:B)";


                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\TeamTouches.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                   false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;
            }
            catch (Exception) { }
        }
    }
}
