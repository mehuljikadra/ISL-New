using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using NUnit.Framework;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace ISL
{
    class GenericTracker
    {
        public IWebDriver driver { get; private set; }
        public string BrowserDataFilePath { get; private set; }

        private string SrcDataFilePath;

        public GenericTracker(IWebDriver driver, string BrowserDataFilePath, string SrcDataFilePath)
        {
            this.driver = driver;
            this.BrowserDataFilePath = BrowserDataFilePath;
            this.SrcDataFilePath = SrcDataFilePath;
        }

        public IWebElement GetData()
        {

            driver.FindElements(By.CssSelector(".si-ftk-opt2"));

            var goals1 = driver.FindElement(By.CssSelector(".si-fkt-bOne .si-fkt-sctn-number")).Text;
            var matches = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col1 .si-fkt-sctn-number")).Text;
            var avggoalsmatch = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col2 .si-fkt-sctn-number")).Text;
            var assist = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col3 .si-fkt-sctn-number")).Text;
            var touches = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col4 .si-fkt-sctn-number")).Text;
            var redcards = driver.FindElement(By.CssSelector(".si-fkt-bThree .si-fkt-sctn-number")).Text;
            var yellowcards = driver.FindElement(By.CssSelector(".si-fkt-bFour .si-fkt-sctn-number")).Text;
            var foulscommited = driver.FindElement(By.CssSelector(".si-fkt-teamA .fkt-nums")).Text;
            var foulssuffered = driver.FindElement(By.CssSelector(".si-fkt-teamB .fkt-nums")).Text;
            var clearences = driver.FindElement(By.CssSelector(".si-fkt-bSix .si-fkt-sctn-number")).Text;
            var tackles = driver.FindElement(By.CssSelector(".si-fkt-bSeven .si-fkt-sctn-number")).Text;
            var cleansheets = driver.FindElement(By.CssSelector(".si-fkt-bEight .si-fkt-sctn-number")).Text;
            var goalconversationrate = driver.FindElement(By.CssSelector(".si-fkt-bTen .fkt-nums")).Text;
            var shotsontarget = driver.FindElement(By.CssSelector(".si-fkt-bEleven .si-fkt-sctn-number")).Text;

            var leaguetracker = new Leaguetracker
            {
                Goals1 = goals1,
                Matches = matches,
                AvgGoalsMatch = avggoalsmatch,
                Assist = assist,
                Touches = touches,

                RedCards = redcards,
                YellowCards = yellowcards,
                Foulscommited = foulscommited,
                Foulssuffered = foulssuffered,
                Clearences = clearences,
                Tackles = tackles,
                CleanSheets = cleansheets,
                Goalconversationrate = goalconversationrate,
                Shotsontarget = shotsontarget
            };
            TestContext.Out.WriteLine($"Goals1: { goals1 } | Matches: { matches } | AvgGoalsMatch: { avggoalsmatch }" +
                $"| Assist: { assist } | Touches: { touches } | RedCards: { redcards } | YellowCards: { yellowcards }" +
                $"| Foulscommited: { foulscommited } | Foulssuffered: { foulssuffered } | Clearences: { clearences } | Tackles: { tackles }" +
                $" | CleanSheets: { cleansheets } | Goalconversationrate: { goalconversationrate} | Shotsontarget:  {shotsontarget}");



            CreateXlSheet(leaguetracker);
            return null;
        }

        public void CreateXlSheet(Leaguetracker leaguetracker)
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
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 1] = "Goals1";
                oSheet.Cells[1, 2] = "Matches";
                oSheet.Cells[1, 3] = "AvgGoalsMatch";

                oSheet.Cells[1, 4] = "Assist";
                oSheet.Cells[1, 5] = "Touches";
                oSheet.Cells[1, 6] = "RedCards";

                oSheet.Cells[1, 7] = "YellowCards";
                oSheet.Cells[1, 8] = "Foulscommited";
                oSheet.Cells[1, 9] = "Foulssuffered";


                oSheet.Cells[1, 10] = "Clearences";
                oSheet.Cells[1, 11] = "Tackles";
                oSheet.Cells[1, 12] = "CleanSheets";
                oSheet.Cells[1, 13] = "Goalconversationrate";
                oSheet.Cells[1, 14] = "Shotsontarget";



                oSheet.Cells[2, 1] = leaguetracker.Goals1;
                oSheet.Cells[2, 2] = leaguetracker.Matches;

                oSheet.Cells[2, 3] = leaguetracker.AvgGoalsMatch;
                oSheet.Cells[2, 4] = leaguetracker.Assist;
                oSheet.Cells[2, 5] = leaguetracker.Touches;
                oSheet.Cells[2, 6] = leaguetracker.RedCards;
                oSheet.Cells[2, 7] = leaguetracker.YellowCards;
                oSheet.Cells[2, 8] = leaguetracker.Foulscommited;
                oSheet.Cells[2, 9] = leaguetracker.Foulssuffered;
                oSheet.Cells[2, 10] = leaguetracker.Clearences;
                oSheet.Cells[2, 11] = leaguetracker.Tackles;
                oSheet.Cells[2, 12] = leaguetracker.CleanSheets;
                oSheet.Cells[2, 13] = leaguetracker.Goalconversationrate;
                oSheet.Cells[2, 14] = leaguetracker.Shotsontarget;



                oXL.UserControl = false;
                oWB.SaveAs(this.BrowserDataFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;
                string srcPath = (this.SrcDataFilePath);

                oWB = (excel._Workbook)(oXL.Workbooks.Open(srcPath));
                oSheet = oWB.Worksheets.get_Item(1);

                string destPath = (this.BrowserDataFilePath);
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(destPath));
                oXL.Visible = true;
                dSheet = oWB.Worksheets.Add();

                excel.Range from = oSheet.Range["A:A,B:B,C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N"];
                excel.Range torange = dSheet.Range["A1:N1"];

                from.Copy(torange);


                oXL.ActiveSheet.Range["A4"] = "=VLOOKUP(A2,Sheet1!B:B,1,FALSE)";
                oXL.ActiveSheet.Range["B4"] = "=VLOOKUP(B2,Sheet1!A:A,1,FALSE)";
                oXL.ActiveSheet.Range["C4"] = "=VLOOKUP(C2,Sheet1!D:D,1,FALSE)";
                oXL.ActiveSheet.Range["D4"] = "=VLOOKUP(D2,Sheet1!N:N,1,FALSE)";
                oXL.ActiveSheet.Range["E4"] = "=VLOOKUP(E2,Sheet1!G:G,1,FALSE)";
                oXL.ActiveSheet.Range["F4"] = "=VLOOKUP(F2,Sheet1!F:F,1,FALSE)";
                oXL.ActiveSheet.Range["G4"] = "=VLOOKUP(G2,Sheet1!L:L,1,FALSE)";
                oXL.ActiveSheet.Range["H4"] = "=VLOOKUP(H2,Sheet1!H:H,1,FALSE)";
                oXL.ActiveSheet.Range["I4"] = "=VLOOKUP(I2,Sheet1!I:I,1,FALSE)";
                oXL.ActiveSheet.Range["J4"] = "=VLOOKUP(J2,Sheet1!M:M,1,FALSE)";
                oXL.ActiveSheet.Range["K4"] = "=VLOOKUP(K2,Sheet1!C:C,1,FALSE)";
                oXL.ActiveSheet.Range["L4"] = "=VLOOKUP(L2,Sheet1!E:E,1,FALSE)";
                oXL.ActiveSheet.Range["M4"] = "=VLOOKUP(M2,Sheet1!K:K,1,FALSE)";
                oXL.ActiveSheet.Range["N4"] = "=VLOOKUP(N2,Sheet1!J:J,1,FALSE)";

                oXL.UserControl = false;
                oWB.SaveAs(this.BrowserDataFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;

            }

            catch (Exception) { }
        }
    }
}
