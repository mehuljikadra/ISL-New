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
    class Leaguetracker
    {

        public string Goals1 { get; set; }
        public string CleanSheets { get; set; }
        public string Fouls { get; set; }

        public string Passes { get; set; }

        public string MinsPerGoal { get; set; }
        public string Matches { get; set; }
        public string AvgGoalsMatch { get; set; }
        public string GoalConversionRate { get; set; }
        public string AvgPassPerGame { get; set; }
        public string RedCards { get; set; }
        public string YellowCards { get; set; }
        public string Tackles { get; set; }
        public string Interceptions { get; set; }
        public string Blocks { get; set; }
        public string Assist { get; set; }
        public string Touches { get; set; }
        public string Foulscommited { get; set; }
        public string Foulssuffered { get; set; }
        public string Clearences { get; set; }
        public string Goalconversationrate { get; set; }
        public string Shotsontarget { get; set; }

    }
    class OverallStats
    {
        private IWebDriver driver;
        private object missing;

        public OverallStats(IWebDriver driver)
        {
            this.driver = driver;
        }
        public IWebElement GetLeagueTrackerData()
        {
          

            var Leaguetrackerdata = driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[7]/div/div/div/div/section/component/div"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(Leaguetrackerdata);
            Thread.Sleep(2000);
            action1.Perform();
            Leaguetrackerdata = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[7]/div/div/div/div/section/component/div")));


            excel.Application x1app = new excel.Application();
            excel.Workbook x1workbook = x1app.Workbooks.Open(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\Excel_stats.xlsx", ReadOnly: false);
            excel.Worksheet xlWorkSheet = (excel.Worksheet)x1workbook.Worksheets[1];

            excel.Range x1range = xlWorkSheet.UsedRange;



            driver.FindElements(By.CssSelector(".si-league-tracker"));

            var goals1 = driver.FindElement(By.CssSelector(".si-fkt-bOne .si-fkt-sctn-number")).Text;
            var minspergoal = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col1 .si-fkt-sctn-number")).Text;
            var avggoalsmatch = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col2 .si-fkt-sctn-number")).Text;
            var goalconversionrate = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col3 .si-fkt-sctn-number")).Text;
            var passes = driver.FindElement(By.CssSelector(".si-fkt-bTwo-col4 .si-fkt-sctn-number")).Text;
            var avgpasspergame = driver.FindElement(By.CssSelector(".si-fkt-bThree .si-fkt-sctn-number")).Text;
            var redcards = driver.FindElement(By.CssSelector(".si-fkt-bSeven .si-fkt-sctn-number")).Text;
            var yellowcards = driver.FindElement(By.CssSelector(".si-fkt-bEight .si-fkt-sctn-number")).Text;
            var tackles = driver.FindElement(By.CssSelector(".si-fkt-bEight.AEight .si-fkt-sctn-number")).Text;
            var fouls = driver.FindElement(By.CssSelector(".si-fkt-bNine .si-fkt-sctn-number")).Text;
            var interceptions = driver.FindElement(By.CssSelector(".si-fkt-bTen .si-fkt-sctn-number")).Text;
            var blocks = driver.FindElement(By.CssSelector(".si-fkt-bEleven .si-fkt-sctn-number")).Text;
            var cleansheet = driver.FindElement(By.CssSelector(".si-fkt-bTwelve .si-fkt-sctn-number")).Text;

            var leaguetracker = new Leaguetracker
            {
                Goals1 = goals1,
                MinsPerGoal = minspergoal,
                AvgGoalsMatch = avggoalsmatch,
                GoalConversionRate = goalconversionrate,
                Passes = passes,
                AvgPassPerGame = avgpasspergame,
                RedCards = redcards,
                YellowCards = yellowcards,
                Tackles = tackles,
                Fouls = fouls,
                Interceptions = interceptions,
                Blocks = blocks,
                CleanSheets = cleansheet
            };
            TestContext.Out.WriteLine($"Goals1: { goals1 } | MinsPerGoal: { minspergoal } | AvgGoalsMatch: { avggoalsmatch } | AvgGoalsMatch: { avggoalsmatch }" +
                $"| GoalConversionRate: { goalconversionrate } | Passes: { passes } | AvgPassPerGame: { avgpasspergame } | RedCards: { redcards }" +
                $"| YellowCards: { yellowcards } | Tackles: { tackles } | Fouls: { fouls } | Interceptions: { interceptions }" +
                $" | Blocks: { blocks } | CleanSheets: { cleansheet }");



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
                oSheet.Cells[1, 2] = "MinsPerGoal";
                oSheet.Cells[1, 3] = "AvgGoalsMatch";

                oSheet.Cells[1, 4] = "GoalConversionRate";
                oSheet.Cells[1, 5] = "Passes";
                oSheet.Cells[1, 6] = "AvgPassPerGame";

                oSheet.Cells[1, 7] = "RedCards";
                oSheet.Cells[1, 8] = "YellowCards";
                oSheet.Cells[1, 9] = "Tackles";


                oSheet.Cells[1, 10] = "Fouls";
                oSheet.Cells[1, 11] = "Interceptions";
                oSheet.Cells[1, 12] = "Blocks";
                oSheet.Cells[1, 13] = "CleanSheets";



                oSheet.Cells[2, 1] = leaguetracker.Goals1;
                oSheet.Cells[2, 2] = leaguetracker.MinsPerGoal;
                
                oSheet.Cells[2, 3] = leaguetracker.AvgGoalsMatch;
                oSheet.Cells[2, 4] = leaguetracker.GoalConversionRate;
                oSheet.Cells[2, 5] = leaguetracker.Passes;
                oSheet.Cells[2, 6] = leaguetracker.AvgPassPerGame;
                oSheet.Cells[2, 7] = leaguetracker.RedCards;
                oSheet.Cells[2, 8] = leaguetracker.YellowCards;
                oSheet.Cells[2, 9] = leaguetracker.Tackles;
                oSheet.Cells[2, 10] = leaguetracker.Fouls;
                oSheet.Cells[2, 11] = leaguetracker.Interceptions;
                oSheet.Cells[2, 12] = leaguetracker.Blocks;
                oSheet.Cells[2, 13] = leaguetracker.CleanSheets;




                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\leaguetracker.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;

                string srcPath = (@"C:\Users\aditya.bhosle\Desktop\ISL\FootballOveralltracker.xls");

                oWB = (excel._Workbook)(oXL.Workbooks.Open(srcPath));
                oSheet = oWB.Worksheets.get_Item(1);

                string destPath = (@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\leaguetracker.xlsx");
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(destPath));
                oXL.Visible = true;
                dSheet = oWB.Worksheets.Add();

                excel.Range from = oSheet.Range["A:A,B:B,C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N"];
                excel.Range torange = dSheet.Range["A1:N1"];

                from.Copy(torange);


                oXL.ActiveSheet.Range["A16"] = "=VLOOKUP(A13,Sheet1!A:A,1,FALSE)";
                oXL.ActiveSheet.Range["B16"] = "=VLOOKUP(B13,Sheet1!H:H,1,FALSE)";
                oXL.ActiveSheet.Range["C16"] = "=VLOOKUP(C13,Sheet1!G:G,1,FALSE)";
                oXL.ActiveSheet.Range["D16"] = "=VLOOKUP(D13,Sheet1!M:M,1,FALSE)";
                oXL.ActiveSheet.Range["E16"] = "=VLOOKUP(E13,Sheet1!J:J,1,FALSE)";
                oXL.ActiveSheet.Range["F16"] = "=VLOOKUP(F13,Sheet1!B:B,1,FALSE)";
                oXL.ActiveSheet.Range["G16"] = "=VLOOKUP(G13,Sheet1!D:D,1,FALSE)";
                oXL.ActiveSheet.Range["H16"] = "=VLOOKUP(H13,Sheet1!I:I,1,FALSE)";
                oXL.ActiveSheet.Range["I16"] = "=VLOOKUP(I13,Sheet1!E:E,1,FALSE)";
                oXL.ActiveSheet.Range["J16"] = "=VLOOKUP(J13,Sheet1!L:L,1,FALSE)";
                oXL.ActiveSheet.Range["K16"] = "=VLOOKUP(K13,Sheet1!F:F,1,FALSE)";

                oXL.UserControl = false;
                oWB.SaveAs(@"C:\Users\aditya.bhosle\source\repos\Data\Data\NewFolder1\leaguetracker.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;

            }
            catch (Exception) { }
        }

    }


}