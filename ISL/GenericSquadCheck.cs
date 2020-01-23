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
using OpenQA.Selenium.Interactions;

namespace ISL
{
    

        class Squadsdetail
    {

        public string PlayerName { get; set; }

        public string JersyNo { get; set; }
    }
    class Urllist
    {
        public string URL { get; set; }
    }

    class GenericSquadCheck
    {
        public IWebDriver driver { get; private set; }
        public string BrowserDataFilePath { get; private set; }

        private string SrcDataFilePath;


        public GenericSquadCheck(IWebDriver driver, string BrowserDataFilePath, string SrcDataFilePath)
        {
            this.driver = driver;
            this.BrowserDataFilePath = BrowserDataFilePath;
            this.SrcDataFilePath = SrcDataFilePath;
        }
        public IWebElement GetGenericSquadCheckData()
        {


            var Urllist = new List<Urllist>();

            var Squadsdetail = new List<Squadsdetail>();
            Thread.Sleep(2000);

            var allDivs = driver.FindElements(By.CssSelector(".si-team-info"));

            foreach (var div in allDivs)
            {
    

                var url = div.GetAttribute("href");


                Urllist.Add(new Urllist { URL = url });
                TestContext.Out.WriteLine($"URL: { url } ");

            }
            Urllist.Reverse();

            foreach (var p in Urllist.Skip(1))
            {
                var link = p.URL.ToString();
                driver.Navigate().GoToUrl(link);
                Thread.Sleep(3000);

                var playerd = driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[2]/div/div/div/div/section/component/div/div/div/div/div/div[3]/div[1]/span"));
                var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
                ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
                action1.MoveToElement(playerd);
                action1.Perform();
                playerd = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                           .Until(driver => driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[2]/div/div/div/div/section/component/div/div/div/div/div/div[3]/div[1]/span")));


                var allDivs1 = driver.FindElements(By.CssSelector(".si-player-details"));
                foreach (var div1 in allDivs1)
                {
                    var playername = div1.FindElement(By.CssSelector(".si-player-name")).Text;

                    var jerseyno = div1.FindElement(By.CssSelector(".si-player-jersey")).Text;
                    Squadsdetail.Add(new Squadsdetail { PlayerName = playername, JersyNo = jerseyno });
                    TestContext.Out.WriteLine($"PlayerName: { playername }  | JersyNo:{jerseyno}");


                }



            }
            CreateXlSheet(Squadsdetail);

            return null;
        }

        public void CreateXlSheet(List<Squadsdetail> Squadsdetail)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel._Worksheet dSheet;

            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                //oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());


                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;



                oSheet.Cells[1, 1] = "TeamTitle";

                oSheet.Cells[1, 2] = "JersyNo";

                for (int i = 0; i < Squadsdetail.Count; i++)
                {
                    int row = i + 1;

                    oSheet.Cells[row, 1] = Squadsdetail[i].PlayerName;
                    oSheet.Cells[row, 2] = Squadsdetail[i].JersyNo;



                }


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
                dSheet = oWB.Worksheets.Add();

                excel.Range from = oSheet.Range["B:B,E:E"];
                excel.Range torange = dSheet.Range["A1:B1"];

                from.Copy(torange);


                oXL.ActiveSheet.Range["D2:D1000"] = "=VLOOKUP(A2,Sheet1!A:B,1,False)";
                oXL.ActiveSheet.Range["E2:E1000"] = "=VLOOKUP(A2,Sheet1!A:B,2,FALSE)";
                oXL.ActiveSheet.Range["F2:F1000"] = "=EXACT(D:D,A:A)";
                oXL.ActiveSheet.Range["G2:G1000"] = "=EXACT(E:E,B:B)";


                oXL.UserControl = false;
                oWB.SaveAs(this.BrowserDataFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXL.Visible = true;


            }


            catch (Exception)
            {

            }





        }
    }
}
