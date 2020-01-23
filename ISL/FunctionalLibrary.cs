using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;


namespace ISL

{
    class FunctionalLibrary
    {
        

        public static void clickAction(IWebDriver driver, string LocaterValue, string LocaterType)
        {
            if (LocaterType == "id")
            {
                driver.FindElement(By.Id(LocaterValue)).Click();
            }
                

            if (LocaterType == "xpath")
            {
                driver.FindElement(By.XPath(LocaterValue)).Click();
            }


        }

       /* public static void TypeAction(IWebDriver driver, string LocaterValue, string LocaterType, string Value)
        {
            if (LocaterType == "id")
            { 
            driver.FindElement(By.Id(LocaterValue)).Clear();
            driver.FindElement(By.Id(LocaterValue)).SendKeys(Value);
        }
            if (LocaterType == "xpath")
            {
                driver.FindElement(By.XPath(LocaterValue)).Clear();
                driver.FindElement(By.XPath(LocaterValue)).SendKeys(Value);
            }
        }*/

        public static void MouseOver(IWebDriver driver, string LocaterValue)

        {
            IWebElement element = driver.FindElement(By.XPath(LocaterValue));

            Actions action = new Actions(driver);

            action.MoveToElement(element).Perform();


        }
        public static void TryFindElement(IWebDriver driver, string LocaterValue)
        {
             
            try
            {

               IWebElement element = driver.FindElement(By.XPath(LocaterValue));
            }
            catch (NoSuchElementException)
            {
               
            }
           
        }

        

        public static void waitForElement(IWebDriver driver, string Locatervalue)

        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(1));

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(Locatervalue)));


        }

        public static void screenShot(IWebDriver driver)
        {

            string imgName = DateTime.Now.ToString("dd/MM/yyyy-HH-mm-ss");


            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();

            ss.SaveAsFile("D:\\ISLproject\\ScreenShots\\" + imgName + ".png");

        }

        public static string ReadExcelData(IWebDriver driver, int s, int i, int j)
        {
            excel.Application xlapp = new excel.Application();

            excel.Workbook xlwb = xlapp.Workbooks.Add(@"D:\\yahoo2.0\\TestInput\\yahooinput.xlsx");

            excel._Worksheet xlsheet = xlwb.Sheets[s];

            excel.Range xlrange = xlsheet.UsedRange;

            string data = xlrange.Cells[i][j].value2;

            return data;
        }



        public static void SetDataExcel(IWebDriver driver,int i,int j,string data)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
           
            object misvalue = System.Reflection.Missing.Value;

            oXL = new Microsoft.Office.Interop.Excel.Application();

          oWB=  (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());

            oSheet = (excel.Worksheet)oWB.Sheets[1];

            oSheet.Cells[i][j] = data;

            oWB.SaveAs("C:\\Users\\shaik\\Desktop\\input.xlsx");


        }

        public static void setdata(IWebDriver driver, int s, int i,int j, string data)
        {
            excel.Application xlapp = new excel.Application();

            excel.Workbook xlwb = xlapp.Workbooks.Add();

            excel._Worksheet sheet = xlwb.Sheets[s];

            sheet.Cells[i][j] = data;

            xlapp.DisplayAlerts = false;

            xlwb.SaveAs("C:\\Users\\shaik\\Desktop\\test\\testinput.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                   false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlwb.Close();
        }
    }
}
