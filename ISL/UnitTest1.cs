using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using SeleniumExtras.WaitHelpers;

namespace ISL
{
    [TestClass]
    public class UnitTest1
    {
            IWebDriver driver = new ChromeDriver();
       

       [TestMethod]
        public void Playerstats()
        {
            driver.Navigate().GoToUrl("https://www.indiansuperleague.com/stats/115-138-goals-player-statistics");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
            
            //Cookies Gotit button
            driver.FindElement(By.CssSelector(".action-btn")).Click();

            Thread.Sleep(2000);
            Seasondropdown();
            Thread.Sleep(1000);
            //reading data for goals
            LoadMoreButton();
          
            // driver.FindElement(By.XPath("/html/body/div[1]/header/section/div/div/div[4]/div/nav/ul/li[7]/a")).Click();
            ReadingExcel ObjData = new ReadingExcel(driver);
            ObjData.GetExcelData();

            Statsdropdown();


            //reading data for cleansheets
           
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[2]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            CleanSheetStats ObjectData1 = new CleanSheetStats(driver);
            ObjectData1.GetCleanSheetData();

            Statsdropdown();

            //reading data for fouls
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[3]")).Click();
            LoadMoreButton();
            FoulsStats ObjectData2 = new FoulsStats(driver);
            ObjectData2.GetFoulSheetData();

            Statsdropdown();

            //reading data for assists
  
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[1]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();

            Assists ObjectData3 = new Assists(driver);
            ObjectData3.GetAssistsSheetData();

            Statsdropdown();
            
            //reading data for passes
            
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[6]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();

            PassesStats ObjectData4 = new PassesStats(driver);
            ObjectData4.GetPassesSheetData();

            Statsdropdown();
            
            //Interception
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[5]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            Interception Objinterception = new Interception(driver);
            Objinterception.GetInterceptionData();

            Statsdropdown();

            //Red Card
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[7]")).Click();
            Thread.Sleep(5000);

            LoadMoreButton();
            Redcard Objredcard = new Redcard(driver);
            Objredcard.GetRedcardData();

            Statsdropdown();

            //Saves
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[8]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            Saves Objsaves = new Saves(driver);
            Objsaves.GetSavesData();

            Statsdropdown();

            //Shots
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[9]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            Shots Objshots = new Shots(driver);
            Objshots.GetShotsData();

            Statsdropdown();
            Statsdropdownscroll();
            //ShotonTarget
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[10]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            ShotonTarget Objshotontrgt = new ShotonTarget(driver);
            Objshotontrgt.GetShotonTargetData();

            Statsdropdown();
            Statsdropdownscroll();
            //Tackles
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[11]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            Tackles Objtackle = new Tackles(driver);
            Objtackle.GetTacklesData();

            Statsdropdown();
            Statsdropdownscroll();
            //Touches
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[12]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            Touches Objtouches = new Touches(driver);
            Objtouches.GetTouchesData();

            Statsdropdown();
            Statsdropdownscroll();
            //Yellowcard
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[13]")).Click();
            Thread.Sleep(5000);
            LoadMoreButton();
            Yellowcard Objyellowcard = new Yellowcard(driver);
            Objyellowcard.GetYellowcardData();

            Statsdropdown();


           


        }
        [TestMethod]
        public void Teamdata()
        {
            driver.Navigate().GoToUrl("https://www.indiansuperleague.com/stats/115-128-goals-club-statistics");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);

            //Cookies Gotit button
            GotITButton();

            Thread.Sleep(10000);
            Seasondropdown();
            Thread.Sleep(1000);
          
            //Team Gaol
            TeamGoal objteamgoal = new TeamGoal(driver);
            objteamgoal.GetTeamGoalData();

            //Team Crossess
            clubstats();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[1]")).Click();

            Teamcrossess objteamcrosses = new Teamcrossess(driver);
            objteamcrosses.GetTeamcrossessData();

            //team cleansheet
            clubstats();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[2]")).Click();

            TeamCleansheet objteamcleansheet = new TeamCleansheet(driver);
            objteamcleansheet.GetTeamCleanSheetsData();

            //Team Fouls
            clubstats();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[4]")).Click();

            Teamfouls objteamfoul = new Teamfouls(driver);
            objteamfoul.GetTeamfoulsData();

            //Team Passes
            clubstats();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[7]")).Click();

            Teampasses objteampasses = new Teampasses(driver);
            objteampasses.GetTeampassesData();

            //Team Saves
            clubstats();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[8]")).Click();

            TeamSaves objteamsaves = new TeamSaves(driver);
            objteamsaves.GetTeamSavesData();

            //Team Shots
            clubstats();
            Statsdropdownscroll();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[9]")).Click();

            TeamShots objteamshots = new TeamShots(driver);
            objteamshots.GetTeamShotsData();
            
            //Team Redcards
            clubstats();
            Statsdropdownscroll();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[10]")).Click();

            TeamRedcard objteamredcard = new TeamRedcard(driver);
            objteamredcard.GetTeamRedcardData();

            //Team Tackle
            clubstats();
            Statsdropdownscroll();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[11]")).Click();

            Teamtackle objteamtackle = new Teamtackle(driver);
            objteamtackle.GetTeamtackleData();

            //Team Touches
            clubstats();
            Statsdropdownscroll();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[12]")).Click();

            Teamtouches objteamtouches = new Teamtouches(driver);
            objteamtouches.GetTeamtouchesData();

            //Team yellowcard
            Thread.Sleep(2000);
            clubstats();
            Statsdropdownscroll();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[14]")).Click();

            Teamyellowcard objteamyellowcard = new Teamyellowcard(driver);
            objteamyellowcard.GetTeamyellowcardData();
        }
        private void clubstats()
        {
            //driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[1]")).Click();


            var clubstat = driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[1]"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(clubstat);
            action1.Perform();
            clubstat = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[1]")));
            clubstat.Click();
        }
        [TestMethod]
        public void LeagueTracker()
        {
            //reading data for league tracker
            driver.Navigate().GoToUrl("https://www.indiansuperleague.com/");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);

            //Cookies Gotit button
            driver.FindElement(By.CssSelector(".action-btn")).Click();

            OverallStats ObjectData5 = new OverallStats(driver);
            ObjectData5.GetLeagueTrackerData();
        }
        [TestMethod]
        public void Squadcheck()
        {
            driver.Navigate().GoToUrl("https://www.indiansuperleague.com/clubs");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);

           
            //Cookies Gotit button
            GotITButton();
            Thread.Sleep(2000);

            //Team Profile tab
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div[1]/div/div/div[2]/div/div[4]/a")).Click();
            //Squad tab
            SquadTab();

            GenericSquadCheck _ATKsquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\Squaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _ATKsquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Teamselection
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[2]/a/div/img")).Click();
            Thread.Sleep(2000);

            //BengluruTeam
            SquadTab();
            GenericSquadCheck _BengaluruSquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\BengaluruSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _BengaluruSquad.GetGenericSquadCheckData();

            TeamsLogo();
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[3]/a/div/img")).Click();
            Thread.Sleep(2000);

            //ChennayianTeam
            SquadTab();
            GenericSquadCheck _Chennaisquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\ChennaiSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Chennaisquad.GetGenericSquadCheckData();

            TeamsLogo();
            //FCGoa
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[4]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Goasquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\FCGoaSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Goasquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Hydrabad FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[5]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Hydrabadsquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\HydrabadSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Hydrabadsquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Jamshedpur FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[6]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Jamshedpursquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\JamshedpurSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Jamshedpursquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Kerala FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[7]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Keralasquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\KeralaSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Keralasquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Mumbai FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[8]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Mumbaisquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\MumbaiSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Mumbaisquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Northeast FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[9]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Northeastsquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\NortheastSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Northeastsquad.GetGenericSquadCheckData();

            TeamsLogo();
            //Odisha FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[10]/a/div/img")).Click();
            Thread.Sleep(2000);
            SquadTab();

            GenericSquadCheck _Odhishasquad = new GenericSquadCheck(driver, @"D:\Automation\ISL\Player Data\OdishaSquaddetail.xlsx", @"D:\Automation\ISL\Player Data\ISLPlayerData.xlsx");
            _Odhishasquad.GetGenericSquadCheckData();

        }

        [TestMethod]
        public void Searchplayers()
        {
            driver.Navigate().GoToUrl("https://www.indiansuperleague.com/");

            driver.Manage().Window.Maximize();
            
            string ReadExcel;
            int rctn = 6;

            excel.Application x1app = new excel.Application();
            excel.Workbook x1workbook = x1app.Workbooks.Open(@"D:\Automation\ISL\Player search file\player_data1.xlsx");
            excel.Worksheet x1worksheet = x1workbook.Sheets[1];


            excel.Range x1range = x1worksheet.UsedRange;



            for (int i = 6; i <= rctn; i++)
            {

                for (int j = 2; j <= 78; j++)
                {

                    ReadExcel = x1range.Cells[i][j].Text.ToString();
                    driver.FindElement(By.XPath("/html/body/div[1]/header/section/div/div/div[3]/div/div[2]/div[1]/ul/li[5]/a")).Click();

                    var search = driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div/div[2]/input"));
                    search.SendKeys(ReadExcel);
                    search.SendKeys(Keys.Enter);

                    var visible = driver.FindElement(By.XPath("//*[@id='cookiebtn']"));
                    
                    if (visible.Displayed)
                    {
                        //Assert.AreEqual(true, visible.Displayed);
                        visible.Click();
                    }
                    
                    
                    /* FunctionalLibrary.TryFindElement(driver, "//*[@id='cookiebtn']" );

                         var visible =  IsElementVisible(element);
                         if (visible)
                         {
                         element.Click();
                         }*/



                    var viewprofile = driver.FindElement(By.Id("player-list"));
                    var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
                    action1.MoveToElement(viewprofile);
                    action1.Perform();
                    viewprofile = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                               .Until(driver => driver.FindElement(By.Id("player-list")));

                    var playername = driver.FindElement(By.CssSelector(".article-name")).Text;

                    if (ReadExcel.Equals(playername))
                    {
                        Debug.WriteLine(playername,"player name is correct");
                    }
                    else
                    {
                        Debug.WriteLine(playername,"player name is incorrect");
                    }
                   var webe =   driver.FindElement(By.Id("player-list"));


                  var href =  webe.FindElement(By.CssSelector("a")).GetAttribute("href");
                   
                    
                   // WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(1));

                    //wait.Until(ExpectedConditions.UrlContains(href));

                    if (href == null)
                    {
                        Debug.WriteLine(playername,"Player not clikable");
                       
                      
                       
                    }
                    else
                    {
                        Debug.WriteLine(playername, "player is clikable");
                        webe.Click();
                        Thread.Sleep(1000);
                        var playerdetail = driver.FindElement(By.CssSelector(".si-player-name")).Text;
                        if (playername.Equals(playerdetail))
                        {
                            Debug.WriteLine("Redirected on detail page");
                        }
                    }



                    Thread.Sleep(2000);
                    /*var viewprofilesroll = driver.FindElement(By.CssSelector(".article-content"));
                    var action2 = new OpenQA.Selenium.Interactions.Actions(driver);
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
                    action2.MoveToElement(viewprofilesroll);
                    action2.Perform();
                    viewprofilesroll = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                               .Until(driver => driver.FindElement(By.CssSelector(".article-content")));*/
                    //svar viewprofilesroll = driver.FindElement(By.CssSelector(".si-player-name"));

                    

                }

            }


        }

       
        [TestMethod]
        public void TeamwiseLeagueTracker()
        {
            driver.Navigate().GoToUrl("https://www.indiansuperleague.com/clubs/499-atk-profile");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
            GotITButton();
            Teamtracker();

           
            GenericTracker _Atkleaguetracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\ATKleaguetracker.xlsx", @"D:\Automation\ISL\Data Files\FCATKtracker.xls");
            _Atkleaguetracker.GetData();

            TeamsLogo();

            //Bengaluru FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[2]/a/div/img")).Click();
            Thread.Sleep(2000);
            Teamtracker();
            
            GenericTracker _Bengulurutrackercs = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\Bengulurutrackercsr.xlsx", @"D:\Automation\ISL\Data Files\FCBengalurutracker.xls");
            _Bengulurutrackercs.GetData();
            TeamsLogo();

            //Chennai FC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[3]/a/div/img")).Click();
            Teamtracker();
            GenericTracker _Chennaitracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\ChennaiTracker.xlsx", @"D:\Automation\ISL\Data Files\FCChennaiFC.xls");
            _Chennaitracker.GetData();


            TeamsLogo();
            //FCGoa 
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[4]/a/div/img")).Click();
            Teamtracker();
            GenericTracker _Goatracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\FCGoa.xlsx", @"D:\Automation\ISL\Data Files\FCGoaFC.xls");
            _Goatracker.GetData();

            TeamsLogo();
            //HydrabadFC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[5]/a/div/img")).Click();
            Teamtracker();
            GenericTracker _hydrabadtracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\HydrabadFC.xlsx", @"D:\Automation\ISL\Data Files\FCHydrabadFC.xls");
            _hydrabadtracker.GetData();

            TeamsLogo();
            //JamshedpurFC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[6]/a/div/img")).Click();
            Teamtracker();

            GenericTracker _jamshedpurtracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\JamshedpurFC.xlsx", @"D:\Automation\ISL\Data Files\FCJamshedpurFC.xls");
            _jamshedpurtracker.GetData();

            TeamsLogo();
            //KeralaFC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[7]/a/div/img")).Click();
            Teamtracker();

            GenericTracker _keralatracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\KeralaFC.xlsx", @"D:\Automation\ISL\Data Files\FCKeralaFC.xls");
            _keralatracker.GetData();
            TeamsLogo();

            //MumbaiFC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[8]/a/div/img")).Click();
            Teamtracker();
            GenericTracker _Mumbaitracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\MumbaiFC.xlsx", @"D:\Automation\ISL\Data Files\FCMumbaiFC.xls");
            _Mumbaitracker.GetData();
            

            TeamsLogo();
            //NortheastFC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[9]/a/div/img")).Click();
            Teamtracker();
            GenericTracker _Northeasttracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\NortheastFC.xlsx", @"D:\Automation\ISL\Data Files\FCNortheastFC.xls");
            _Northeasttracker.GetData();

            TeamsLogo();

            //OdhissaFC
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[1]/div/div/div[1]/div/div/div[10]/a/div/img")).Click();
            Teamtracker();
            GenericTracker _Odishatracker = new GenericTracker(driver, @"D:\Automation\ISL\Club Stats\OdishaFC.xlsx", @"D:\Automation\ISL\Data Files\ATKtracker.xls");
            _Odishatracker.GetData();

        }
        private void Teamtracker()
        {
            var tracker = driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[7]/div/div/div/div/section/component/div[1]/h2"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(tracker);
            action1.Perform();
            tracker = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[7]/div/div/div/div/section/component/div[1]/h2")));
        }
        private void SquadTab()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section[1]/div/div/div/div/section/component[2]/div/component/ul/li[2]/a")).Click();
        }
        private void TeamsLogo()
        {
            var Logo = driver.FindElement(By.CssSelector(".teams-logo"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(Logo);
            action1.Perform();
            Logo = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.CssSelector(".teams-logo")));
        }
        private void Seasondropdown()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div/div[2]/div[1]")).Click();

            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[2]/div/div[2]/div[2]/ul/li[2]")).Click();
        }
        private void GotITButton()
        {
            //Cookies Gotit button
            driver.FindElement(By.CssSelector(".action-btn")).Click();
        }
      
            
        
        private void LoadMoreButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 1, 0));
            var button = wait.Until(driver => driver.FindElement(By.CssSelector(".si-stats-more-btn")));

            while (button.GetCssValue("display") != "none")
            {
                var action = new OpenQA.Selenium.Interactions.Actions(driver);
                action.MoveToElement(button);
                action.Perform();
                button.Click();
                button = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                            .Until(driver => driver.FindElement(By.CssSelector(".si-stats-more-btn")));
            }
        }
        private void Statsdropdown()
        {
            var DropDown = driver.FindElement(By.CssSelector(".si-stats-dropdown-container"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(DropDown);
            action1.Perform();
            DropDown = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.CssSelector(".si-stats-dropdown-container")));
            DropDown.Click();
        }
        private void Statsdropdownscroll()
        {
            var DropDown = driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[13]"));
            var action1 = new OpenQA.Selenium.Interactions.Actions(driver);
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
            action1.MoveToElement(DropDown);
            action1.Perform();
            DropDown = new WebDriverWait(driver, new TimeSpan(0, 1, 0))
                       .Until(driver => driver.FindElement(By.XPath("/html/body/div[1]/section/myapp/section/div/div/div/div/section/component/div[2]/div/div/div/div/div[2]/div[1]/div[1]/div/div/div[2]/ul/li[13]")));
        }
       
        
       

    }
}
