using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using System.Data;

namespace BrowserAutomation
{
    public class clsBrowserAutomation
    {

    #region "Open/Close Browsers"

        /// <summary>
        /// Open the new Chrome Browser and return the chrome driver as IwebDriver
        /// Note: this fucntion will close the chrome/chromedriver if its opened already, otherwise new chrome driver will not work perfectly
        /// Note: Ensure the chromedriver is downloaded and placed into your exe folder
        /// </summary>
        /// <returns></returns>
		public IWebDriver openChrome()
        {
            OpenQA.Selenium.Chrome.ChromeOptions choptions = new OpenQA.Selenium.Chrome.ChromeOptions();
            //choptions.BinaryLocation = "C:\\Program Files\\Google\\\\Application\\Chrom.exe";
            choptions.AddArgument("--test-type");
            choptions.AddArgument("--disable-plugins");
            choptions.AddArgument("--disable-extensions");
            //choptions.AddArgument("--enable-automation");
            choptions.AddArgument("--no-sandbox");
            choptions.AddUserProfilePreference("credentials_enable_service", false);
            choptions.AddUserProfilePreference("profile.password_manager_enabled", false);
            //choptions.AddArgument("--start-maximized");

            //choptions.AddArgument("--allow-external-pages");
            //choptions.AddArgument("--allow-running-insecure-content");
            //choptions.AddArgument("--new-window");
            //choptions.AddArguments("--enable-strict-powerful-feature-restrictions");
            //choptions.AddUserProfilePreference("profile.default_content_setting_values.images", 0);

            //Close the previous  driver is it have opened            
            foreach (System.Diagnostics.Process proTem in System.Diagnostics.Process.GetProcessesByName("chromedriver"))
            {
                proTem.CloseMainWindow();
            }
            //Close the chrome browser if its opened already
            foreach (System.Diagnostics.Process proTem in System.Diagnostics.Process.GetProcessesByName("chrome"))
            {
                proTem.CloseMainWindow();
            }

            //Open the new 
            return (new OpenQA.Selenium.Chrome.ChromeDriver(Environment.CurrentDirectory, choptions, new TimeSpan(0, 5, 0))); 
        }

        /// <summary>
        /// Open the opear window and return the opeardriver as iwebdriver
        /// Opera portable allows opening multiple browsers seemlessly, so upto you to decide to close the existing windows by passing true to the parameter p_clnCloseOpenedOperaDriver
        /// Note: Ensure the opera driver is downloaded and placed into your exe folder in the right folder(FolderName:32bit-"operadriver_win32" 64bit-"operadriver_win64")
        /// </summary>
        /// <param name="p_clnCloseOpenedOperaDriver"></param>
        /// <returns></returns>
        public IWebDriver openOpera(bool p_clnCloseOpenedOperaDriver = true)
        {
            string strDriverLocation = "";
            if (Environment.Is64BitOperatingSystem == true)
                strDriverLocation = Environment.CurrentDirectory + "\\operadriver_win64";
            else
                strDriverLocation = Environment.CurrentDirectory + "\\operadriver_win32";

            OpenQA.Selenium.Opera.OperaOptions opOptions = new OpenQA.Selenium.Opera.OperaOptions();
            opOptions.AddArgument("--test-type");
            opOptions.AddArgument("disable-plugins");
            opOptions.AddArgument("disable-extensions");
            opOptions.AddArgument("no-sandbox");
            opOptions.AddUserProfilePreference("credentials_enable_service", false);
            opOptions.AddUserProfilePreference("profile.password_manager_enabled", false);
            //Disable java script
            opOptions.AddArgument("--disable-javascript");

            //opOptions.AddArgument("--start-maximized");
            //Opeara executable location
            opOptions.BinaryLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\opera\\launcher.exe";


            if (p_clnCloseOpenedOperaDriver == true)
            {
                //Close the previous  driver is it have opened            
                foreach (System.Diagnostics.Process proTem in System.Diagnostics.Process.GetProcessesByName("operadriver"))
                {
                    proTem.CloseMainWindow();
                }
            }

            //Open the new 
            return (new OpenQA.Selenium.Opera.OperaDriver(strDriverLocation, opOptions, new TimeSpan(0, 5, 0)));
        }

        /// <summary>
        /// Open the new IE window and return the IEDriver as iwebdriver
        ///Note: Ensure the IE driver is downloaded and placed into your exe folder in the right folder(FolderName:32bit-"IEdriver_win32" 64bit-"IEdriver_win64") , if not pass the driver location to the argument "p_strDriverLocation"
        /// </summary>
        /// <param name="p_clnCloseOpenedChromDriver">Whether to close the existing IE or not</param>
        /// <param name="p_strDriverLocation">Location of the webdriver</param>
        /// <returns></returns>
        public IWebDriver openIE(bool p_clnCloseOpenedChromDriver = true, string p_strDriverLocation = "")
        {
            string strDriverLocation = "";
            if (p_strDriverLocation == "")
            {
                if (Environment.Is64BitOperatingSystem == true)
                    strDriverLocation = Environment.CurrentDirectory + "\\IEdriver_win64";
                else
                    strDriverLocation = Environment.CurrentDirectory + "\\IEdriver_win32";
            }
            strDriverLocation = p_strDriverLocation;

            OpenQA.Selenium.IE.InternetExplorerOptions ieOptions = new OpenQA.Selenium.IE.InternetExplorerOptions();
            ieOptions.IgnoreZoomLevel = true;
            ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            //Open the new 
            return (new OpenQA.Selenium.IE.InternetExplorerDriver(strDriverLocation, ieOptions, new TimeSpan(0, 5, 0)));
        }

        /// <summary>
        /// This function closes the google  and googl driver
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <returns></returns>
        public bool close(IWebDriver p_IWebDriver)
        {
            //Exit the  Browser 
            p_IWebDriver.Close();
            //Quit the  driver
            p_IWebDriver.Quit();
            return true;
        }

	#endregion

    #region Get/Set Values
        public void SetText(IWebDriver p_IWebDriver, string p_strControlID, string p_strText)
        {
            try
            {
                p_IWebDriver.FindElement(By.Id(p_strControlID)).Clear();
                p_IWebDriver.FindElement(By.Id(p_strControlID)).SendKeys(p_strText);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string GetText(IWebDriver p_IWebDriver, string p_strControlID)
        {
            IWebElement iwebEle;
            try
            {
                iwebEle = p_IWebDriver.FindElement(By.Id(p_strControlID));
                string strValue = iwebEle.GetAttribute("Value");
                if (strValue == null) strValue = "";
                return strValue;
            }
            catch (Exception)
            {
                throw;
            }

        }

        public void SetDate(IWebDriver p_IWebDriver, DateTime p_dtValue, string p_strHours, string p_strMinutes, string p_strYearControlID, string p_strMonthControlID, string p_strDayControlID, string p_strHoursControlID, string p_strMinutesControlID)
        {
            SetText(p_IWebDriver, p_strYearControlID, p_dtValue.Year.ToString());
            SetText(p_IWebDriver, p_strMonthControlID, stringMakeTwoDigitString(p_dtValue.Month.ToString()));
            SetText(p_IWebDriver, p_strDayControlID, stringMakeTwoDigitString(p_dtValue.Day.ToString()));
            SetText(p_IWebDriver, p_strHoursControlID, p_strHours);
            SetText(p_IWebDriver, p_strMinutesControlID, p_strMinutes);
        }

        public DateTime GetDate(IWebDriver p_IWebDriver, string p_strYearControlID, string p_strMonthControlID, string p_strDayControlID, string p_strHoursControlID, string p_strMinutesControlID)
        {
            //Get all Values, then combine, then convert into date then return....
            DateTime dtReturnValue;
            string strDay, strMonth, strYear, strHour, strMinute;
            strYear = GetText(p_IWebDriver, p_strYearControlID);
            strMonth = GetText(p_IWebDriver, p_strMonthControlID);
            strDay = GetText(p_IWebDriver, p_strDayControlID);

            strHour = GetText(p_IWebDriver, p_strHoursControlID);
            strMinute = GetText(p_IWebDriver, p_strMinutesControlID);

            dtReturnValue = Convert.ToDateTime(strYear + "/" + strMonth + "/" + strDay + " " + strHour + ":" + strMinute);
            return dtReturnValue;
        }

        public void SelectElement_SelectText(IWebDriver p_IWebDriver, string p_strControlID, string p_strText)
        {
            SelectElement selectControl = new SelectElement(p_IWebDriver.FindElement(By.Id(p_strControlID)));
            //For Sea
            if (selectControl.SelectedOption.Text != p_strText)
                selectControl.SelectByText(p_strText);
        }

        /// <summary>
        /// This function is to select the item in the combo box
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <param name="strControlID"></param>
        /// <param name="strSelectText"></param>
        public void SelectItem(IWebDriver p_IWebDriver, string strControlID, string strSelectText)
        {
            SelectElement SelectElement = new SelectElement(p_IWebDriver.FindElement(By.Id(strControlID))); 
            //Select the option
            if (SelectElement.SelectedOption.Text != strSelectText)
                SelectElement.SelectByText(strSelectText);

            //Validation
            if (SelectElement.SelectedOption.Text != strSelectText)
            {
                throw new Exception("Validation failed for selecting the text:" + strSelectText + " for Control: " + strControlID);
            }
        }

        /// <summary>
        /// This function will return two letter string, if it has only one letter then it will add "0" at the beginning.
        /// </summary>
        /// <param name="p_strText"></param>
        /// <returns></returns>
        private string stringMakeTwoDigitString(string p_strText)
        {
            if (p_strText == null) return "";
            if (p_strText.Length == 1) p_strText = "0" + p_strText;

            return p_strText;
        }

        #endregion

    #region Table Works

        public List<string> GetTableSingleColumnValues(IWebDriver p_IWebDriver, int intCon1ColumnNo, string p_strTableXpath, string p_strTableID = "")
        {
            List<string> lstColumnValues = new List<string>();

            IWebElement tblData = null;

            try
            {
                if (p_strTableXpath != "")
                    tblData = p_IWebDriver.FindElement(By.XPath(p_strTableXpath));
                else
                    tblData = p_IWebDriver.FindElement(By.Id(p_strTableID));

            }
            catch (Exception)
            {
                return lstColumnValues;
            }

            ReadOnlyCollection<IWebElement> rwEvents = tblData.FindElements(By.TagName("tr"));
            if (rwEvents != null)
            {
                lstColumnValues = new List<string>();
                string strCellValue;
                ReadOnlyCollection<IWebElement> cellsEvents = null;
                //Select each item in the table            
                for (int i = 0; i < rwEvents.Count; i++)
                {
                    cellsEvents = rwEvents[i].FindElements(By.TagName("td"));
                    if (cellsEvents.Count == 0) continue;
                    strCellValue = cellsEvents[intCon1ColumnNo].Text;
                    if (strCellValue == null) strCellValue = "";
                    lstColumnValues.Add(strCellValue);
                }
            }


            return lstColumnValues;
        }

        public ReadOnlyCollection<IWebElement> GetTableRow(IWebDriver p_IWebDriver, string p_strTableXpath, int intCon1ColumnNo, string strCon1MatchValue, int intCon2ColumnNo, string strCon2MatchValue)
        {
            IWebElement tblEvents = p_IWebDriver.FindElement(By.XPath(p_strTableXpath));
            ReadOnlyCollection<IWebElement> rwEvents = tblEvents.FindElements(By.TagName("tr"));
            if (rwEvents == null)
            {
                return null;
            }

            ReadOnlyCollection<IWebElement> cellsEvents = null;
            //Select each item in the table            
            for (int i = 0; i < rwEvents.Count; i++)
            {
                cellsEvents = rwEvents[i].FindElements(By.TagName("td"));
                if (cellsEvents == null) continue;
                if (cellsEvents.Count == 0) continue;
                //Check whether its RP Route plan or not.
                if (cellsEvents[intCon1ColumnNo].Text.ToLower() == strCon1MatchValue.ToLower() && cellsEvents[intCon2ColumnNo].Text.ToLower() == strCon2MatchValue.ToLower())
                {
                    return cellsEvents;
                }
            }

            return null;
        }

        /// <summary>
        /// This function will return the datatable form the input
        /// HeaderTableXPath is needed - if the header is in the different table
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="p_strHeaderTableXPath"></param>
        /// <param name="strArrColumns"></param>
        /// <returns></returns>
        public DataTable GetTable(IWebDriver p_IWebDriver, string p_strTableIDorXPath, string p_strHeaderTableXPath = "", string[] strArrColumns = null)
        {
            DataTable tblResult = new DataTable();
            //Maximize the window to get the tables clearly
            try { p_IWebDriver.Manage().Window.Maximize(); }
            catch { };

            int intFirstRow = 1;
            try
            {
                IWebElement tblData = null;
                //Find whether its ID or Xpath
                if (p_strTableIDorXPath.Contains("/") == false && p_strTableIDorXPath.Contains("[") == false)
                {
                    //Convert ID into xpath
                    //tblData = p_IWebDriver.FindElement(By.Id(p_strTableIDorXPath));     
                    p_strTableIDorXPath = "//*[@id=\"" + p_strTableIDorXPath + "\"]";
                    Console.WriteLine("Converted table-ID into xpath is :" + p_strHeaderTableXPath);
                }

                ////Wait until the browser busy
                //WebDriverWait wait = new WebDriverWait(p_IWebDriver, new TimeSpan(0, 1, 0));
                //tblData = wait.Until<OpenQA.Selenium.IWebElement>(ExpectedConditions.ElementExists(OpenQA.Selenium.By.XPath(p_strTableIDorXPath)));
                //tblData = wait.Until<OpenQA.Selenium.IWebElement>(ExpectedConditions.ElementToBeClickable(OpenQA.Selenium.By.XPath(p_strTableIDorXPath)));
                
                tblData = p_IWebDriver.FindElement(By.XPath(p_strTableIDorXPath));
                ReadOnlyCollection<IWebElement> rwData = tblData.FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"));
                if (rwData == null)
                {
                    return null;
                }

                ReadOnlyCollection<IWebElement> headerCells = null;
                ReadOnlyCollection<IWebElement> cellsEvents = null;
                List<string> lstColumns = null;

                //Convert Array into list.....
                if (strArrColumns != null) lstColumns = strArrColumns.ToList<string>();

                //Find whether it has the thead tag...
                try
                {
                    if (tblData.FindElements(By.TagName("thead")).Count > 0)
                    {
                        p_strHeaderTableXPath = p_strTableIDorXPath + "/thead/tr";
                        Console.WriteLine("found the thead, the xpath is :" + p_strHeaderTableXPath);
                    }
                }
                catch 
                {
            
                }

                //Find the header row
                if (p_strHeaderTableXPath == "")
                {
                    headerCells = rwData[0].FindElements(By.TagName("td"));
                }
                else
                {
                    IWebElement tblHeader = p_IWebDriver.FindElement(By.XPath(p_strHeaderTableXPath));
                    headerCells = tblHeader.FindElements(By.TagName("td"));
                    if (headerCells.Count == 0) headerCells = tblHeader.FindElements(By.TagName("th"));
                    intFirstRow = 0;
                }

                //Add columns
                for (int i = 0; i < headerCells.Count; i++)
                {
                    if (lstColumns == null || lstColumns.Contains(headerCells[i].Text.Trim()) == true)
                        tblResult.Columns.Add(headerCells[i].Text.Trim(), typeof(string));
                }

                tblResult.Columns.Add("RowDblClickValue", typeof(string));

                //Add the values
                //Select each item in the table
                for (int i = intFirstRow; i < rwData.Count; i++)
                {
                    DataRow drData = tblResult.NewRow();

                    //Store the attribute value
                    string strKeyValue = rwData[i].GetAttribute("ondblclick");
                    if (strKeyValue != null && strKeyValue != "" && strKeyValue.IndexOf('\'') != 0)
                    {
                        strKeyValue = strKeyValue.Split(new char[] { '\'' })[1];
                        drData["RowDblClickValue"] = strKeyValue;
                    }

                    //Check the columns
                    cellsEvents = rwData[i].FindElements(By.TagName("td"));
                    if (cellsEvents != null || cellsEvents.Count > 0)
                    {
                        for (int j = 0; j < cellsEvents.Count; j++)
                        {
                            if (lstColumns == null || lstColumns.Contains(headerCells[j].Text) == true)
                                drData[headerCells[j].Text.Trim()] = cellsEvents[j].Text.Trim();
                        }
                    }

                    tblResult.Rows.Add(drData);
                }

            }
            catch
            {
 
                throw;
            }

            return tblResult;
        }

        /// <summary>
        /// This function will add a new tab using javascript function "window.open"
        /// Make sure JavaScript is enabled in the browser
        /// </summary>
        /// <param name="p_Browser"></param>
        /// <returns></returns>
        public string AddNewTab(IWebDriver p_Browser)
        {
            //Open the new window with details
            ////Method1
            //p_Browser.SwitchTo().DefaultContent();
            //IWebElement bodyele = p_Browser.FindElement(By.CssSelector("body"));
            //bodyele.Click();
            //bodyele.SendKeys(Keys.Control + "t");
            //System.Threading.Thread.Sleep(1000);

            ////Method2
            //p_Browser.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "T");
            //bodyele.SendKeys(Keys.Control + "T");

            //Method 3
            //Using javascript
            //Get the previous handles
            IEnumerable<string> IEnumWindows = new List<string>();
            IEnumWindows = p_Browser.WindowHandles;
            ((IJavaScriptExecutor)p_Browser).ExecuteScript("window.open('about:blank','_blank');");
            //Find the new tab name and return the same            
            return p_Browser.WindowHandles.Except(IEnumWindows).FirstOrDefault().ToString();

            ////Method 2
            //Actions actDClick = new Actions(p_chromeDriver);
            ////actDClick.KeyDown(Keys.Control).MoveToElement(bodyele).Click().Perform();
            //actDClick.KeyDown(Keys.Control).SendKeys("t").KeyUp(Keys.Control).Build().Perform();
            //actDClick.Perform();

            ////act.keyDown(Keys.CONTROL).sendKeys("t").keyUp(Keys.CONTROL).build().p erform(); //Open new tab          
            ////string strUrl = "http://" + Environment.UserName + ":" + p_strWinPwd + "@scp.panalpina.com/XCP/";
            ////String selectLinkOpeninNewTab = Keys.CONTROL, Keys.RETURN);

            //p_chromeDriver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
            //p_chromeDriver.SwitchTo().Window(p_chromeDriver.WindowHandles.Last());
            //p_chromeDriver.Navigate().GoToUrl("http://www.google.com");

            ////p_chromeDriver.FindElement(By.TagName("body")).SendKeys(Keys.Control + "t");
            //p_chromeDriver.Navigate().GoToUrl(strUrl);

            ////driver.findElement(By.cssSelector("body")).sendKeys(selectLinkOpeninNewTab);

            //if (chromeSwitchToPage(p_chromeDriver, "") == false)
            //{
            //    //Code to quit
            //    return null;
            //}

            ////Open  new tab window and request for the reading details
        }

        /// <summary>
        /// This function search a specific row in a table and returns the matched row
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="p_strColumn1Name"></param>
        /// <param name="p_strCon1MatchValue"></param>
        /// <param name="p_strColumn2Name"></param>
        /// <param name="strCon2MatchValue"></param>
        /// <returns>Return type is ReadOnlyCollection-Add System.Collections.ObjectModel in using statement </returns>
        public IWebElement getTableRow(IWebDriver p_IWebDriver, string p_strTableXpath, string p_strColumn1Name, string p_strCon1MatchValue, string p_strColumn2Name, string strCon2MatchValue)
        {
            IWebElement tblEvents = p_IWebDriver.FindElement(By.XPath(p_strTableXpath));
            ReadOnlyCollection<IWebElement> rwEvents = tblEvents.FindElements(By.TagName("tr"));
            int intColumn1Index = -1;
            int intColumn2Index = -1;
            if (rwEvents == null)
            {
                return null;
            }

            ReadOnlyCollection<IWebElement> cellsEvents = null;
            //Match the title and get the column index
            cellsEvents = rwEvents[0].FindElements(By.TagName("td"));
            for (int i = 0; i < cellsEvents.Count; i++)
            {
                if (cellsEvents[i].Text.Trim().ToLower() == p_strColumn1Name.ToLower())
                    intColumn1Index = i;
                else if (cellsEvents[i].Text.Trim().ToLower() == p_strColumn2Name.ToLower())
                    intColumn2Index = i;
            }

            //Check both index has been found, if not found then return null
            if (intColumn1Index == -1 || intColumn2Index == -1)
                return null;

            //Select each item in the table
            for (int i = 1; i < rwEvents.Count; i++)
            {
                cellsEvents = rwEvents[i].FindElements(By.TagName("td"));
                if (cellsEvents == null) continue;
                if (cellsEvents.Count == 0) continue;

                if (cellsEvents[intColumn1Index].Text.Trim().ToLower() == p_strCon1MatchValue.ToLower() && cellsEvents[intColumn2Index].Text.Trim().ToLower() == strCon2MatchValue.ToLower())
                {
                    return rwEvents[i];
                }
            }

            return null;
        }

        //public DataTable HTMLGetTable(HtmlAgilityPack.HtmlNode p_htmlTable, string p_strTableXPath = null, string[] strArrColumns = null)
        //{
        //    DataTable tblResult = new DataTable();
        //    try
        //    {
        //        //HtmlAgilityPack.HtmlNode tblEvents = p_IWebDriver.FindElement(By.XPath(p_strTableXPath));
        //        HtmlAgilityPack.HtmlNodeCollection htmlRows = p_htmlTable.SelectNodes("tr");
        //        //if (rwEvents == null)
        //        //{
        //        //    return null;
        //        //}

        //        HtmlAgilityPack.HtmlNodeCollection headerCells = null;
        //        HtmlAgilityPack.HtmlNodeCollection cellsEvents = null;

        //        List<string> lstColumns = null;

        //        //Convert Array into list.....
        //        if (strArrColumns != null) lstColumns = strArrColumns.ToList<string>();

        //        //Add columns
        //        headerCells = htmlRows[0].SelectNodes("td");
        //        for (int i = 0; i < headerCells.Count; i++)
        //        {
        //            if (lstColumns == null || lstColumns.Contains(headerCells[i].InnerText.Trim()) == true)
        //                tblResult.Columns.Add(headerCells[i].InnerText, typeof(string));
        //        }

        //        //Add the values
        //        //Select each item in the table            
        //        for (int i = 1; i < htmlRows.Count; i++)
        //        {
        //            DataRow drData = tblResult.NewRow();
        //            //Check the columns
        //            cellsEvents = htmlRows[i].SelectNodes("td");
        //            if (cellsEvents != null || cellsEvents.Count > 0)
        //            {
        //                for (int j = 0; j < cellsEvents.Count; j++)
        //                {
        //                    if (lstColumns == null || lstColumns.Contains(headerCells[j].InnerText) == true)
        //                        drData[headerCells[j].InnerText] = cellsEvents[j].InnerText.Trim();
        //                }
        //            }

        //            tblResult.Rows.Add(drData);
        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        e = e;
        //        throw;
        //    }

        //    return tblResult;
        //}        

        /// <summary>
        /// This function will return the datatable form the input
        /// HeaderTableXPath is needed - if the header is in the different table
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="p_strHeaderTableXPath"></param>
        /// <param name="strArrColumns">columd can be specified with type or without... For example - with data type: "int|ID, without data type: "ID"</param>
        /// <returns></returns>
        public DataTable GetTableWithDataType(IWebDriver p_IWebDriver, string p_strTableIDorXPath, string p_strHeaderTableXPath = "", string[] strArrColumnsWithDataType = null)
        {
            DataTable dataTable = new DataTable();
            Dictionary<string, string> dictionary = (Dictionary<string, string>)null;
            try
            {
                p_IWebDriver.Manage().Window.Maximize();
            }
            catch
            {
            }
            int num = 1;
            try
            {
                if (!p_strTableIDorXPath.Contains("/") && !p_strTableIDorXPath.Contains("["))
                {
                    p_strTableIDorXPath = "//*[@id=\"" + p_strTableIDorXPath + "\"]";
                    Console.WriteLine("Converted table-ID into xpath is :" + p_strHeaderTableXPath);
                }
                IWebElement element1 = p_IWebDriver.FindElement(By.XPath(p_strTableIDorXPath));
                ReadOnlyCollection<IWebElement> elements1 = element1.FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"));
                if (elements1 == null)
                    return (DataTable)null;
                if (strArrColumnsWithDataType != null)
                {
                    dictionary = new Dictionary<string, string>();
                    foreach (string key in ((IEnumerable<string>)strArrColumnsWithDataType).ToList<string>())
                    {
                        if (key.Contains("|"))
                            dictionary.Add(key.Substring(key.IndexOf("|") + 1), key.Substring(0, key.IndexOf("|")).ToLower());
                        else
                            dictionary.Add(key, "string");
                    }
                }
                try
                {
                    if (element1.FindElements(By.TagName("thead")).Count > 0)
                    {
                        p_strHeaderTableXPath = p_strTableIDorXPath + "/thead/tr";
                        Console.WriteLine("found the thead, the xpath is :" + p_strHeaderTableXPath);
                    }
                }
                catch 
                {
                }
                ReadOnlyCollection<IWebElement> elements2;
                if (p_strHeaderTableXPath == "")
                {
                    elements2 = elements1[0].FindElements(By.TagName("td"));
                }
                else
                {
                    IWebElement element2 = p_IWebDriver.FindElement(By.XPath(p_strHeaderTableXPath));
                    elements2 = element2.FindElements(By.TagName("td"));
                    if (elements2.Count == 0)
                        elements2 = element2.FindElements(By.TagName("th"));
                    num = 0;
                }
                for (int index = 0; index < elements2.Count; ++index)
                {
                    if (dictionary == null)
                        dataTable.Columns.Add(elements2[index].Text.Trim(), typeof(string));
                    else if (dictionary.ContainsKey(elements2[index].Text.Trim()))
                    {
                        string str = dictionary[elements2[index].Text.Trim()];
                        if (str == "int")
                            dataTable.Columns.Add(elements2[index].Text.Trim(), typeof(int));
                        else if (str == "string")
                            dataTable.Columns.Add(elements2[index].Text.Trim(), typeof(string));
                        else if (str == "date")
                            dataTable.Columns.Add(elements2[index].Text.Trim(), typeof(DateTime));
                        else if (str == "bool" || str == "boolean")
                            dataTable.Columns.Add(elements2[index].Text.Trim(), typeof(bool));
                        else
                            dataTable.Columns.Add(elements2[index].Text.Trim(), typeof(string));
                    }
                }
                dataTable.Columns.Add("RowDblClickValue", typeof(string));
                for (int index1 = num; index1 < elements1.Count; ++index1)
                {
                    DataRow row = dataTable.NewRow();
                    string attribute = elements1[index1].GetAttribute("ondblclick");
                    if (attribute != null && attribute != "" && attribute.IndexOf('\'') != 0)
                    {
                        string str = attribute.Split('\'')[1];
                        row["RowDblClickValue"] = (object)str;
                    }
                    ReadOnlyCollection<IWebElement> elements3 = elements1[index1].FindElements(By.TagName("td"));
                    if (elements3 != null || elements3.Count > 0)
                    {
                        for (int index2 = 0; index2 < elements3.Count; ++index2)
                        {
                            if (dictionary == null || dictionary.ContainsKey(elements2[index2].Text.Trim()))
                                row[elements2[index2].Text.Trim()] = (object)elements3[index2].Text.Trim();
                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }
            catch 
            {
                throw;
            }
            return dataTable;
        }


    #endregion
        
    #region Alert Window Handling

        /// <summary>
        /// This function is to wait until an alert display, then accept the same alert, may click okay, yes. This function will wait for 50 seconds/based on the input parameter to display the alert if not then it will return false        
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <param name="p_strAlertTitle"> Title of the alert window</param>
        /// <param name="p_strAlertContentContains">Alert window text, this will be used for contains match</param>
        /// <returns></returns>
        public bool AcceptAlertWindow(IWebDriver p_IWebDriver, string p_strAlertContentContains, int intWaitingSeconds = 50)
        {
            //Wait till popup comes - wait for 50 seconds/based on the input
            WebDriverWait webWait = new WebDriverWait(p_IWebDriver, new TimeSpan(0, 0, intWaitingSeconds));

            try
            {
                webWait.Until(ExpectedConditions.AlertIsPresent());
            }
            catch (WebDriverTimeoutException) //Alert window not found
            {
                return false;
            }

            //Check whether the alert is present if not then show error
            try
            {
                //Check the alert title and content is same as the parameter
                if (p_IWebDriver.SwitchTo().Alert().Text.Contains(p_strAlertContentContains) == true)
                {
                    //Click do not show again
                    if (p_strAlertContentContains == "Please select Shipment Creation Application")
                    {
                        p_IWebDriver.SwitchTo().Alert().SendKeys(OpenQA.Selenium.Keys.Space);
                    }
                    //Accept the alert  //Equal to click Ok /Okay / Yes
                    p_IWebDriver.SwitchTo().Alert().Accept();
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (NoAlertPresentException)
            {
                //return false - no alert found in 300 seconds
                return false;
            }
        }        

    #endregion

    #region Find Element

        /// <summary>
        /// This function is to find the element , if the  driver throws the error then null will be returned
        /// </summary>
        /// <param name="p_IWebDriver"></param>
        /// <param name="p_strElementID"></param>
        /// <returns></returns>
        public IWebElement FindElementByID(IWebDriver p_IWebDriver, string p_strElementID)
        {
            try
            {
                return p_IWebDriver.FindElement(By.Id(p_strElementID));
            }
            catch (Exception)
            {
                return null;
            }
        }        

    #endregion
        
    }
}
