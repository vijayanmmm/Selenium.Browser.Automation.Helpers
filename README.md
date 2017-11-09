# Selenium.Browser.Automation.Helpers
This project to provide an easy way to start automating browser using selenium. It has set of methods inbuild which can help to automate quickly.

# Instructions
1. Add the nuget package
2. Depends on which browser you gonna use, download and include the driver "chrome driver/ IE Driver/ Opera Driver" in your project. You can download these from the webpage [http://www.seleniumhq.org/download]" - listed under "Third Party Drivers, Bindings, and Plugins".
3. Follow the sample code.

# Samples - C#
## NamesSpace
```c#
using BrowserAutomation;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions; //For the actions like Double click and item in the webpage
```
## Data Filling(Only text) & Click Button
The following code to open the facebook web page, then enter user name and password.
```c#
            clsBrowserAutomation objBrowser = new clsBrowserAutomation();
            IWebDriver chromeBrowser = objBrowser.openChrome();
            //Open the website
            chromeBrowser.Navigate().GoToUrl("https://www.facebook.com/");
            //Enter the username and password
            objBrowser.SetText(chromeBrowser, "email", "sample@domain.com");
            objBrowser.SetText(chromeBrowser, "pass", "password");
            //Click the login button
            objBrowser.FindElementByID(chromeBrowser,"u_0_5").Click();
```

## Data Filling(text & options) & Click Button
The following code to open the facebook web page, then fill the signup details.

```c#
            clsBrowserAutomation objBrowser = new clsBrowserAutomation();
            IWebDriver chromeBrowser = objBrowser.openChrome();
            //Open the website
            chromeBrowser.Navigate().GoToUrl("https://www.facebook.com/");
            //Enter Signup details
            //First name
            objBrowser.SetText(chromeBrowser, "u_0_g", "Vijayan");
            //LastName
            objBrowser.SetText(chromeBrowser, "u_0_i", "Venkateshan");
            //Email
            objBrowser.SetText(chromeBrowser, "u_0_1", "sample@domain.com");
            //Date of Birth
            //If the input box is not combo box
            //objBrowser.SetDate(chromeBrowser, new DateTime(2017, 11, 10), "00", "00","year","month","day", "minutes", "hours");
            //Set Value in day, month, year combo box
            objBrowser.SelectItem(chromeBrowser, "day", "10");
            objBrowser.SelectItem(chromeBrowser, "month", "Nov");
            objBrowser.SelectItem(chromeBrowser, "year", "2017");
            //Select the option Male/Female - same code as button click
            objBrowser.FindElementByID(chromeBrowser, "u_0_6").Click();
            //Click the Signup button
            objBrowser.FindElementByID(chromeBrowser, "u_0_y").Click();
  ```
  
