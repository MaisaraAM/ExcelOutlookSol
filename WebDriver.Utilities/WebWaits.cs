using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static WebDriver.Utilities.ElementFactory;
using By = WebDriver.Utilities.ElementFactory.by;
using Assert = NUnit.Framework.Assert;

namespace WebDriver.Utilities
{
    public class WebWaits
    {
        private static IWebDriver _WebDriver;
        static WebUIActions WebUIActions;
        public static string Implicit_wait_time = "60";
        // public string Implicit_wait_time = "60";
        public WebWaits(IWebDriver webDriver)
        {
            _WebDriver = webDriver;
            WebUIActions = new WebUIActions(_WebDriver);
        }


        public static void Implicit_Wait(double? time = null)
        {

            if (time.Equals(null))
                _WebDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(double.Parse(Implicit_wait_time));
            else
                _WebDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds((double)time);
        }
        public static void Implicit_Wait(IWebDriver _WebDriver, double? time = null)
        {

            if (time.Equals(null))
                _WebDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(double.Parse(Implicit_wait_time));
            else
                _WebDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds((double)time);
        }

        public static bool WaitAndCheckTillControlToDisplay(IWebElement element, int Waittime = 120)
        {
            bool checkDisplayed = false;
            DateTime dt = DateTime.Now;
            do
            {
                if (element.Displayed)
                {
                    checkDisplayed = true;
                    break;
                }
            }
            while (dt.AddSeconds(Waittime) > DateTime.Now);
            return checkDisplayed;
        }

        public void WaitTillControlToDisplay(IWebElement element, int Waittime = 120)
        {
            DateTime dt = DateTime.Now;
            do
            {
                if (element.Displayed)
                {
                    return;
                }
            }
            while (dt.AddSeconds(Waittime) > DateTime.Now);
            Assert.Fail("Time Out : Control - " + element + " Did not loaded within " + Waittime + " Seconds");
        }


        public bool IsUIElementPresent(WebUIElement element,bool checkException=false)
        {
            try
            {
                Implicit_Wait(_WebDriver, 0);
                bool Check_element = WebUIActions.checkDisplayed(element, false, checkException);
                Implicit_Wait(_WebDriver);
                return true;
            }
            catch
            {
                Implicit_Wait(_WebDriver);
                return false;
            }
        }
        public void WaitTillToUIElementPresent(WebUIElement element, int Waittime = 120)
        {
            DateTime dt = DateTime.Now;
            do
            {
                if (IsUIElementPresent(element))
                {
                    return;
                }
            }
            while (dt.AddSeconds(Waittime) > DateTime.Now);
        }
        public void waitingGridToRowsDipalyed(string headerName, int countRowsToWait = 1)
        {
            int count, itr = 1;
            WaitTillToUIElementPresent(new WebUIElement(By.Name, headerName));
            count = WebUIActions.getGridColumnValues(new WebUIElement(by.Name, headerName), 0).ToList().Count;
            while (count <= countRowsToWait)
            {
                Thread.Sleep(1000);
                try
                {
                    WaitTillToUIElementPresent(new WebUIElement(By.Name, headerName));
                    count = WebUIActions.getGridColumnValues(new WebUIElement(By.Name, headerName), 0).ToList().Count;
                }
                catch { }
                itr++;
                if (itr == 180)
                    break;
            }
        }

    }
}
