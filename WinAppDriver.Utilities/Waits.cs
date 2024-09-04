using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static WinAppDriver.Utilities.ElementFactory;
using Assert = NUnit.Framework.Assert;

namespace WinAppDriver.Utilities
{
    public class Waits
    {
        private WindowsDriver<WindowsElement> _session;
        string Implicit_wait_time = ConfigurationManager.AppSettings["Implicit_wait_time"];
        static WindowsUIActions windowsUIActions;

        public Waits(WindowsDriver<WindowsElement> session)
        {
            _session = session;
            windowsUIActions = new WindowsUIActions(_session);
        }


        public void Implicit_Wait(double? time = null)
        {

            if (time.Equals(null))
                _session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(double.Parse(Implicit_wait_time));
            else
                _session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds((double)time);
        }
        public void Implicit_Wait(WindowsDriver<WindowsElement> _session, double? time = null)
        {

            if (time.Equals(null))
                _session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(double.Parse(Implicit_wait_time));
            else
                _session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds((double)time);
        }
        public void wait_new_window(int number_of_windows)
        {
            try
            {
                Thread.Sleep(1000);
                int itr = 1;
                var allowwindowshandler = _session.WindowHandles;
                int cnt = allowwindowshandler.Count;
                while (cnt < number_of_windows)
                {
                    allowwindowshandler = _session.WindowHandles;
                    cnt = allowwindowshandler.Count;
                    Thread.Sleep(1000);
                    itr++;
                    if (itr == 60)
                        break;
                }
                _session.SwitchTo().Window(allowwindowshandler[0]);
            }
            catch { }
        }
        public void wait_new_window(int number_of_windows, WindowsDriver<WindowsElement> _session, int numOfSeconds)
        {
            try
            {
                Thread.Sleep(1000);
                int itr = 1;
                var allowwindowshandler = _session.WindowHandles;
                int cnt = allowwindowshandler.Count;
                while (cnt < number_of_windows)
                {
                    allowwindowshandler = _session.WindowHandles;
                    cnt = allowwindowshandler.Count;
                    Thread.Sleep(1000);
                    itr++;
                    if (itr == numOfSeconds)
                        break;
                }
                _session.SwitchTo().Window(allowwindowshandler[0]);
            }
            catch { }
        }
        public static bool WaitAndCheckTillControlToDisplay(WindowsElement element, int Waittime = 120)
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
      
        public void WaitTillControlToDisplay(WindowsElement element, int Waittime = 120)
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


        public bool IsUIElementPresent(WindowsUIElement element, bool checkException = false)
        {
            try
            {
                Implicit_Wait(_session, 0);
                bool Check_element = windowsUIActions.checkEnabled(element,false,checkException);
                Implicit_Wait(_session);
                return true;
            }
            catch
            {
                Implicit_Wait(_session);
                return false;
            }
        }
        public void WaitTillToUIElementPresent(WindowsUIElement element, int Waittime = 120)
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
            WaitTillToUIElementPresent(new WindowsUIElement(By.Name, headerName));
            count = windowsUIActions.getGridColumnValues(new WindowsUIElement(By.Name, headerName), 0).ToList().Count;
            while (count <= countRowsToWait)
            {
                Thread.Sleep(1000);
                try
                {
                    WaitTillToUIElementPresent(new WindowsUIElement(By.Name, headerName));
                    count = windowsUIActions.getGridColumnValues(new WindowsUIElement(By.Name, headerName), 0).ToList().Count;
                }
                catch { }
                itr++;
                if (itr == 320)
                    break;
            }
        }

    }
}
