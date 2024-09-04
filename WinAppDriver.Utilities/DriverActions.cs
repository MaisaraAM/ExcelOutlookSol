using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keys = WinAppDriver.Utilities.Keys;

namespace WinAppDriver.Utilities
{
   public class driverActions
    {
        private WindowsDriver<WindowsElement> _session;
        ElementFactory elementFactory;
        WindowsUIActions windowsUIActions;

        public driverActions(WindowsDriver<WindowsElement> session)
        {
            _session = session;
        }

        #region widowActions
        public void maximizeWindow()
        {
            _session.Manage().Window.Maximize();
        }
        public void switchTo(string window)
        {
            _session.SwitchTo().Window(window);
        }
        public List<string> windowHandlesList()
        {
            return _session.WindowHandles.ToList();
        }
        public void closeWindow()
        {
            _session.Close();
        }
        public void closeWindow(WindowsUIElement ele)
        {
         
            windowsUIActions.sendKeyboardKeys(ele, Keys.Alt + Keys.F4);
        }
        public void Screenshot(string screen_name, WindowsDriver<WindowsElement> _session)
        {
            try
            {
                Screenshot scr_sh = ((ITakesScreenshot)_session).GetScreenshot();
                screen_name = screen_name + System.DateTime.Now.ToString("yyyyMMddHHmmss");
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
              
                string pathString = System.IO.Path.Combine(appDir.Replace("\\bin\\Debug", ""), "Results");
              
                string fullPath = System.IO.Path.Combine(pathString, screen_name + ".Png");
               
                 scr_sh.SaveAsFile(fullPath);
            }


            catch { }
        }
        public void Screenshot(string screen_name, WindowsDriver<WindowsElement> _session, out string fullPath,out string imageName)
        {
            fullPath = "";
            imageName= "";
            try
            {
                Screenshot scr_sh = ((ITakesScreenshot)_session).GetScreenshot();
                imageName = screen_name + System.DateTime.Now.ToString("yyyyMMddHHmmss");
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string pathString = System.IO.Path.Combine(appDir.Replace("\\bin\\Debug", ""), "Results");
                fullPath = System.IO.Path.Combine(pathString, imageName + ".Png");
                scr_sh.SaveAsFile(fullPath);
            }
            catch { }
        }
        #endregion

    }
}
