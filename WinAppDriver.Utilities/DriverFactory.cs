using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
 

namespace WinAppDriver.Utilities
{
    public class DriverFactory 
    {
        public static string Application_Path = ConfigurationManager.AppSettings["Appliction_Path"];
        public static string Implicit_wait_time = ConfigurationManager.AppSettings["Implicit_wait_time"];
        public WindowsDriver<WindowsElement> InitializeDriver(string deviceName, Uri uri, string applicationPath = null)
        {
           Run_WinappDriver();
            AppiumOptions opt = new AppiumOptions();

            opt.AddAdditionalCapability("app", applicationPath == null ? @Application_Path : applicationPath);
            opt.AddAdditionalCapability("diviceName", deviceName);
            opt.AddAdditionalCapability("platformName", "windows");
            opt.AddAdditionalCapability("ms:waitForAppLaunch", "5");
           
              return new WindowsDriver<WindowsElement>(uri, opt,TimeSpan.FromSeconds(int.Parse(Implicit_wait_time)));
            
        }
        public void InitializeDriver_ConsoleApp(string deviceName, Uri uri, string applicationPath = null)
        {
            Run_WinappDriver();
            AppiumOptions opt = new AppiumOptions();

            opt.AddAdditionalCapability("app", applicationPath == null ? @Application_Path : applicationPath);
            opt.AddAdditionalCapability("diviceName", deviceName);
            opt.AddAdditionalCapability("platformName", "windows");
            opt.AddAdditionalCapability("ms:waitForAppLaunch", "5");

            try
            {
                 new WindowsDriver<WindowsElement>(uri, opt);
            }
            catch
            { }
        }
        public WindowsDriver<WindowsElement> InitializeRunningProcess(string appTopLevelWindowHandleHex, string Device_Name, Uri uri)
        {          
            AppiumOptions opt = new AppiumOptions();
            opt.AddAdditionalCapability("diviceName", Device_Name);
            opt.AddAdditionalCapability("platformName", "windows");
            opt.AddAdditionalCapability("appTopLevelWindow", appTopLevelWindowHandleHex);
            return  new WindowsDriver<WindowsElement>(uri, opt);
        }

        public static void Run_WinappDriver()
        {

            string WinDriver_Path = ConfigurationManager.AppSettings["WinDriver_Path"];
            System.Diagnostics.Process.Start(@WinDriver_Path);
        }
        public static void Run_WindowsProcesses_EXE()
        {

            string WindowsProcesses_EXE_Path = ConfigurationManager.AppSettings["WindowsProcesses_EXE_Path"];
            //System.Diagnostics.Process.Start(@WindowsProcesses_EXE_Path);
     
        }
        public void TestCleanup(WindowsDriver<WindowsElement> _session)
        {

            if (_session!= null)
            {
                _session.Quit();
                _session = null;

            }
        }

    }
}
