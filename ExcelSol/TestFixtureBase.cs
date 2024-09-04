using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium.Appium.Windows;
using WinAppDriver.Utilities;

namespace ExcelSol
{
    public class TestFixtureBase
    {
        public static WindowsDriver<WindowsElement> _session;
        driverActions driverActions;
        DriverFactory driverFactory = new DriverFactory();
        Waits waits;

        public string WinDriver_Path = ConfigurationManager.AppSettings["WinDriver_Path"];
    }
}
