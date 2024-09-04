using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Threading;
using System.Web.SessionState;

namespace WebDriver.Utilities
{
    public class WebDriverFactory
    {
        public IWebDriver _webdriver;
        

        public IWebDriver InitializeDriver(string Url)
        {       
            var options = new ChromeOptions();
            options.AddArguments("disable-infobars");
            options.AddArguments("--start-maximized");
            options.AddArgument("--disable-features=VizDisplayCompositor");
            options.AddAdditionalCapability("useAutomationExtension", false);
            options.AddExcludedArgument("enable-automation");
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);

            options.AddArgument("--ignore-ssl-errors=yes");
            options.AddArgument("--ignore-certificate-errors");

            _webdriver = new ChromeDriver(options);
            _webdriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
            _webdriver.Manage().Window.Maximize();
            _webdriver.Navigate().GoToUrl(Url);
            return _webdriver;
        }

        public void CloseBrowser(IWebDriver _webdriver)
        {

            if (_webdriver != null)
            {
                _webdriver.Quit();
                _webdriver = null;
            }
        }


    }
}
