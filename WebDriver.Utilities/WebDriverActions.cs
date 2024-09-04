using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keys = WebDriver.Utilities.Keys;

namespace WebDriver.Utilities
{
    public class WebDriverActions
    {
        private IWebDriver _webdriver;
        ElementFactory elementFactory;
        WebUIActions webUIActions;

        public WebDriverActions(IWebDriver session)
        {
            _webdriver = session;
        }

        #region webActions
        public void maximizeWindow()
        {
            _webdriver.Manage().Window.Maximize();
        }
        public void switchTo(string window)
        {
            _webdriver.SwitchTo().Window(window);
        }
        public List<string> windowHandlesList()
        {
            return _webdriver.WindowHandles.ToList();
        }
        public void closeWindow()
        {
            _webdriver.Close();
        }
        public void closeWindow(WebUIElement ele)
        {

            webUIActions.sendKeyboardKeys(ele, Keys.Alt + Keys.F4);
        }
        public void Screenshot(string screen_name, IWebDriver _webdriver,out string screenShotPath)
        {
            screenShotPath = "";
            try
            {
                if (_webdriver != null)
                {
                    Screenshot scr_sh = ((ITakesScreenshot)_webdriver).GetScreenshot();
                    screen_name = screen_name + System.DateTime.Now.ToString("yyyyMMddHHmmss");
                    var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
                    var applicationPath_new = application_path.Replace("\\bin\\Debug", "");
                    var applicationPath_new_name = applicationPath_new + "\\Results";
                    screenShotPath = $@"{applicationPath_new_name}\{screen_name}.jpg";
                    scr_sh.SaveAsFile($@"{applicationPath_new_name}\{screen_name}.jpg");
                }
               
            }
            catch { }
        }
        #endregion

    }
}
