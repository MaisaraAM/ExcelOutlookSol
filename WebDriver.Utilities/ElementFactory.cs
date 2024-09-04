using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;


namespace WebDriver.Utilities
{
    public class ElementFactory
    {
        private IWebDriver _webDriver;
        public ElementFactory(IWebDriver webDriver)
        {
            this._webDriver = webDriver;
        }

        public IWebElement findElement(by b, string locator, bool Optional = false,bool checkException=true)
        {
            try
            {
                switch (b)
                {
                    case by.Id:
                        return _webDriver.FindElement(By.Id(locator));
                    case by.Name:
                        return _webDriver.FindElement(By.Name(locator));
                    case by.TagName:
                        return _webDriver.FindElement(By.TagName(locator));
                    case by.Xpath:
                        return _webDriver.FindElement(By.XPath(locator));
                    case by.ClassName:
                        return _webDriver.FindElement(By.ClassName(locator));
                    case by.CssSelector:
                        return _webDriver.FindElement(By.CssSelector(locator));
                    case by.LinkText:
                        return _webDriver.FindElement(By.LinkText(locator));


                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                if (Optional)

                    return null;

                else
                {
                    if (checkException)
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot Find element with locator {0} and here is the exception ", locator));
                    throw e;
                }
            }
        }
        public IWebElement findElement(by b, string locator, IWebElement ele, bool Optional = false)
        {

            try
            {
                switch (b)
                {
                    case by.Id:
                        return ele.FindElement(By.Id(locator));
                    case by.Name:
                        return ele.FindElement(By.Name(locator));
                    case by.TagName:
                        return ele.FindElement(By.TagName(locator));
                    case by.Xpath:
                        return ele.FindElement(By.XPath(locator));
                    case by.ClassName:
                        return ele.FindElement(By.ClassName(locator));
                    case by.CssSelector:
                        return ele.FindElement(By.CssSelector(locator));
                    case by.LinkText:
                        return ele.FindElement(By.LinkText(locator));

                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                if (Optional)

                    return null;

                else
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot Find element with locator {0} and here is the exception ", locator));
                    throw e;
                }
            }
        }

        public IList<IWebElement> findElements(by b, string locator)
        {
            try
            {
                switch (b)
                {
                    case by.Id:
                        return _webDriver.FindElements(By.Id(locator));
                    case by.Name:
                        return _webDriver.FindElements(By.Name(locator));
                    case by.TagName:
                        return _webDriver.FindElements(By.TagName(locator));
                    case by.Xpath:
                        return _webDriver.FindElements(By.XPath(locator));
                    case by.ClassName:
                        return _webDriver.FindElements(By.ClassName(locator));
                    case by.CssSelector:
                        return _webDriver.FindElements(By.CssSelector(locator));
                    case by.LinkText:
                        return _webDriver.FindElements(By.LinkText(locator));

                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot Find element with locator {0} and here is the exception ", locator));
                return null;
            }
        }
        public dynamic findElements(by b, string locator, IWebElement ele)
        {
            try
            {
                switch (b)
                {
                    case by.ClassName:
                        return ele.FindElements(By.ClassName(locator));
                    case by.Name:
                        return ele.FindElements(By.Name(locator));
                    case by.TagName:
                        return ele.FindElements(By.TagName(locator));
                    case by.Xpath:
                        return ele.FindElements(By.XPath(locator));
                    case by.CssSelector:
                        return ele.FindElements(By.CssSelector(locator));
                    case by.LinkText:
                        return ele.FindElements(By.LinkText(locator));
                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot Find element with locator {0} and here is the exception ", locator));
                return null;
            }
        }

        public enum by
        {
            Id,
            Name,
            ClassName,
            TagName,
            Xpath,
            CssSelector,
            LinkText,
            none,
            XPath
        }

    }
}
