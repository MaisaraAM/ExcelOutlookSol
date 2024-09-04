using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinAppDriver.Utilities
{
    public class ElementFactory
    {
        private WindowsDriver<WindowsElement> _session;
        public ElementFactory(WindowsDriver<WindowsElement> session)
        {
            _session = session;
        }

        public WindowsElement findElement(By b, string locator, bool Optional = false, bool checkException = true)
        {

            try
            {
                switch (b)
                {
                    case By.AccessibilityId:
                        return _session.FindElementByAccessibilityId(locator);
                    case By.Name:
                        return _session.FindElementByName(locator);
                    case By.TagName:
                        return _session.FindElementByTagName(locator);
                    case By.Xpath:
                        return _session.FindElementByXPath(locator);
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
                    //if (checkException)
                    //    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot Find element with locator {0} and here is the exception ", locator));
                    throw e;
                }
            }
        }
        public WindowsElement findElement(By b, string locator, WindowsElement ele, bool Optional = false,bool checkException=true)
        {

            try
            {
                switch (b)
                {
                    case By.AccessibilityId:
                        return (WindowsElement)ele.FindElementByAccessibilityId(locator);
                    case By.Name:
                        return (WindowsElement)ele.FindElementByName(locator);
                    case By.TagName:
                        return (WindowsElement)ele.FindElementByTagName(locator);
                    case By.Xpath:
                        return _session.FindElementByXPath(locator);
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
                    //if (checkException)
                    //    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot Find element with locator {0} and here is the exception ", locator));
                    throw e;
                }
            }
        }

        public IList<WindowsElement> findElements(By b, string locator)
        {
            try
            {
                switch (b)
                {
                    case By.ClassName:
                        return _session.FindElementsByClassName(locator);
                    case By.Name:
                        return _session.FindElementsByName(locator);
                    case By.TagName:
                        return _session.FindElementsByTagName(locator);
                    case By.Xpath:
                        return _session.FindElementsByXPath(locator);
                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                //Please report the exception once the reporting module is ready 
                return null;
            }
        }
        public dynamic findElements(By b, string locator, WindowsElement ele)
        {
            try
            {
                switch (b)
                {
                    case By.ClassName:
                        return (IList<WindowsElement>)ele.FindElementsByClassName(locator);
                    case By.Name:
                        var bb= ele.FindElementsByName(locator);
                        return bb;
                    case By.TagName:
                        return (IList<WindowsElement>)ele.FindElementsByTagName(locator);
                    case By.Xpath:
                        return (IList<WindowsElement>)ele.FindElementsByXPath(locator);
                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                //Please report the exception once the reporting module is ready 
                return null;
            }
        }

        public enum By
        {
            AccessibilityId,
            Name,
            ClassName,
            TagName,
            Xpath,
            none
        }

    }
}
