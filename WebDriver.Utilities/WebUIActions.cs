using System;
using System.Collections.Generic;
using static WebDriver.Utilities.WebWaits;
using OpenQA.Selenium.Interactions;
using System.Linq;
using OpenQA.Selenium;
using By = WebDriver.Utilities.ElementFactory.by;
using System.Threading;

namespace WebDriver.Utilities
{
    public class WebUIActions
    {

        private IWebDriver _webDriver;
        ElementFactory elementFactory;
        WebWaits waits;

        public WebUIActions(IWebDriver webDriver)
        {
            this._webDriver = webDriver;
            elementFactory = new ElementFactory(webDriver);

        }

        #region elementsActions

        public void clickOnElement(WebUIElement ele, bool assert = false, WebUIElement expectedElement = null)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {

                try
                {
                    element.Click();
                    if (assert)
                    {
                        if (expectedElement == null)
                            throw new Exception(String.Format("Cannot find the expected element with locator {0} ", expectedElement.Locator));
                        IWebElement expectedElementObject = elementFactory.findElement(expectedElement.By, expectedElement.Locator);

                    }

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with Locator: {0} and Name:{1} and here is the exception ", ele.Locator,ele.Name));
                }
            }
        }
        public void clickOnElement(WebUIElement ele, WebUIElement eleParent, bool assert = false, WebUIElement expectedElement = null)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {

                try
                {
                    element.Click();
                    if (assert)
                    {
                        if (expectedElement == null)
                            throw new Exception(String.Format("Cannot find the expected element with locator {0} ", expectedElement.Locator));
                        IWebElement expectedElementObject = elementFactory.findElement(expectedElement.By, expectedElement.Locator);

                    }

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception ", ele.Locator));
                }
            }
        }
        public void setText(WebUIElement ele, string text, bool clear = true, bool assert = false, bool exactMatch = false)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    try
                    {
                        int trial = 3;
                        bool textAssertions = false;
                        if (clear) element.Clear();
                        if (assert)
                        {
                            do
                            {
                                trial--;
                                element.Click();
                                element.Clear();
                                element.SendKeys(text);

                                textAssertions = exactMatch ? String.Equals(element.Text, text) : String.Equals(element.Text.ToLower(), text.ToLower());

                            }
                            while (!textAssertions && trial > 0);

                        }
                        else
                        {
                            element.Click();
                            element.SendKeys(text);
                        }
                        if (textAssertions)
                        {
                            throw new Exception(String.Format("Couldn't match the text of element with locator {0} with current text {1} and the expected text {2}", ele.Locator, element.Text, text));
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot set text to element with locator {0} and here is the exception  ", ele.Locator));
                    }
                }
            }
        }
        public void setText(WebUIElement ele, WebUIElement eleParent, string text, bool clear = true, bool assert = false, bool exactMatch = false)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
                var element = elementFactory.findElement(ele.By, ele.Locator, parent);
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    try
                    {
                        int trial = 3;
                        bool textAssertions = false;
                        if (clear) element.Clear();
                        if (assert)
                        {
                            do
                            {
                                trial--;
                                element.Click();
                                element.Clear();
                                element.SendKeys(text);

                                textAssertions = exactMatch ? String.Equals(element.Text, text) : String.Equals(element.Text.ToLower(), text.ToLower());

                            }
                            while (!textAssertions && trial > 0);

                        }
                        else
                        {
                            element.Click();
                            element.SendKeys(text);
                        }
                        if (textAssertions)
                        {
                            throw new Exception(String.Format("Couldn't match the text of element with locator {0} with current text {1} and the expected text {2}", ele.Locator, element.Text, text));
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot set test to element with locator {0} and here is the exception  ", ele.Locator));
                    }
                }
            }
        }
        public void sendKeyboardKeys(WebUIElement ele, string text, bool withClick = true)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
                try
                {
                    if (element != null && WaitAndCheckTillControlToDisplay(element))
                    {
                        if (withClick)
                            element.Click();
                        element.Clear();
                        element.SendKeys(text);
                    }
                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot send controls to element with locator {0} and here is the exception  ", ele.Locator));
                }

            }
        }
        public void sendKeyboardKeys(WebUIElement ele, WebUIElement eleParent, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
                var element = elementFactory.findElement(ele.By, ele.Locator, parent);
                try
                {
                    if (element != null && WaitAndCheckTillControlToDisplay(element))
                    {
                        element.Click();
                        element.Clear();
                        element.SendKeys(text);
                    }
                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot send controls to element with locator {0} and here is the exception  ", ele.Locator));
                }

            }
        }
        public void sendControls(WebUIElement ele, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
                try
                {
                    if (element != null && WaitAndCheckTillControlToDisplay(element))
                    {
                        element.Clear();
                        element.SendKeys(text);
                    }
                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot send controls to element with locator {0} and here is the exception  ", ele.Locator));
                }

            }
        }
        public string getText(WebUIElement ele)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            string element_text = "";
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
                try
                {
                    element_text = element.Text;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get text from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_text;
        }
        public string getText(WebUIElement ele, WebUIElement eleParent)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            string element_text = "";
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
                try
                {
                    element_text = element.Text;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get text from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_text;
        }
        public string getTextByValueAttribute(WebUIElement ele)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            string element_text = "";
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
                try
                {
                    element_text = element.GetAttribute("Value.Value").ToString();

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get text from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_text;
        }
        public string getTextByValueAttribute(WebUIElement ele, WebUIElement eleParent)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            string element_text = "";
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
                try
                {
                    string sss = _webDriver.PageSource.ToString();
                    element_text = element.GetAttribute("Value.Value").ToString();

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get text from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_text;
        }
        public void selectDropDownListItem(WebUIElement ele, string itemText, bool selectFirstItem = false)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            if ((getText(ele) != itemText) || (selectFirstItem))
            {
                clickOnElement(ele);
                if (checkDisplayed(new WebUIElement(By.Name, itemText)))
                {
                    moveToElementAndClick(new WebUIElement(By.Name, itemText));
                }
                else
                {
                    bool checkItemDiplayed = checkDisplayed(new WebUIElement(By.Name, itemText));
                    bool foundWithUp = true;
                    string currentDropdowntext = getText(ele);
                    string currentDropdowntextBefore = "";
                    while (checkItemDiplayed == false)
                    {
                        currentDropdowntextBefore = getText(ele);
                        element.SendKeys(Keys.ArrowUp + Keys.ArrowUp + Keys.ArrowUp + Keys.ArrowUp + Keys.ArrowUp);
                        currentDropdowntext = getText(ele);
                        checkItemDiplayed = checkDisplayed(new WebUIElement(By.Name, itemText));
                        if (currentDropdowntext == currentDropdowntextBefore)
                        {
                            foundWithUp = false;
                            break;
                        }
                    }
                    if (foundWithUp == false)
                    {
                        checkItemDiplayed = checkDisplayed(new WebUIElement(By.Name, itemText));
                        while (checkItemDiplayed == false)
                        {

                            element.SendKeys(Keys.ArrowDown + Keys.ArrowDown + Keys.ArrowDown + Keys.ArrowDown + Keys.ArrowDown);
                            checkItemDiplayed = checkDisplayed(new WebUIElement(By.Name, itemText));

                        }
                    }
                    moveToElementAndClick(new WebUIElement(By.Name, itemText));

                }
            }
        }
        public void selectDropDownListItem(WebUIElement ele, WebUIElement eleParent, string itemText, bool selectFirstItem = false)
        {

            if ((getText(ele, eleParent) != itemText) || (selectFirstItem))
            {
                clickOnElement(ele, eleParent);
                moveToElementAndClick(new WebUIElement(By.Name, itemText));
            }
        }

        public void selectDropDownListItemFirst(WebUIElement ele)
        {
            string currItem = getText(ele);
            if ((currItem.Trim().Length == 0) || (currItem.Trim() == "غير معرف"))
            {
                sendKeyboardKeys(ele, Keys.ArrowDown + Keys.Enter);
                currItem = getText(ele);
                if (currItem.Trim() == "غير معرف")
                    sendKeyboardKeys(ele, Keys.ArrowDown + Keys.Enter);
            }
            else
                selectDropDownListItem(ele, currItem, true);
        }

        public void selectDropDownListItemFirst(WebUIElement ele, WebUIElement eleParent)
        {
            string currItem = getText(ele);
            selectDropDownListItem(ele, eleParent, currItem, true);
        }
        public void selectDropDownListItemSecond(WebUIElement ele)
        {
            string currItem = getText(ele);
            if (currItem.Trim().Length == 0)
            {
                sendKeyboardKeys(ele, Keys.ArrowDown + Keys.Enter);
            }
            else
                selectDropDownListItem(ele, currItem, true);
        }
        public void setCalenderDate(WebUIElement ele, string year, string month, string day, int year_X_Offset = 40, int month_X_Offset = 55, int day_X_Offset = 72)
        {
            moveToElementAndClickAndSendText(ele, year, true, year_X_Offset, 8);
            moveToElementAndClickAndSendText(ele, month, true, month_X_Offset, 8);
            moveToElementAndClickAndSendText(ele, day, true, day_X_Offset, 8);
        }
        public void setCalenderDate(WebUIElement ele, WebUIElement eleParent, string year, string month, string day)
        {
            moveToElementAndClickAndSendText(ele, eleParent, year, true, 40, 8);
            moveToElementAndClickAndSendText(ele, eleParent, month, true, 55, 8);
            moveToElementAndClickAndSendText(ele, eleParent, day, true, 72, 8);
        }

        #endregion

        #region moveToElementActions

        public void moveToElementAndDoubleClick(WebUIElement ele, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).DoubleClick().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot DoubleClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndDoubleClick(WebUIElement ele, WebUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).DoubleClick().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot DoubleClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndRightClick(WebUIElement ele, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).ContextClick().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).ContextClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot RightClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndRightClick(WebUIElement ele, WebUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).ContextClick().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).ContextClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot RightClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClick(WebUIElement ele, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).Click().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).Click().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClick(WebUIElement ele, WebUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).Click().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).Click().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElement(WebUIElement ele, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot find element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElement(WebUIElement ele, WebUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot find element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClickMenus(WebUIElement ele, bool heightDivededByY = false, int x = 0, int y = 0)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (heightDivededByY)
                        new Actions(_webDriver).MoveToElement(element, element.Size.Width - x, element.Size.Height / y).Click().Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element, element.Size.Width - x, element.Size.Height - y).Click().Perform();
                    Thread.Sleep(2000);
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveByOffsetAndClick(int x = 0, int y = 0)
        {
            try
            {
                new Actions(_webDriver).MoveByOffset(x, y).Click().Perform();
                Thread.Sleep(2000);
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot MoveByOffset click action perform "));
            }
        }
        public void moveByOffsetAndPerform(int x = 0, int y = 0)
        {
            try
            {
                new Actions(_webDriver).MoveByOffset(x, y).Perform();
                Thread.Sleep(2000);
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot MoveByOffset action perform "));
            }
        }
        public void moveToElementAndClickAndSendText(WebUIElement ele, string txt, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).Click().SendKeys(txt).Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).Click().SendKeys(txt).Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot write text on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClickAndSendText(WebUIElement ele, WebUIElement eleParent, string txt, bool withOffset = false, int x = 0, int y = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_webDriver).MoveToElement(element, x, y).Click().SendKeys(txt).Perform();
                    else
                        new Actions(_webDriver).MoveToElement(element).Click().SendKeys(txt).Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot write text on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }

        #endregion

        #region propertiesActions

        public bool checkEnabled(WebUIElement ele, bool withHandleException = true)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            bool element_enabled = false;
            if (element != null)
            {
                try
                {
                    element_enabled = element.Enabled;

                }
                catch (Exception e)
                {
                    if (withHandleException)
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_enabled;
        }
        public bool checkEnabled(WebUIElement ele, WebUIElement eleParent)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            bool element_enabled = false;
            if (element != null)
            {
                try
                {
                    element_enabled = element.Enabled;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_enabled;
        }
        public bool checkDisplayed(WebUIElement ele, bool withHandleException = true,bool checkException=true)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator,false,checkException);
            bool element_displyaed = false;
            if (element != null)
            {
                try
                {
                    element_displyaed = element.Displayed;

                }
                catch (Exception e)
                {
                    if (withHandleException)
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_displyaed;
        }
        public bool checkDisplayed(WebUIElement ele, WebUIElement eleParent)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            bool element_displyaed = false;
            if (element != null)
            {
                try
                {
                    element_displyaed = element.Displayed;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_displyaed;
        }
        public bool checkSelected(WebUIElement ele)
        {
            IWebElement element = elementFactory.findElement(ele.By, ele.Locator);
            bool element_selected = false;
            if (element != null)
            {
                try
                {
                    element_selected = element.Selected;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_selected;
        }
        public bool checkSelected(WebUIElement ele, WebUIElement eleParent)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            bool element_selected = false;
            if (element != null)
            {
                try
                {
                    element_selected = element.Selected;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_selected;
        }

        #endregion

        #region gridActions

        public List<String> getGridColumnValues(WebUIElement ele, int skipCount = 0)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            if (column != null)
            {
                List<String> values = new List<string>();
                foreach (IWebElement element in column)
                {
                    if (element != null)
                    {
                        values.Add(String.IsNullOrEmpty(element.Text) ? "" : element.Text);
                    }
                    else
                    {
                        values.Add("");

                    }

                }
                return values;
            }
            return null;

        }
        public List<String> getGridColumnValues(WebUIElement ele, WebUIElement eleParent, int skipCount = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator, parent);

            if (column != null)
            {
                List<String> values = new List<string>();
                foreach (IWebElement element in column)
                {
                    if (element != null)
                    {
                        values.Add(String.IsNullOrEmpty(element.Text) ? "" : element.Text);
                    }
                    else
                    {
                        values.Add("");

                    }

                }
                return values;
            }
            return null;

        }
        public void setGridItemText(WebUIElement ele, int index, string text)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator);
            try
            {
                if (column != null)
                {
                    column[index].Clear();
                    new Actions(_webDriver).MoveToElement(column[index]).Click().SendKeys(text).Perform();

                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot set text on Grid element with locator {0} and here is the exception ", ele.Locator));
            }

        }
        public void setGridItemText(WebUIElement ele, WebUIElement eleParent, int index, string text)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator, parent);
            try
            {
                if (column != null)
                {
                    column[index].Clear();
                    new Actions(_webDriver).MoveToElement(column[index]).Click().SendKeys(text).Perform();

                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot set text on Grid element with locator {0} and here is the exception ", ele.Locator));
            }

        }
        public void clickOnGridItemByItsIndex(WebUIElement ele, int index, int skipCount = 0)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            try
            {
                if (column != null)
                {
                    column[index].Click();
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on Grid element with locator {0} and here is the exception  ", ele.Locator));
            }

        }
        public void clickOnGridItemByItsIndex(WebUIElement ele, WebUIElement eleParent, int index, int skipCount = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator, parent).Skip(skipCount).ToList();
            try
            {
                if (column != null)
                {
                    column[index].Click();
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on Grid element with locator {0} and here is the exception  ", ele.Locator));
            }

        }
        public void clickOnGridItemByItsText(WebUIElement ele, string text, int skipCount = 0)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (IWebElement element in column)
                {
                    try
                    {
                        if (element != null)
                        {
                            if (element.Text == text)
                                element.Click();
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on Grid element with locator {0} and here is the exception  ", ele.Locator));
                    }

                }
            }
        }
        public void clickOnGridItemByItsText(WebUIElement ele, WebUIElement eleParent, string text, int skipCount = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator, parent).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (IWebElement element in column)
                {
                    try
                    {
                        if (element != null)
                        {
                            if (element.Text == text)
                                element.Click();
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on Grid element with locator {0} and here is the exception  ", ele.Locator));
                    }

                }
            }
        }
        public void clickAndKeyUpOnGridItemByText(WebUIElement ele, string text, int skipCount = 0)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (IWebElement element in column)
                {
                    try
                    {
                        if (element != null)
                        {
                            if (element.Text == text)
                                new Actions(_webDriver).KeyDown(Keys.Control).Click(element).KeyUp(Keys.Control).Build().Perform();
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
                    }

                }
            }

        }
        public void clickAndKeyUpOnGridItemByText(WebUIElement ele, WebUIElement eleParent, string text, int skipCount = 0)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator, parent).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (IWebElement element in column)
                {
                    try
                    {
                        if (element != null)
                        {
                            if (element.Text == text)
                                new Actions(_webDriver).KeyDown(Keys.Control).Click(element).KeyUp(Keys.Control).Build().Perform();
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
                    }

                }
            }

        }
        public void doubleClickGridItem(WebUIElement ele, int index)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator);
            try
            {
                if (column != null)
                {
                    new Actions(_webDriver).MoveToElement(column[index]).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void doubleClickGridItem(WebUIElement ele, WebUIElement eleParent, int index)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator, parent);
            try
            {
                if (column != null)
                {
                    new Actions(_webDriver).MoveToElement(column[index]).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void rightClickGridItem(WebUIElement ele, int index)
        {
            IList<IWebElement> column = elementFactory.findElements(ele.By, ele.Locator);
            try
            {
                if (column != null)
                {
                    new Actions(_webDriver).MoveToElement(column[index]).ContextClick().Perform();
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void rightClickGridItem(WebUIElement ele, WebUIElement eleParent, int index)
        {
            IWebElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator, parent);
            try
            {
                if (column != null)
                {
                    new Actions(_webDriver).MoveToElement(column[index]).ContextClick().Perform();
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        #endregion

        public void scrollIntoViewElement(WebUIElement element)
        {

            Implicit_Wait(_webDriver, 60);
            IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", elementFactory.findElement(element.By, element.Locator));
        }





    }
}
