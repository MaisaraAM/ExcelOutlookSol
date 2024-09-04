using System;
using System.Collections.Generic;
using OpenQA.Selenium.Appium.Windows;
using static WinAppDriver.Utilities.Waits;
using static WinAppDriver.Utilities.ElementFactory;
using OpenQA.Selenium.Interactions;
using System.Linq;
using OpenQA.Selenium;
using By = WinAppDriver.Utilities.ElementFactory.By;
using System.Threading;

namespace WinAppDriver.Utilities
{
    public class WindowsUIActions
    {

        private WindowsDriver<WindowsElement> _session;
        ElementFactory elementFactory;

        public WindowsUIActions(WindowsDriver<WindowsElement> session)
        {
            this._session = session;
            elementFactory = new ElementFactory(session);
        }

        #region elementsActions

        public void clickOnElement(WindowsUIElement ele, bool assert = false, WindowsUIElement expectedElement = null)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
                
                try
                {
                    element.Click();
                    if (assert)
                    {
                        if (expectedElement == null)
                            throw new Exception(String.Format("Cannot find the expected element with locator {0} ", expectedElement.Locator));
                        WindowsElement expectedElementObject = elementFactory.findElement(expectedElement.By, expectedElement.Locator);

                    }

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception ", ele.Locator));
                }
            }
        }
        public void clickOnElement(WindowsUIElement ele, WindowsUIElement eleParent, bool assert = false, WindowsUIElement expectedElement = null)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
               
                try
                {
                    element.Click();
                    if (assert)
                    {
                        if (expectedElement == null)
                            throw new Exception(String.Format("Cannot find the expected element with locator {0} ", expectedElement.Locator));
                        WindowsElement expectedElementObject = elementFactory.findElement(expectedElement.By, expectedElement.Locator);

                    }

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception ", ele.Locator));
                }
            }
        }
        public void setText(WindowsUIElement ele, string text, bool clear = true, bool assert = false, bool exactMatch = false)
        {
            if (!String.IsNullOrWhiteSpace(text)) {
                WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
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
        public void setText(WindowsUIElement ele, WindowsUIElement eleParent, string text, bool clear = true, bool assert = false, bool exactMatch = false)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
                var element = elementFactory.findElement(ele.By, ele.Locator,parent);
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
        public void sendKeyboardKeys(WindowsUIElement ele, string text,bool withClick=true)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
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
        public void sendKeyboardKeys(WindowsUIElement ele, WindowsUIElement eleParent, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
                var element = elementFactory.findElement(ele.By, ele.Locator,parent);
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
        public void sendControls(WindowsUIElement ele, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
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
        public string getText(WindowsUIElement ele)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            string element_text = "";
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {              
                try
                {
                    element_text= element.Text;

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get text from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_text;
        }
        public string getText(WindowsUIElement ele, WindowsUIElement eleParent)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
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
        public string getTextByValueAttribute(WindowsUIElement ele)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
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
        public string getTextByValueAttribute(WindowsUIElement ele, WindowsUIElement eleParent)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator, parent);
            string element_text = "";
            if (element != null && WaitAndCheckTillControlToDisplay(element))
            {
                try
                {
                    string sss = _session.PageSource.ToString();
                    element_text = element.GetAttribute("Value.Value").ToString();

                }
                catch (Exception e)
                {
                    HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get text from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_text;
        }
        public void selectDropDownListItem(WindowsUIElement ele, string itemText,bool selectFirstItem=false)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            if ((getText(ele) != itemText)||(selectFirstItem))
            {
                clickOnElement(ele);
                if (checkDisplayed(new WindowsUIElement(By.Name, itemText)))
                {
                    moveToElementAndClick(new WindowsUIElement(By.Name, itemText));
                }
                else
                {
                    bool checkItemDiplayed = checkDisplayed(new WindowsUIElement(By.Name, itemText));
                    bool foundWithUp = true;
                    string currentDropdowntext = getText(ele);
                    string currentDropdowntextBefore = "";
                    while (checkItemDiplayed == false)
                    {
                        currentDropdowntextBefore = getText(ele);
                        element.SendKeys(Keys.ArrowUp+ Keys.ArrowUp + Keys.ArrowUp + Keys.ArrowUp + Keys.ArrowUp );
                        currentDropdowntext = getText(ele);
                        checkItemDiplayed = checkDisplayed(new WindowsUIElement(By.Name, itemText));
                        if (currentDropdowntext == currentDropdowntextBefore)
                        {
                            foundWithUp = false;
                            break;
                        }
                    }
                    if (foundWithUp == false)
                    {
                        checkItemDiplayed = checkDisplayed(new WindowsUIElement(By.Name, itemText));
                        while (checkItemDiplayed == false)
                        {

                            element.SendKeys(Keys.ArrowDown+ Keys.ArrowDown+ Keys.ArrowDown+ Keys.ArrowDown+ Keys.ArrowDown);
                            checkItemDiplayed = checkDisplayed(new WindowsUIElement(By.Name, itemText));

                        }
                    }
                    moveToElementAndClick(new WindowsUIElement(By.Name, itemText));

                }
            }
        }
        public void selectDropDownListItem(WindowsUIElement ele, WindowsUIElement eleParent, string itemText, bool selectFirstItem = false)
        {

            if ((getText(ele, eleParent) != itemText) || (selectFirstItem))
            {
                clickOnElement(ele, eleParent);
                moveToElementAndClick(new WindowsUIElement(By.Name, itemText));
            }
        }

        public void selectDropDownListItemFirst(WindowsUIElement ele)
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

        public void selectDropDownListItemFirst(WindowsUIElement ele, WindowsUIElement eleParent)
        {
            string currItem = getText(ele);
            selectDropDownListItem(ele, eleParent, currItem, true);
        }
        public void selectDropDownListItemSecond(WindowsUIElement ele)
        {
            string currItem = getText(ele);
            if (currItem.Trim().Length == 0)
            {
                sendKeyboardKeys(ele, Keys.ArrowDown + Keys.Enter);
            }
            else
                selectDropDownListItem(ele, currItem, true);
        }
        public void setCalenderDate(WindowsUIElement ele,string year, string month, string day, int year_X_Offset= 40, int month_X_Offset = 55, int day_X_Offset = 72)
        {
            moveToElementAndClickAndSendText(ele,year,true, year_X_Offset, 8);
            moveToElementAndClickAndSendText(ele,month,true,month_X_Offset,8);
            moveToElementAndClickAndSendText(ele,day,true,day_X_Offset,8);
        }
        public void setCalenderDate(WindowsUIElement ele, WindowsUIElement eleParent, string year, string month, string day)
        {
            moveToElementAndClickAndSendText(ele,eleParent, year, true, 40, 8);
            moveToElementAndClickAndSendText(ele,eleParent, month,true, 55, 8);
            moveToElementAndClickAndSendText(ele,eleParent, day,  true, 72, 8);
        }

        #endregion

        #region moveToElementActions

        public void moveToElementAndDoubleClick(WindowsUIElement ele, bool withOffset = false, int x=0,int y=0)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).DoubleClick().Perform();
                    else
                        new Actions(_session).MoveToElement(element).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot DoubleClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndDoubleClick(WindowsUIElement ele, WindowsUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).DoubleClick().Perform();
                    else
                        new Actions(_session).MoveToElement(element).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot DoubleClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndRightClick(WindowsUIElement ele, bool withOffset = false, int x = 0, int y = 0)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).ContextClick().Perform();
                    else
                        new Actions(_session).MoveToElement(element).ContextClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot RightClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndRightClick(WindowsUIElement ele, WindowsUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).ContextClick().Perform();
                    else
                        new Actions(_session).MoveToElement(element).ContextClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot RightClick on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClick(WindowsUIElement ele, bool withOffset=false,int x = 0, int y = 0)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if(withOffset)
                        new Actions(_session).MoveToElement(element,x,y).Click().Perform();
                    else
                    new Actions(_session).MoveToElement(element).Click().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClick(WindowsUIElement ele, WindowsUIElement eleParent, bool withOffset = false, int x = 0, int y = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).Click().Perform();
                    else
                        new Actions(_session).MoveToElement(element).Click().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClickMenus(WindowsUIElement ele,bool heightDivededByY=false, int x = 0, int y = 0)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                   if(heightDivededByY)
                        new Actions(_session).MoveToElement(element, element.Size.Width - x, element.Size.Height / y).Click().Perform();
                   else
                        new Actions(_session).MoveToElement(element, element.Size.Width - x, element.Size.Height - y).Click().Perform();
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
                new Actions(_session).MoveByOffset(x, y).Click().Perform();
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
                new Actions(_session).MoveByOffset(x, y).Perform();
                Thread.Sleep(2000);
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot MoveByOffset action perform "));
            }
        }
        public void moveToElementAndClickAndSendText(WindowsUIElement ele,string txt, bool withOffset = false, int x = 0, int y = 0)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {     
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).Click().SendKeys(txt).Perform();
                    else
                        new Actions(_session).MoveToElement(element).Click().SendKeys(txt).Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot write text on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void moveToElementAndClickAndSendText(WindowsUIElement ele, WindowsUIElement eleParent, string txt, bool withOffset = false, int x = 0, int y = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
            try
            {
                if (element != null && WaitAndCheckTillControlToDisplay(element))
                {
                    if (withOffset)
                        new Actions(_session).MoveToElement(element, x, y).Click().SendKeys(txt).Perform();
                    else
                        new Actions(_session).MoveToElement(element).Click().SendKeys(txt).Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot write text on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }

        #endregion

        #region propertiesActions

        public bool checkEnabled(WindowsUIElement ele,bool withHandleException=true, bool checkException = true)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator,checkException);
            bool element_enabled = false;
            if (element != null)
            {
                try
                {
                    element_enabled = element.Enabled;

                }
                catch (Exception e)
                {
                    if(withHandleException)
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot get proparty from element with locator {0} and here is the exception ", ele.Locator));
                }
            }
            return element_enabled;
        }
        public bool checkEnabled(WindowsUIElement ele, WindowsUIElement eleParent)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
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
        public bool checkDisplayed(WindowsUIElement ele)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
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
        public bool checkDisplayed(WindowsUIElement ele, WindowsUIElement eleParent)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
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
        public bool checkSelected(WindowsUIElement ele)
        {
            WindowsElement element = elementFactory.findElement(ele.By, ele.Locator);
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
        public bool checkSelected(WindowsUIElement ele, WindowsUIElement eleParent)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var element = elementFactory.findElement(ele.By, ele.Locator,parent);
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

        public List<String> getGridColumnValues(WindowsUIElement ele, int skipCount = 0)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            if (column != null)
            {
                List<String> values = new List<string>();
                foreach (WindowsElement element in column)
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
        public List<String> getGridColumnValues(WindowsUIElement ele, WindowsUIElement eleParent, int skipCount = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator,parent);

            if (column != null)
            {
                List<String> values = new List<string>();
                foreach (WindowsElement element in column)
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
        public void setGridItemText(WindowsUIElement ele, int index, string text)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator);
            try
            {
                if (column != null)
                {
                    column[index].Clear();
                    new Actions(_session).MoveToElement(column[index]).Click().SendKeys(text).Perform();

                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot set text on Grid element with locator {0} and here is the exception ", ele.Locator));
            }

        }
        public void setGridItemText(WindowsUIElement ele, WindowsUIElement eleParent, int index, string text)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator, parent);
            try
            {
                if (column != null)
                {
                    column[index].Clear();
                    new Actions(_session).MoveToElement(column[index]).Click().SendKeys(text).Perform();

                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot set text on Grid element with locator {0} and here is the exception ", ele.Locator));
            }

        }
        public void clickOnGridItemByItsIndex(WindowsUIElement ele, int index, int skipCount = 0)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
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
        public void clickOnGridItemByItsIndex(WindowsUIElement ele, WindowsUIElement eleParent, int index, int skipCount = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator,parent).Skip(skipCount).ToList();
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
        public void clickOnGridItemByItsText(WindowsUIElement ele, string text, int skipCount = 0)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (WindowsElement element in column)
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
        public void clickOnGridItemByItsText(WindowsUIElement ele, WindowsUIElement eleParent, string text, int skipCount = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator,parent).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (WindowsElement element in column)
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
        public void clickAndKeyUpOnGridItemByText(WindowsUIElement ele, string text, int skipCount = 0)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (WindowsElement element in column)
                {
                    try
                    {
                        if (element != null)
                        {
                            if (element.Text == text)
                            new Actions(_session).KeyDown(Keys.Control).Click(element).KeyUp(Keys.Control).Build().Perform();
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
                    }

                }
            }

        }
        public void clickAndKeyUpOnGridItemByText(WindowsUIElement ele, WindowsUIElement eleParent, string text, int skipCount = 0)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator,parent).Skip(skipCount).ToList();
            if (column != null)
            {
                foreach (WindowsElement element in column)
                {
                    try
                    {
                        if (element != null)
                        {
                            if (element.Text == text)
                                new Actions(_session).KeyDown(Keys.Control).Click(element).KeyUp(Keys.Control).Build().Perform();
                        }
                    }
                    catch (Exception e)
                    {
                        HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
                    }

                }
            }

        }
        public void doubleClickGridItem(WindowsUIElement ele, int index)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator);
            try
            {
                if (column != null)
                {
                    new Actions(_session).MoveToElement(column[index]).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void doubleClickGridItem(WindowsUIElement ele, WindowsUIElement eleParent, int index)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator,parent);
            try
            {
                if (column != null)
                {
                    new Actions(_session).MoveToElement(column[index]).DoubleClick().Perform();
                }

            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void rightClickGridItem(WindowsUIElement ele, int index)
        {
            IList<WindowsElement> column = elementFactory.findElements(ele.By, ele.Locator);
            try
            {
                if (column != null)
                {
                    new Actions(_session).MoveToElement(column[index]).ContextClick().Perform();
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        public void rightClickGridItem(WindowsUIElement ele, WindowsUIElement eleParent, int index)
        {
            WindowsElement parent = elementFactory.findElement(eleParent.By, eleParent.Locator);
            var column = elementFactory.findElements(ele.By, ele.Locator,parent);
            try
            {
                if (column != null)
                {
                    new Actions(_session).MoveToElement(column[index]).ContextClick().Perform();
                }
            }
            catch (Exception e)
            {
                HandleExceptions.LogAnyExceptionAndFailTestCase(e, String.Format("Cannot click on element with locator {0} and here is the exception  ", ele.Locator));
            }
        }
        #endregion







    }
}
