using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static WinAppDriver.Utilities.ElementFactory;

namespace WinAppDriver.Utilities
{
    public class WindowsUIElement
    {
        public string Name;
        public string Locator;
        public By By;

        public WindowsUIElement(By by,string Locator,string Name="default")
        {
            this.Name = Name;
            By = by;
            this.Locator = Locator;
        }
       

    }
}
