using static WebDriver.Utilities.ElementFactory;

namespace WebDriver.Utilities
{
    public class WebUIElement
    {
        public string Name;
        public string Locator;
        public by By;
        public bool Displayed;

        public WebUIElement(by by, string Locator, string Name = "default")
        {
            this.Name = Name;
            By = by;
            this.Locator = Locator;
        }


    }
}
