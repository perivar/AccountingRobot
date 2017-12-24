using System;
using OpenQA.Selenium;
using OpenQA.Selenium.PhantomJS;

namespace AccountingRobot
{
    public static class PhantomJSUtils
    {
        public static IWebDriver GetDriver()
        {
            var options = new PhantomJSOptions();
            options.AddAdditionalCapability("IsJavaScriptEnabled", true);
            options.AddAdditionalCapability("phantomjs.page.settings.userAgent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.94 Safari/537.36");
            var driver = new PhantomJSDriver(options);
            driver.Manage().Window.Size = new System.Drawing.Size(1440, 1000);

            return driver;
        }
    }
}
