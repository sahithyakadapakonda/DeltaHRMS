using System;
using System.Collections.Generic;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace DeltaHRMS
{
    class SetMethods
        
    {
        public static void EnterText(IWebDriver driver, String element, string value, string elementtype)
        {
            if (elementtype == "Id")
                driver.FindElement(By.Id(element)).SendKeys(value);
            if (elementtype == "Name")
                driver.FindElement(By.Name(element)).SendKeys(value);
        }
        public static void Click(IWebDriver driver, String element, string elementtype)
        {
            if (elementtype == "Id")
                driver.FindElement(By.Id(element)).Click();
            if (elementtype == "Name")
                driver.FindElement(By.Name(element)).Click();
            if (elementtype == "XPath")
                driver.FindElement(By.XPath(element)).Click();
            if (elementtype == "LinkText")
                driver.FindElement(By.LinkText(element)).Click();
            if (elementtype == "CssSelector")
                driver.FindElement(By.CssSelector(element)).Click();
        }

        public static void SelectDropdown(IWebDriver driver, String element, string value, string elementtype)
        {
            if (elementtype == "Id")
                new SelectElement(driver.FindElement(By.Id(element))).SelectByText(value);
            if (elementtype == "Name")
                new SelectElement(driver.FindElement(By.Name(element))).SelectByText(value);
            if (elementtype == "XPath")
                new SelectElement(driver.FindElement(By.XPath(element))).SelectByText(value);
        }
        
    }

    
    
}
 