using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Xml;
using System.Linq;
using Aspose.Cells;


namespace DeltaHRMS
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
           

                Workbook wb = new Workbook("DeltaHRMS_StatusReport.xlsx");
                Worksheet sheet = wb.Worksheets[0];
                IWebDriver driver = new ChromeDriver();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            Cell cell;
                string d;
                
                String[] B = { "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13" ,"B14"};
                string[] C = { "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13","C14" };
                String[] D = { "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13","D14" };

                for (int i = 0; i <= 12; i++)
                {

                    string ValueB = B[i];
                    string ValueC = C[i];
                    string ValueD = D[i];
                    cell = sheet.Cells[ValueB];
                    string TestStep = cell.Value.ToString();
                    try
                    {
                       switch (ValueB)
                       {
                            case "B2":
                                if (TestStep.Contains("Browser"))
                                {
                                    driver.Navigate().GoToUrl("https://www.google.com/");
                                    driver.Manage().Window.Maximize();
                                    if (driver.Url.Contains("google"))
                                    {
                                        cell = sheet.Cells[ValueC];
                                        cell.PutValue("Pass");
                                    }
                                }
                                break;

                            case "B3":
                                if (TestStep.Contains("HRMS"))
                                {
                                    driver.Navigate().GoToUrl("http://deltahrmsqa.deltaintech.com/");
                                    if (driver.Url.Contains("deltaintech.com"))
                                    {

                                        cell = sheet.Cells[ValueC];
                                        cell.PutValue("Pass");
                                    }
                                }
                                break;

                            case "B4":
                                if (TestStep.Contains("Login"))
                                {
                                    XmlDocument xDoc = new XmlDocument();
                                    xDoc.Load("LoginDetails.Xml");
                                    string UserName = xDoc.DocumentElement.SelectSingleNode("UserName").InnerText;
                                    string Password = xDoc.DocumentElement.SelectSingleNode("Password").InnerText;
                                    SetMethods.EnterText(driver, "username", UserName, "Id");
                                    SetMethods.EnterText(driver, "password", Password, "Id");
                                    SetMethods.Click(driver, "loginsubmit", "Id");
                                    if (driver.Url.Contains("welcome"))
                                    {
                                        cell = sheet.Cells[ValueC];
                                        cell.PutValue("Pass");
                                    }

                                }
                                break;

                            case "B5":
                                if (TestStep.Contains("Services"))
                                {
                                    SetMethods.Click(driver, "thumbnail_4", "Id");
                                    if (driver.Url.Contains("welcome"))
                                    {
                                        cell = sheet.Cells[ValueC];
                                        cell.PutValue("Pass");
                                    }

                                }
                                break;

                        case "B6":
                            if (TestStep.Contains("Leaves"))
                            {
                                SetMethods.Click(driver, "acc_li_toggle_31", "Id");
                                if (driver.Url.Contains("welcome"))
                                {
                                    cell = sheet.Cells[ValueC];
                                    cell.PutValue("Pass");
                                }
                            }
                            break;

                        //case "B7":
                        //    if (TestStep.Contains("Request"))
                        //    {
                        //        SetMethods.Click(driver, "Leave Request", "LinkText");
                        //        if (driver.Url.Contains("leaverequest"))
                        //        {
                        //            cell = sheet.Cells[ValueC];
                        //            cell.PutValue("Pass");
                        //        }

                        //    }
                        //    break;

                        //case "B8":
                        //    if (TestStep.Contains("Apply"))
                        //    {
                        //        SetMethods.Click(driver, "//input[@value='Apply Leave']", "XPath");
                        //        if (driver.Url.Contains("leaverequest"))
                        //        {
                        //            cell = sheet.Cells[ValueC];
                        //            cell.PutValue("Pass");
                        //        }

                        //    }
                        //    break;

                        //case "B9":
                        //    if (TestStep.Contains("type"))
                        //    {
                        //        SetMethods.Click(driver, "//span[text()='Select Leave Type']", "XPath");
                        //        SetMethods.Click(driver, "//span[contains(text(),'Earned Leave ')]", "XPath");
                        //        if (driver.Url.Contains("leaverequest"))
                        //        {
                        //            cell = sheet.Cells[ValueC];
                        //            cell.PutValue("Pass");


                        //        }
                        //    }
                        //    break;

                        //case "B10":
                        //    if (TestStep.Contains("From"))
                        //    {
                        //        SetMethods.Click(driver, "//input[@id='from_date']", "XPath");
                        //        var days = driver.FindElements(By.CssSelector("a[class='ui-state-default']"));
                        //        var builder = new Actions(driver);
                        //        builder.Click(days[13]).Build().Perform();
                        //        if (driver.Url.Contains("leaverequest"))
                        //        {
                        //            cell = sheet.Cells[ValueC];
                        //            cell.PutValue("Pass");
                        //        }

                        //    }
                        //    break;

                        //case "B11":
                        //    if (TestStep.Contains("To"))
                        //    {
                        //        SetMethods.Click(driver, "//input[@id='to_date']", "XPath");
                        //        SetMethods.Click(driver, "a[class='ui-state-default']", "CssSelector");
                        //        SetMethods.Click(driver, "//input[@id='to_date']", "XPath");
                        //        var days = driver.FindElements(By.CssSelector("a[class='ui-state-default']"));
                        //        var builder = new Actions(driver);
                        //        builder.Click(days[2]).Build().Perform();
                        //        if (driver.Url.Contains("leaverequest"))
                        //        {
                        //            cell = sheet.Cells[ValueC];
                        //            cell.PutValue("Pass");

                        //        }

                        //    }

                        //    break;

                        //case "B12":
                        //    if (TestStep.Contains("Reason"))
                        //    {
                        //        SetMethods.EnterText(driver, "reason", "Personal Work", "Id");
                        //        if (driver.Url.Contains("leaverequest"))
                        //        {

                        //            cell = sheet.Cells[ValueC];
                        //            cell.PutValue("Pass");

                        //        }

                        //    }
                        //    break;

                        //case "B13":
                        //    if (TestStep.Contains("Apply"))
                        //    {

                        //        SetMethods.Click(driver, "submit", "Name");
                        //        cell = sheet.Cells[ValueC];
                        //        cell.PutValue("Pass");
                        //        Thread.Sleep(10000);

                        //    }
                        //    break;


                        case "B14":
                              if (TestStep.Contains("Allocated"))
                              {
                                SetMethods.Click(driver, "//div[contains(@class,'side-menu')]//li//a[@id='62']", "XPath");
                                SetMethods.Click(driver, "filter_all", "Id");
                                SetMethods.SelectDropdown(driver, "perpage_pendingleaves", "100","Id");
                                Thread.Sleep(3000);


                                IList<IWebElement> rows = driver.FindElements(By.XPath("//table[@class='grid']/tbody/tr"));
                                int NumberOfRows = rows.Count();
                                String FirstPath = "//*[@id='pendingleaves']/table/tbody/tr[";
                                String SecondPath = "]/td[6]/span";
                                double Sum = 0;

                                for (int k = 2; k < NumberOfRows; i++)
                                {
                                   string innerText = driver.FindElement(By.XPath(FirstPath + k + SecondPath)).Text;
                                    Double Days = Convert.ToDouble(innerText);
                                    Sum += Days;

                                }
                                cell = sheet.Cells["C14"];
                                cell.PutValue(Sum);

                              }
                            
                            break;


                            
                       }

                   


                    }
                    catch (Exception e)
                    {
                        cell = sheet.Cells["E2"];
                        cell.PutValue("Fail");
                        cell = sheet.Cells[ValueC];
                        cell.PutValue("Fail");
                        driver.Quit();
                    throw e;
           

                    }
                    finally
                    {
                        d = DateTime.Now.ToString();
                        cell = sheet.Cells[ValueD];
                        cell.PutValue(d);
                        wb.Save("DeltaHRMS_StatusReport_Updated.Xlsx", SaveFormat.Xlsx);
                        
                    }

                }
            driver.Quit();


        }

    }
}