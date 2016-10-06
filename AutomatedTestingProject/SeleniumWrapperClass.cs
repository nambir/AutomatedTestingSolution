using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Firefox;
using System.Threading;
using System.Drawing.Imaging;
using System.IO;
using OpenQA.Selenium.IE;
using System.Data;
using System.Reflection;
using System.Diagnostics;
using OpenQA.Selenium.Remote;

namespace AutomatedTestingProject
{
    public class SeleniumWrapperClass
    {
        public static IWebDriver driver;
        public static string IE_DRIVER_PATH;
        public static string baseURL;
        public static string fileSavePath;
        public static string excelFilePath;
        public static string TestClassName;
        public static string TestMethodName;
        public static string environment;
        public static int Timeout_ControlLevel = 5;
        public static int Timeout_Button = 5;
        public static int Timeout_Receipt_Screenshot = 15;
        public static string Authentication;

        public SeleniumWrapperClass()
        {

        }

        public void Initialize(string _filePath)
        {


            if (driver == null)
            {
                InitializeIE();
            }
        }

        private void InitializeIE()
        {
            var options = new InternetExplorerOptions();
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            //options.i

            var capabilities = DesiredCapabilities.InternetExplorer();
            capabilities.SetCapability("ie.ensureCleanSession", true);

            driver = new InternetExplorerDriver(IE_DRIVER_PATH, options, new TimeSpan(0, 0, 5, 0, 0));
        }


        public void GoToUrl(string URL)
        {
            driver.Navigate().GoToUrl(URL);
        }

        public void AspComboBox(string controlID, string textToSelect)
        {
            Thread.Sleep(1000);
            WaitControlToLoad(Timeout_ControlLevel);

            SelectElement select = new SelectElement(driver.FindElement(By.Id(controlID)));
            select.SelectByText(textToSelect);
        }

        public void RadComboBox(string controlID, string textToSelect)
        {
            try
            {
                FixedWait(Timeout_ControlLevel);

                DoWait(1);
                var cbo = driver.FindElement(By.Id(controlID + "_Arrow"));
                cbo.Click();
                DoWait(3);

                IWebElement List = driver.FindElement(By.Id(controlID + "_Input"));
                var item = driver.FindElements(By.TagName("li")).Where(elem => elem.Text.Trim().ToUpper() == "" + textToSelect.Trim().ToUpper() + "").FirstOrDefault();

                item.Click();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "textToSelect:" + textToSelect + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void AspTextBox(string controlID, string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                var radtb = driver.FindElement(By.Id(controlID));
                radtb.SendKeys(value);
                radtb.SendKeys(Keys.Tab);
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadTextBox(string controlID, string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);
                var radtb = driver.FindElement(By.Id(controlID));
                radtb.Clear();
                radtb.SendKeys(value);
                radtb.SendKeys(Keys.Tab);
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadTextBox_OLD(string controlID, string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                var radtb = driver.FindElement(By.Id(controlID + "_text"));
                //var radtb = driver.FindElement(By.Id(controlID));
                radtb.Clear();
                radtb.SendKeys(value);
                radtb.SendKeys(Keys.Tab);
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadSubmitButton(string controlID)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_Button);

                var btn = driver.FindElement(By.Id(controlID));
                btn.Click();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadCheckBoxButton(string controlID)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                var radChkbtn = driver.FindElement(By.Id(controlID));
                radChkbtn.Click();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadDatePicker(string controlID, string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                string IEVersion = (((OpenQA.Selenium.Remote.RemoteWebDriver)(driver)).Capabilities).Version;

                IWebElement raddp = null;

                if (IEVersion.Equals("10"))
                {
                    raddp = driver.FindElement(By.Id(controlID + "_dateInput"));
                }
                else
                {
                    //raddp = driver.FindElement(By.Id(controlID + "_dateInput_text"));
                    raddp = driver.FindElement(By.Id(controlID + "_dateInput"));
                }
                raddp.Clear();
                raddp.SendKeys(value);
                raddp.SendKeys(Keys.Tab);
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadDatePicker_OLD(string controlID, string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                var raddp = driver.FindElement(By.Id(controlID + "_dateInput_text"));
                raddp.Clear();
                raddp.SendKeys(value);
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void RadTabSelect(string controlID, string tabNo)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);
                var radtab = driver.FindElement(By.XPath("//div[@id='" + controlID + "']/div/ul/li[" + tabNo + "]/a/span/span/span"));
                radtab.Click();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "value:" + tabNo + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public void Hyperlink(string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                driver.FindElement(By.XPath("//a[contains(.,'" + value + "')]")).Click();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "Hyperlink" + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + value + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception("Hyperlink", ex);
            }
        }

        public void Hyperlink_OLD(string value)
        {
            FixedWait(Timeout_ControlLevel);
            driver.FindElement(By.CssSelector("div[@id=\"rtMid\"]/descendant::span[text()='" + value + "']")).Click();

            //IList<IWebElement> links = driver.FindElements(By.TagName("a"));
            //links.First(element => element.Text == value).Click();

            // your logic with traditional foreach loop
            //foreach (var link in links)
            //{
            //    if (link.Text == "YouTube")
            //    {
            //        link.Click();
            //        break;
            //    }
            //}
        }

        public void Label(string value)
        {
            try
            {
                Thread.Sleep(1000);
                FixedWait(Timeout_ControlLevel);

                driver.FindElement(By.XPath("//label[contains(.,'" + value + "')]")).Click();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "Label" + newLine + "value:" + value + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + value + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception("Label", ex);
            }
        }

        public void ConfirmAlertBox(bool accept)
        {
            try
            {
                Thread.Sleep(2000);

                if (accept)
                {
                    driver.SwitchTo().Alert().Accept();
                }
                else
                {
                    driver.SwitchTo().Alert().Dismiss();
                }
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "ConfirmAlertBox" + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + "ConfirmAlertBox" + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception("ConfirmAlertBox", ex);
            }
        }

        public void AlertOk()
        {
            try
            {
                Thread.Sleep(1000);

                var myAlert = driver.SwitchTo().Alert();
                myAlert.Accept();
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "AlertOk" + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + "AlertOk" + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception("AlertOk", ex);
            }
        }

        public string ReadRadTextBox(string controlID)
        {
            try
            {
                FixedWait(Timeout_Button);
                var radtb = driver.FindElement(By.Id(controlID));
                string value = radtb.GetAttribute("value");
                return value;
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public string ReadRadTextBox_OLD(string controlID)
        {
            try
            {
                FixedWait(Timeout_ControlLevel);
                var radtb = driver.FindElement(By.Id(controlID + "_text"));
                string value = radtb.GetAttribute("value");
                return value;
            }
            catch (Exception ex)
            {
                var callingMethod = new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod();
                TestClassName = callingMethod.ReflectedType.Name;
                TestMethodName = callingMethod.Name;
                string newLine = Environment.NewLine;
                Log("Error:" + newLine + "TestClassName:" + TestClassName + newLine + "TestMethodName:" + TestMethodName + newLine
                    + "controlID:" + controlID + newLine + "ex.Message:" + ex.Message);

                TakeScreenShot("Error-" + TestClassName + "-" + TestMethodName + "-" + controlID + "-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(controlID, ex);
            }
        }

        public string ReadCRMHeader(string Class)
        {
            try
            {
                WaitControlToLoad(5);
                var h1 = driver.FindElement(By.CssSelector("h1." + Class));
                string value = h1.Text;
                return value;
            }
            catch (Exception ex)
            {
                Log(Class + "-" + ex.Message);
                TakeScreenShot("Error-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(Class, ex);
            }
        }

        public string Readlabel(string ControlId)
        {
            try
            {
                WaitControlToLoad(5);
                var h1 = driver.FindElement(By.Id(ControlId));
                string value = h1.Text;
                return value;
            }
            catch (Exception ex)
            {
                Log(ControlId + "-" + ex.Message);
                TakeScreenShot("Error-" + DateTime.Now.ToString().Replace('/', '-').Replace(':', '-'));
                throw new Exception(ControlId, ex);
            }
        }

        public static void TakeScreenShot(string fileName)
        {
            try
            {
                FixedWait(Timeout_Receipt_Screenshot);

                ITakesScreenshot screenshotDriver = (ITakesScreenshot)driver;
                Screenshot screenshot = screenshotDriver.GetScreenshot();
                fileName = fileSavePath + @"\" + fileName;
                screenshot.SaveAsFile(fileName + ".jpeg", ImageFormat.Jpeg);
            }
            catch (Exception ex)
            {
                Log("TakeScreenShot" + "-" + ex.Message);
                throw new Exception("TakeScreenShot", ex);
            }
        }

        public static void FixedWait(int seconds)
        {
            DoWait(seconds);
            WaitControlToLoad(seconds);
        }

        public static void DoWait(int seconds)
        {

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds));

            var waitComplete = wait.Until<bool>(

                arg =>
                {

                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(seconds));

                    return true;
                });
        }

        public static void WaitControlToLoad(int timeOutInSeconds)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            String command = "return document.readyState";

            //Check the readyState before doing any waits
            //if (js.ExecuteScript(command).ToString().Equals("complete"))
            //{
            //    return;
            //}

            for (int i = 0; i < timeOutInSeconds; i++)
            {
                try
                {
                    Thread.Sleep(1000);

                    if (js.ExecuteScript(command).ToString().Equals("complete"))
                    {
                        break;
                    }
                }
                catch (Exception e)
                {
                    Log(e.Message);

                    if (js.ExecuteScript(command).ToString().Equals("complete"))
                    {
                        break;
                    }
                }
            }


            //bool pageLoaded = false;
            //object state = null;
            //int count = 0;
            //while (pageLoaded.Equals(false))
            //{
            //    if (((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState") != null)
            //    {
            //        state = ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState");
            //        if(state.Equals("complete"))
            //        pageLoaded = true;
            //    }
            //    else
            //    {
            //        count++;
            //    }
            //}
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds));

            //wait.Until<Boolean>(delegate(IWebDriver drv)
            //{
            //    return (Boolean)((IJavaScriptExecutor)drv).ExecuteScript("return document.readyState;").Equals("complete");

            //    //while (((IJavaScriptExecutor)drv).ExecuteScript("return document") != null)
            //    //{
            //    //    return (Boolean)((IJavaScriptExecutor)drv).ExecuteScript("return document.readyState;").Equals("complete");
            //    //}
            //    //return false;
            //});
        }

        public static void Log(string msg)
        {
            try
            {
                System.IO.FileStream wFile;
                DateTime dateTime = DateTime.Now;
                byte[] byteData = null;
                byteData = Encoding.ASCII.GetBytes(Environment.NewLine + dateTime + ":" + msg);
                wFile = new FileStream(fileSavePath + @"\log.txt", FileMode.Append);
                wFile.Write(byteData, 0, byteData.Length);
                wFile.Close();
            }
            finally
            {
            }
        }

 
    }
}
