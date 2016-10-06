using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace AutomatedTestingProject
{

    [TestClass]
    public class UnitTest1 : SeleniumWrapperClass
    {

        static bool OneTimeSetup = false;

        [TestInitialize]
        public void EnvironmentalSettings()
        {
            if (!OneTimeSetup)
            {
                OneTimeSetup = true;
                string path = @"C:\AutomatedTestingSolution";
                IE_DRIVER_PATH = path + @"\Selenium Setup\IEDriverServer_Win32_2.42.0";
                fileSavePath = path + @"\Screenshots";

                Initialize(fileSavePath);
            }
            
        }

        [TestMethod]
        public void OpenTheLinkTest()
        {
            driver.Navigate().GoToUrl(@"http://www.telerik.com/");
            RadTextBox("TextBoxID", "Text1");
            RadDatePicker("DatePickerID", "01/01/2016");
            RadComboBox("ComboBoxID", "Text1");
            RadSubmitButton("ButtonID");
            TakeScreenShot("Screenshot1");

        }
    }
}