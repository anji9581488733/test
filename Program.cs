using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SLV_for_dooku_2_devices
{
    internal class Program
    {


        public static void Anji(WindowsDriver<WindowsElement> ft, WindowsElement UU)
        {
            Actions RT = new Actions(ft);
            RT.MoveToElement(UU).Click().Build().Perform();
        }

        static void Main(string[] args)
        {
            WindowsDriver<WindowsElement> Storagel;
            AppiumOptions desiredCapabilities1 = new AppiumOptions();
            desiredCapabilities1.AddAdditionalCapability("app", @"C:\Program Files (x86)\ReSound\Dooku2.9.78.1\StorageLayoutViewer.exe");
            desiredCapabilities1.AddAdditionalCapability("appWorkingDir", @"C:\Program Files (x86)\ReSound\Dooku2.9.78.1");
            Storagel = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723/"), desiredCapabilities1);
            Storagel.Manage().Window.Maximize();
            WebDriverWait wait = new WebDriverWait(Storagel, TimeSpan.FromSeconds(30));
            Assert.IsNotNull(Storagel);
            Thread.Sleep(4000);



            Storagel.SwitchTo().Window(Storagel.WindowHandles.First());
            if (Storagel.FindElementByClassName("ConnectLeftRightUserControl").Displayed)
            {

                Storagel.FindElementByAccessibilityId("ToggleButton").Click();
                Storagel.FindElementByAccessibilityId("FittingDongle:0").Click();
            }


            

            Storagel.SwitchTo().Window(Storagel.WindowHandles.First());

            if (!Storagel.FindElementByClassName("DataGrid").Displayed)
            {

                var bu = Storagel.FindElementByClassName("Button");
                Anji(Storagel, (WindowsElement)bu);
            }


            String Serialnumber = "2000800246";
            

            Storagel.SwitchTo().Window(Storagel.WindowHandles.First());
            WebDriverWait Continue = new WebDriverWait(Storagel, TimeSpan.FromSeconds(500));
            

            do
            {
                var non = Storagel.FindElementByClassName("DataGrid");
                //var pan = non.FindElementsByClassName("DataGridCell");

                ReadOnlyCollection<AppiumWebElement> boxs = (ReadOnlyCollection<AppiumWebElement>)non.FindElementsByClassName("DataGridCell");
                Storagel.SwitchTo().Window(Storagel.WindowHandles.First());

                foreach (var element in boxs)
                {

                    if (element.Text == Serialnumber)
                    {
                        element.Click();
                        Storagel.FindElementByName("Connect").Click();

                        Thread.Sleep(5000);
                        Storagel.SwitchTo().Window(Storagel.WindowHandles.First());
                        Storagel.FindElementByName("_Read from").Click();

                        break;
                    }


                }


                
   
            }while(!Storagel.FindElementByName("_Read from").Displayed);

            Storagel.SwitchTo().Window(Storagel.WindowHandles.First());

            
            Storagel.SwitchTo().Window(Storagel.WindowHandles[0]);


            Continue.Until(ExpectedConditions.ElementToBeClickable(By.Name("Open")));
            Storagel.FindElementByName("Open").Click();


            var allCheckbox = Continue.Until(ExpectedConditions.ElementToBeClickable(By.Name("All")));
            Anji(Storagel, (WindowsElement)allCheckbox);

            Thread.Sleep(1000);
            Storagel.FindElementByName("Apply selection").Click();
            Storagel.FindElementByName("File").Click();

            Thread.Sleep(3000);



            var pp = Storagel.FindElementByName("Save as CDI file");
            Anji(Storagel, (WindowsElement)pp);


            Thread.Sleep(2000);
            Storagel.FindElementByAccessibilityId("1").Click();

            do
            {
                Storagel.SwitchTo().Window(Storagel.WindowHandles.First());

                if ((Storagel.FindElementByAccessibilityId("TitleBar")).Text == "Read Error")
                {
                    Storagel.FindElementByAccessibilityId("checkBoxIgnoreAll").Click();
                    Storagel.FindElementByAccessibilityId("buttonOk").Click();
                }

                else if (Storagel.FindElementByAccessibilityId("TitleBar").Text == "Saved CDI/XML")
                {
                    Thread.Sleep(2000);
                    Storagel.SwitchTo().Window(Storagel.WindowHandles.First());
                    Storagel.FindElementByName("Ok").Click();

                    break;
                }

                Storagel.SwitchTo().Window(Storagel.WindowHandles.First());
            } while ((Storagel.FindElementByAccessibilityId("TitleBar").Text) != "Create CDI file from layout");

            Storagel.SwitchTo().Window(Storagel.WindowHandles.First());
            Storagel.FindElementByAccessibilityId("Close").Click();


            







        }
    }
}
