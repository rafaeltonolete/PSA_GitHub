using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA002_BulkCreateTimeEntries
{


    [TestClass]
    public class BulkCreateTimeEntries
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmuri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());



        private XrmApp xrmApp;
        private WebClient client;


        /*Initialize XrmApp for UCI*/

        [TestInitialize]
        public void AppSetup()
        {
            this.client = new WebClient(TestSettings.Options);
            this.xrmApp = new XrmApp(client);
        }

        //***Given there are Logged Time Entries in a specified week
        [TestMethod]
        public void TC05_BulkCopyAndCreateTimeEntries()
        {


            using (xrmApp)
            {

                //Log In
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //(During this delay, enter the number code then proceed to welcome page of D365)
                xrmApp.ThinkTime(30000);


                //Navigate to Project Service
                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Navigate to Time Entries
                xrmApp.Navigation.OpenSubArea("Projects", "Time Entries");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click on Copy Week
                xrmApp.CommandBar.ClickCommand("Copy Week");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Select Date
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@id,'DatePicker')][1]")).Click();
                client.Browser.Driver.FindElement(By.XPath(".//button[contains(@aria-label,' 7,')]")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.FindElement(By.XPath(".//button[contains(@data-id,'Cancel')]")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);



                //reclick Copy Week
                xrmApp.CommandBar.ClickCommand("Copy Week");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //select Date
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@id,'DatePicker')][1]")).Click();
                client.Browser.Driver.FindElement(By.XPath(".//button[contains(@aria-label,' 7,')]")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Select Time Entries to copy
                client.Browser.Driver.FindElement(By.XPath(".//div[contains(@id,'dialogTabsContainer')]//button[@title='Select All']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Click Copy button
                client.Browser.Driver.FindElement(By.XPath(".//button[contains(@data-id,'Copy')][@aria-label='Copy']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Navigate to Time Entries
                xrmApp.Navigation.OpenSubArea("Projects", "Time Entries");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Switch view to My Time Entries For This week
                xrmApp.Grid.SwitchView("My Time Entries For This Week");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Search Time Entry , there should be a time entry record copied from a week or else this will cause an Error
                client.Browser.Driver.FindElement(By.XPath(".//div[@data-id='cell-0-9']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

            }


        }







        [TestCleanup]
        public void DisposeDriver()
        {
            client.Browser.Driver.Quit();
            xrmApp.Dispose();
        }



    }
}
