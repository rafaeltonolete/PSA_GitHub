using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA001_CreateAndSubmitTimeEntry
{
    [TestClass]
    public class CreateAndSubmitTimeEntry
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

        [TestMethod]
        public void TC01_01_CreateAndSubmitTimeEntry()
        {

            using (xrmApp)
            {
                //Launch and login to D365 as User
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Enter code during this delay then proceed to CRM welcome page
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                //Click Project Service
                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                xrmApp.ThinkTime(5000);

                //Navigate To Time Entries
                xrmApp.Navigation.OpenSubArea("Projects", "Time Entries");
                xrmApp.ThinkTime(5000);

                //Click "+New"
                xrmApp.CommandBar.ClickCommand("New");

                //Supply required fields (Date)
                client.Browser.Driver.FindElement(By.Id("DatePicker11-label")).Click();
                xrmApp.ThinkTime(5000);
                client.Browser.Driver.FindElement(By.Id("DatePicker11-label")).Click();
                xrmApp.ThinkTime(5000);
                client.Browser.Driver.FindElement(By.Id("DatePicker11-label")).Clear();
                xrmApp.ThinkTime(5000);
                client.Browser.Driver.FindElement(By.Id("DatePicker11-label")).SendKeys(DateTime.Now.ToString("d/M/yyyy"));
                xrmApp.ThinkTime(5000);

                xrmApp.QuickCreate.SetValue(new LookupItem { Name = "msdyn_project", Value = "Project v4" });
                xrmApp.ThinkTime(5000);

                //Save time entry
                client.Browser.Driver.FindElement(By.XPath("//button[@id='Main_Button']")).Click();
                xrmApp.ThinkTime(10000);
            }

        }

        [TestMethod]
        public void TC01_02_UpdateAndSubmitTimeEntry()
        {
            using (xrmApp)
            {
                //Launch and login to D365 as User
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Add page delay
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                //Navigate to Project Service
                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                xrmApp.ThinkTime(10000);

                //Navigate to Projects area and Time Entries subarea
                xrmApp.Navigation.OpenSubArea("Projects", "Time Entries");
                xrmApp.ThinkTime(5000);

                //Switch view to My Time Entries by Date
                xrmApp.Grid.SwitchView("My Time Entries by Date");
                xrmApp.ThinkTime(10000);

                //Select a record in the grid for editing
                client.Browser.Driver.FindElement(By.XPath("//div[@class='wj-row'] and //label[contains(text(),'Draft')][1]")).Click();
                //client.Browser.Driver.ExecuteScript("var script = document.createElement('script');script.src='https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js';document.getElementsByTagName('head')[0].appendChild(script);");
                xrmApp.ThinkTime(25000);

                //client.Browser.Driver.ExecuteScript("$(document).ready(function(){$('div#Grid79784c7c-cb45-ad19-9329-dde9bed81a0b-id-cell-0-1').dblclick();});");
                //xrmApp.ThinkTime(25000);

                xrmApp.Entity.SelectLookup(new LookupItem { Name = "msdyn_project" });

                xrmApp.Entity.SetValue(new LookupItem { Name = "msdyn_project", Value = "Project v4" });
                xrmApp.Entity.Save();

                xrmApp.ThinkTime(5000);
            }
        }

        [TestMethod]
        public void TC01_03_SubmitTimeEntry()
        {
            using (xrmApp)
            {
                //Launch and login to D365 as User
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Add page delay
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                //Navigate to Project Service
                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                xrmApp.ThinkTime(5000);

                //Navigate to Projects area and Time Entries subarea
                xrmApp.Navigation.OpenSubArea("Projects", "Time Entries");
                xrmApp.ThinkTime(5000);

                //Switch view to My Time Entries by Date
                xrmApp.Grid.SwitchView("My Time Entries by Date");
                xrmApp.ThinkTime(5000);

                //Inject Script
                client.Browser.Driver.ExecuteScript("var script = document.createElement('script');script.src='https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js';document.getElementsByTagName('head')[0].appendChild(script);");
                xrmApp.ThinkTime(5000);

                //Click on filter by status
                client.Browser.Driver.ExecuteScript("$('div#msdyn_entrystatus:contains(Entry Status) span').click()");
                xrmApp.ThinkTime(5000);
                client.Browser.Driver.ExecuteScript("$('div#msdyn_entrystatus:contains(Entry Status) span').click()");
                xrmApp.ThinkTime(5000);

                //Click on Select All checkbox
                client.Browser.Driver.ExecuteScript("$('div.wj-listbox-item:contains(Select All) input').click()");
                xrmApp.ThinkTime(5000);

                //Click on Draft checkbox
                client.Browser.Driver.ExecuteScript("$('div.wj-listbox-item:contains(Draft) input').click()");
                xrmApp.ThinkTime(5000);

                //Click on Apply button
                client.Browser.Driver.ExecuteScript("$('button.wj-btn:contains(Apply)').click()");
                xrmApp.ThinkTime(5000);

                //Click on Save and Close button
                client.Browser.Driver.FindElement(By.XPath("//div[@data-id='cell-0-1']")).Click();
                xrmApp.ThinkTime(5000);

                //Submit
                xrmApp.CommandBar.ClickCommand("Submit");
                xrmApp.ThinkTime(15000);
            }
        }

        [TestMethod]
        public void TC01_04_RecallUpdateSubmitTimeEntry()
        {
            using (xrmApp)
            {
                //Launch and login to D365 as User
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Add page delay
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                //Navigate to Project Service
                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                xrmApp.ThinkTime(5000);

                //Navigate to Projects area and Time Entries subarea
                xrmApp.Navigation.OpenSubArea("Projects", "Time Entries");
                xrmApp.ThinkTime(5000);

                //Switch view to My Time Entries by Date
                xrmApp.Grid.SwitchView("My Time Entries by Date");
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.FindElement(By.XPath("//div[@data-id='cell-3-1']")).Click();
                xrmApp.ThinkTime(5000);

                //Recall
                xrmApp.CommandBar.ClickCommand("Recall");
                xrmApp.ThinkTime(15000);

                //Update and Save
                //xrmApp.Entity.SelectLookup(new LookupItem { Name = "msdyn_project" });
                //xrmApp.Entity.SetValue(new LookupItem { Name = "msdyn_project", Value = "Project v5" });
                //xrmApp.Entity.Save();

                //xrmApp.ThinkTime(5000);

                //Submit
                //client.Browser.Driver.FindElement(By.XPath("//span[contains(text(),'Submit')]")).Click();
                //xrmApp.ThinkTime(5000);
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
