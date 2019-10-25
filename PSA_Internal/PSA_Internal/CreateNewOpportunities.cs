using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA009_CreateNewOpportunity
{

    [TestClass]
    public class CreateNewOpportunity
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
        public void TC04_01_CreateNewOpportunityRecord()
        {
            using (xrmApp)
            {

                string rndLastname = TestSettings.GetRandomString(5, 10);

                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                xrmApp.ThinkTime(5000);

                xrmApp.Navigation.OpenSubArea("Sales", "Opportunities");
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("New");
                xrmApp.ThinkTime(5000);

                //Account
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "parentaccountid" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "parentaccountid", Value = "Test 1" });

                //Topic
                xrmApp.Entity.SetValue("name", "OpportunityTopic" + rndLastname);
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.FindElement(By.XPath("//body")).SendKeys(Keys.ArrowDown);
                xrmApp.ThinkTime(5000);


                //Contact
                for (int i = 0; i < 5; i++)
                {
                    client.Browser.Driver.FindElement(By.XPath("//body")).SendKeys(Keys.ArrowDown);
                }
                xrmApp.ThinkTime(5000);

                //SOLUTION 1
                //xrmApp.Entity.SelectLookup(new LookupItem { Name = "parentcontactid" });
                //xrmApp.ThinkTime(5000);
                //xrmApp.Entity.SetValue(new LookupItem { Name = "parentcontactid", Value = "AutoContact" });
                //xrmApp.ThinkTime(30000);

                //SOLUTION 2
                //var parentcontactId = new LookupItem { Name = "parentcontactid", Value = "Test Contact", Index = 0 };
                //xrmApp.Entity.SelectLookup(parentcontactId);
                //xrmApp.ThinkTime(30000);

                //SOLUTION 3
                client.Browser.Driver.FindElement(By.XPath("//input[contains(@data-id, 'parentcontactid.fieldControl')]")).Click();
                xrmApp.ThinkTime(5000);
                client.Browser.Driver.FindElement(By.XPath("//input[contains(@data-id, 'parentcontactid.fieldControl')]")).SendKeys("Test Contact");
                xrmApp.ThinkTime(10000);
                client.Browser.Driver.FindElement(By.XPath("//span[contains(text(),'Test Contact')]")).Click();
                xrmApp.ThinkTime(10000);

                client.Browser.Driver.FindElement(By.XPath("//body")).Click();
                xrmApp.ThinkTime(5000);

                for (int i = 0; i < 6; i++)
                {
                    client.Browser.Driver.FindElement(By.XPath("//body")).SendKeys(Keys.ArrowDown);
                }
                xrmApp.ThinkTime(5000);

                //Location

                //SOLUTION 1
                //xrmApp.Entity.SetValue("a1a_location", "Makati");
                //xrmApp.ThinkTime(30000);

                //SOLUTION 2
                client.Browser.Driver.FindElement(By.XPath("//select[contains(@data-id, 'a1a_location.fieldControl-option-set-select')]")).Click();
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.FindElement(By.XPath("//option[contains(text(),'Makati')]")).Click();
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("Save & Close");
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
