using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA008_UpdateExistingLeads
{
    [TestClass]
    public class UpdateExistingLeads
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
        public void TC03_02_UpdateExistingLeadRecord()
        {
            using (xrmApp)
            {

                string rndTopic = TestSettings.GetRandomString(5, 10);

                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                xrmApp.ThinkTime(5000);

                xrmApp.Navigation.OpenSubArea("Sales", "Leads");
                xrmApp.ThinkTime(5000);

                xrmApp.Grid.SwitchView("My Open Leads");
                xrmApp.ThinkTime(10000);

                client.Browser.Driver.FindElement(By.XPath("//div[@data-id='cell-0-1']")).Click();
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("Edit");
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.ClearValue("subject");
                xrmApp.ThinkTime(5000);
                xrmApp.Entity.SetValue("subject", "Contact Topic Update" + rndTopic);
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.Save();
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
