using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA007_CreateNewLeads
{

    [TestClass]
    public class CreateNewLeads
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
        public void TC03_01_CreateNewLeadRecord()
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

                xrmApp.Navigation.OpenSubArea("Sales", "Leads");
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("New");
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.SetValue("firstname", "Test");
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.SetValue("lastname", "User" + rndLastname);
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
