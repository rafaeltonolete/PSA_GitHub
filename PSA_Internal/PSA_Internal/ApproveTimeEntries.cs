using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA003_ApproveTimeEntries
{
    [TestClass]
    public class ApproveTimeEntries
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


        /*Given that there are submitted time entries available for Approval*/

        [TestMethod]
        public void TC06_ApproveTimeEntries()
        {


            using (xrmApp)
            {
                //LogIn
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Enter code during this delay then proceed to CRM welcome page
                xrmApp.ThinkTime(30000);


                //Navigate to Project Service
                xrmApp.Navigation.OpenApp(UCIAppName.ProjectService);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Navigate to TimeEntries Approvals list
                xrmApp.Navigation.OpenSubArea("Projects", "Approvals");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Select the 1st row Time Entry
                client.Browser.Driver.FindElement(By.XPath(".//div[@data-id='cell-0-1']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click Approve 
                xrmApp.CommandBar.ClickCommand("Approve");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click Ok button on pop up
                client.Browser.Driver.FindElement(By.XPath(". //button[@id='okButton']")).Click();
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
