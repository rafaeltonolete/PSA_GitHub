using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA005_CreateNewActivities
{


    [TestClass]
    public class CreateNewActivities
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

        /**/

        [TestMethod]
        public void TC01_CreateNewTask()
        {


            using (xrmApp)
            {
                //LogIn
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Enter code during this delay then proceed to CRM welcome page
                xrmApp.ThinkTime(30000);


                //Navigate to Sales Hub
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Navigate to Activities
                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click Task from command button ribbon
                xrmApp.CommandBar.ClickCommand("Task");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter subject
                xrmApp.Entity.SetValue("subject", "Test Project Review");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Select Project Name on Regarding field
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "regardingobjectid" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "regardingobjectid", Value = "Project 3QA" });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Save Task
                xrmApp.Entity.Save();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


            }


        }



        /**/
        [TestMethod]
        public void TC01_CreateNewEmail()
        {


            using (xrmApp)
            {
                //LogIn
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Enter code during this delay then proceed to CRM welcome page
                xrmApp.ThinkTime(30000);


                //Navigate to Sales Hub
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Navigate to Activities
                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click Email from Command button ribbon
                xrmApp.CommandBar.ClickCommand("Email");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Enter Recipient
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "to" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "to", Value = "Tristan Perper" });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter Subject
                xrmApp.Entity.SetValue("subject", "Request For Your Action");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                String myMessage = "Hi,\n\n\nKindly Review Time Entry.\nThen submit a reply to confirm.\n\n\nRegards,\n\nSender";
                myMessage = myMessage.Replace("\n", Environment.NewLine);

                /*parent iframe that contains the Email message body*/
                IWebElement parentIframe = client.Browser.Driver.FindElement(By.XPath(".//iframe[@class='fullPageContentEditorFrame']"));
                client.Browser.Driver.SwitchTo().Frame(parentIframe);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                /*child iframe where the editable field is located*/
                IWebElement frame1 = client.Browser.Driver.FindElement(By.XPath(".//iframe[@class='cke_wysiwyg_frame cke_reset']"));
                client.Browser.Driver.SwitchTo().Frame(frame1);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Enter Message
                client.Browser.Driver.FindElement(By.XPath("//body[contains(@id,'FullPage_')]")).SendKeys(myMessage);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                /*takes back webdriver focus to main content*/
                client.Browser.Driver.SwitchTo().DefaultContent();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                client.Browser.Driver.FindElement(By.XPath(".//label[text()='Duration']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter regarding
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "regardingobjectid" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "regardingobjectid", Value = "Project 3QA" });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Save Email
                xrmApp.CommandBar.ClickCommand("Save");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


            }


        }


        /**/
        [TestMethod]
        public void TC03_CreateNewAppointment()
        {


            using (xrmApp)
            {

                //LogIn
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                //Enter code during this delay then proceed to CRM welcome page
                xrmApp.ThinkTime(30000);


                //Navigate to Sales Hub
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Navigate to Activities
                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click Appointment from command button ribbon
                xrmApp.CommandBar.ClickCommand("Appointment");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter Required attendee
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "requiredattendees" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "requiredattendees", Value = "Tristan Perper", Index = 0 });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //xrmApp.Entity.SetValue(new LookupItem { Name = "requiredattendees", Value = "Rafael Paolo Tonolete" , Index = 1});
                //client.Browser.Driver.WaitForPageToLoad();
                //xrmApp.ThinkTime(5000);

                client.Browser.Driver.FindElement(By.XPath(".//label[text()='Optional']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Enter subject
                xrmApp.Entity.SetValue("subject", "Daily Update Discussion");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                String desCription = "Purpose:\n\nTo Discuss Updates of Project on daily basis.";
                desCription = desCription.Replace("\n", Environment.NewLine);

                //Enter Description
                client.Browser.Driver.FindElement(By.XPath(".//textarea[@aria-label='Description']")).SendKeys(Keys.Backspace);
                client.Browser.Driver.FindElement(By.XPath(".//textarea[@aria-label='Description']")).SendKeys(Keys.Backspace);
                client.Browser.Driver.FindElement(By.XPath(".//textarea[@aria-label='Description']")).SendKeys(Keys.Backspace);
                client.Browser.Driver.FindElement(By.XPath(".//textarea[@aria-label='Description']")).SendKeys(desCription);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Save
                xrmApp.CommandBar.ClickCommand("Save");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


            }


        }


        [TestMethod]
        public void TC02_04_CreateNewPhoneCall()
        {
            using (xrmApp)
            {
                string rndPhoneCall = TestSettings.GetRandomString(5, 10);

                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.ThinkTime(5000);

                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("Phone Call");
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.SetValue("subject", "Auto Phone Call " + rndPhoneCall);
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.SelectLookup(new LookupItem { Name = "to" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "to", Value = "Test Contact", Index = 0 });
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("Save & Close");
                xrmApp.ThinkTime(5000);
            }
        }

        [TestMethod]
        public void TC02_05_CreateOtherActivities()
        {
            using (xrmApp)
            {
                string rndEmailCampaign = TestSettings.GetRandomString(5, 10);

                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.ThinkTime(5000);

                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("Other Activities", "Email Campaign");
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.SetValue("subject", "Email Campaign " + rndEmailCampaign);
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
