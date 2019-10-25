using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Sample;
using OpenQA.Selenium;

namespace PSA006_UpdateActivities
{

    [TestClass]
    public class UpdateActivities
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


        /*Given that a Task record exists*/
        [TestMethod]
        public void TC06_UpdateTask()
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

                //Search for Task to Update
                xrmApp.Grid.Search("Test Project Review");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Select  Task listed
                client.Browser.Driver.FindElement(By.XPath(".//div[@data-id='cell-0-1']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Click Edit
                xrmApp.CommandBar.ClickCommand("Edit");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter new Subject
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@aria-label,'Subject')]")).Clear();
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@aria-label,'Subject')]")).Click();
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@aria-label,'Subject')]")).SendKeys("Test Project Review V1");
                //xrmApp.Entity.SetValue("subject","Test Project Review V1");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter regarding
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "regardingobjectid" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "regardingobjectid", Value = "Project 3QA" });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Save
                xrmApp.Entity.Save();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


            }


        }




        /*Given there is a draft email record*/
        [TestMethod]
        public void TC07_UpdateEmail()
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

                //Switch Grid view
                xrmApp.Grid.SwitchView("My Draft Emails");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Open email to edit
                xrmApp.Grid.OpenRecord(0);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Change Recipient
                //xrmApp.Entity.SelectLookup(new LookupItem { Name = "to" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "to", Value = "Tristan Perper" });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                client.Browser.Driver.FindElement(By.XPath(".//label[text()='Bcc']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Enter new Subject
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@aria-label,'Subject')]")).Clear();
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@aria-label,'Subject')]")).Click();
                client.Browser.Driver.FindElement(By.XPath(".//input[contains(@aria-label,'Subject')]")).SendKeys("Urgent: Request For Your Action");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                String myMessage = "Hi,\n\n\nKindly Review Time Entry.\nThen submit a reply to confirm.\n\n\nRegards,\n\nSender\n\n\n\n";
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


                //Enter message
                client.Browser.Driver.FindElement(By.XPath("//body[contains(@id,'FullPage_')]")).SendKeys(myMessage);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.SwitchTo().DefaultContent();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Save
                xrmApp.CommandBar.ClickCommand("Save");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


            }


        }


        /**/
        [TestMethod]
        public void TC08_UpdateAppointment()
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

                //Switch Grid view
                xrmApp.Grid.SwitchView("My Appointments");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Open record to edit
                xrmApp.Grid.OpenRecord(0);
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                //Enter required attendee
                xrmApp.Entity.SelectLookup(new LookupItem { Name = "requiredattendees" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "requiredattendees", Value = "Tristan Perper", Index = 1 });
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.FindElement(By.XPath(".//label[text()='Optional']")).Click();
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                //Enter subject
                client.Browser.Driver.FindElement(By.XPath(".//input[@aria-label='Subject']")).Clear();
                client.Browser.Driver.FindElement(By.XPath(".//input[@aria-label='Subject']")).Click();
                client.Browser.Driver.FindElement(By.XPath(".//input[@aria-label='Subject']")).SendKeys("Update Discussion");
                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(5000);


                String desCription = "Purpose:\n\nTo Discuss Updates of Project on daily basis.";
                desCription = desCription.Replace("\n", Environment.NewLine);


                //Enter description
                client.Browser.Driver.FindElement(By.XPath(".//textarea[@aria-label='Description']")).Clear();
                client.Browser.Driver.FindElement(By.XPath(".//textarea[@aria-label='Description']")).Click();
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
        public void TC02_09_UpdatePhoneCall()
        {
            using (xrmApp)
            {
                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(5000);

                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.ThinkTime(5000);

                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                xrmApp.ThinkTime(5000);

                xrmApp.Grid.SwitchView("All Phone Calls");
                xrmApp.ThinkTime(5000);

                xrmApp.Grid.OpenRecord(0);
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.ClearValue("phonenumber");
                xrmApp.ThinkTime(5000);
                xrmApp.Entity.SetValue("phonenumber", "1234567");
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.Save();
                xrmApp.ThinkTime(5000);
            }
        }

        [TestMethod]
        public void TC02_10_UpdateOtherActivities()
        {
            using (xrmApp)
            {

                string rndEmailCampaign = TestSettings.GetRandomString(5, 10);

                xrmApp.OnlineLogin.Login(_xrmuri, _username, _password);
                xrmApp.ThinkTime(30000);

                client.Browser.Driver.WaitForPageToLoad();
                xrmApp.ThinkTime(30000);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.ThinkTime(5000);

                xrmApp.Navigation.OpenSubArea("Sales", "Activities");
                xrmApp.ThinkTime(5000);

                xrmApp.Grid.Search("Test Campaign");
                xrmApp.ThinkTime(5000);

                xrmApp.Grid.OpenRecord(0);
                xrmApp.ThinkTime(5000);

                xrmApp.Entity.ClearValue("subject");
                xrmApp.ThinkTime(5000);
                xrmApp.Entity.SetValue("subject", "Test Campaign Update " + rndEmailCampaign);
                xrmApp.ThinkTime(5000);

                xrmApp.CommandBar.ClickCommand("Save");
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
