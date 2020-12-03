namespace TimeUnityPortal
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Configuration;
    using System.Diagnostics;
    using System.Drawing.Imaging;
    using System.Net;
    using System.Net.Mail;
    //using System.Web.Mail;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Support.Events;
    using OpenQA.Selenium.Support.Extensions;
    //using OpenQA.Selenium.Support.PageObjects;   //Obsolete
    using SeleniumExtras.PageObjects;
    using Protractor;
    using TechTalk.SpecFlow;
    using TimeUnityPortal.Pages;
    using OpenQA.Selenium.Support.UI;
    using System.Threading;

    class UIHelper : TestBase
    {
        public static string emailOnError = ConfigurationManager.AppSettings["ERROR_EMAIL"];
        public static int TriageFlag = 0;
        public static int InformationFlag = 0;
        public static int ReferralFlag = 0;
        public static int IndividualFlag = 0;

        public static T PageInit<T>(NgWebDriver driver) where T : class, new()
        {
            var i = Int32.Parse(ScenarioContext.Current["stepcounter"].ToString());
            var log = ScenarioContext.Current.Get<TextWriterTraceListener>("report");
            var page = new T();
            PageFactory.InitElements(driver, page);
            log.WriteLine("-----------------------------------------------------------");
            ScenarioContext.Current["stepcounter"] = i;

            return page;
        }

        public static void GoTo<T>(string host, bool isAngular) where T : IPage, new()
        {
            var i = int.Parse(ScenarioContext.Current["stepcounter"].ToString());
            var log = ScenarioContext.Current.Get<TextWriterTraceListener>("report");
            var page = new T();
            var url = host + page.Url;
            if (isAngular)
            {
                driver.IgnoreSynchronization = false;
                driver.Navigate().GoToUrl(url);
            }
            else
            {
                driver.WrappedDriver.Navigate().GoToUrl(url);
            }

            log.WriteLine(i++ + ". Then I go to " + driver.Title + " page: " + driver.Url);
            //log.WriteLine(". Then I go to " + driver.Title + " page: " + driver.Url);
            log.WriteLine("-----------------------------------------------------------");
            ScenarioContext.Current["stepcounter"] = i;
        }

        public static IWebElement ElementIsReady(IWebElement webElement)
        {
            var isReady = webElement.Displayed && webElement.Enabled.Equals(true);
            while (isReady.Equals(false))
            {
                isReady = webElement.Displayed;
            }

            return webElement;
        }

        public static Func<IWebDriver, IWebElement> ElementIsClickable(IWebElement webElement)
        {
            return dr => (webElement.Displayed && webElement.Enabled) ? webElement : null;
        }

        public static IWebDriver ClickOnLink(IWebElement webElement)
        {
            var i = Int32.Parse(ScenarioContext.Current["stepcounter"].ToString());
            var log = ScenarioContext.Current.Get<TextWriterTraceListener>("report");
            log.WriteLine(i++ + ". And I clicked on '" + webElement.Text + "' link");
            //log.WriteLine(". And I clicked on '" + webElement.Text + "' link");
            Wait.Until(ElementIsClickable(webElement)).Click();
            log.WriteLine("-----------------------------------------------------------");
            log.Flush();
            log.Close();
            ScenarioContext.Current["stepcounter"] = i;
            return driver;
        }

        public static IWebDriver ClickOnButton(IWebElement webElement)
        {
            var i = Int32.Parse(ScenarioContext.Current["stepcounter"].ToString());
            var log = ScenarioContext.Current.Get<TextWriterTraceListener>("report");
            log.WriteLine(i++ + ". And I clicked on '" + webElement.Text + "' button");
            //log.WriteLine(". And I clicked on '" + webElement.Text + "' button");
            webElement.Click();
            log.WriteLine("-----------------------------------------------------------");
            log.Flush();
            log.Close();
            ScenarioContext.Current["stepcounter"] = i;
            return driver;
        }

        public static void SetCheckbox(IWebElement webElement, string status)
        {
            var log = ScenarioContext.Current.Get<TextWriterTraceListener>("report");
            var i = Int32.Parse(ScenarioContext.Current["stepcounter"].ToString());
            if (webElement.Selected)
            {
                if (!status.Equals("OFF"))
                {
                    return;
                }

                webElement.Click();
                log.WriteLine(i++ + ".  Then I set checkbox '" + webElement.GetAttribute("id") + "' to *OFF*");
                //log.WriteLine(".  Then I set checkbox '" + webElement.GetAttribute("id") + "' to *OFF*");
                log.WriteLine(".  Then I set checkbox '" + webElement.GetAttribute("id") + "' to *OFF*");
                log.WriteLine("-----------------------------------------------------------");
            }
            else
            {
                if (!status.Equals("ON"))
                {
                    return;
                }

                webElement.Click();
                //log.WriteLine(i++ + ".  Then I set checkbox '" + webElement.GetAttribute("id") + "' to *ON*");
                //log.WriteLine(".  Then I set checkbox '" + webElement.GetAttribute("id") + "' to *ON*");
                log.WriteLine("-----------------------------------------------------------");
            }
        }

        public static void EnterText(IWebElement webElement, string text)
        {
            webElement = new NgWebElement(driver, webElement);
            var log = ScenarioContext.Current.Get<TextWriterTraceListener>("report");
            var i = Int32.Parse(ScenarioContext.Current["stepcounter"].ToString());
            var placeholder = webElement.GetAttribute("placeholder");

            if (placeholder.Length == 0)
            {
                placeholder = webElement.GetAttribute("id");
            }

            if (text.Length > 0)
            {
                if (text[text.Length - 1] > (char)57349)
                {
                    log.WriteLine(
                        i++ + ". And I entered '" + text.TrimEnd().Substring(0, text.Length - 1) + "' in '"
                        + placeholder + "' field");
                    log.WriteLine("-----------------------------------------------------------");
                    log.WriteLine(i++ + ". Then I pressed 'ENTER'");
                }
                else
                {
                    log.WriteLine(i++ + ". And I entered '" + text + "' in '" + placeholder + "' field");
                }
            }
            else
            {
                log.WriteLine(i++ + ". And I left field '" + placeholder + "' blank");
            }
            //webElement.Click();
            webElement.Clear();
            webElement.SendKeys(text);
            log.WriteLine("-----------------------------------------------------------");
            log.Close();
            ScenarioContext.Current["stepcounter"] = i;
        }

        public static void SelectRandomComboElement(IWebDriver driver, IWebElement Table)
        {
            Random rnd = new Random();
            IndividualFlag = 0;

            var itemCount = Table.FindElements(By.TagName("tr")).Count();
            IList<IWebElement> allOptions = Table.FindElements(By.TagName("tr"));

            if ((allOptions[0].Text == " ") || (allOptions[0].Text == ""))
            {
                if (itemCount == 1)
                {
                    allOptions[0].Click();
                }
                else
                {
                    var index = rnd.Next(1, itemCount - 1);

                    if (itemCount > 2)
                    {
                        allOptions[index].Click();
                    }

                    else
                    {
                        if (itemCount > 1)
                        {
                            allOptions[1].Click();
                        }
                    }
                }
            }

            else
            {
                var index = rnd.Next(0, itemCount - 1);
                var elementText = allOptions[index].Text;
                if (elementText == "Triage")
                {
                    TriageFlag = 1;
                }
                if (elementText == "Information")
                {
                    InformationFlag = 1;
                }
                if (elementText == "Referral")
                {
                    ReferralFlag = 1;
                }
                if (elementText == "Individual")
                {
                    IndividualFlag = 1;
                }
                allOptions[index].Click();
            }
        }

        public static void SelectRandomCheckboxElement(IWebDriver driver, IWebElement Table)
        {
            Random rnd = new Random();

            var itemCount = Table.FindElements(By.TagName("tr")).Count();
            IList<IWebElement> allOptions = Table.FindElements(By.TagName("tr"));
            var index = rnd.Next(1, itemCount - 1);

            if (itemCount > 2)
            {
                IList<IWebElement> OptionParts = allOptions[index].FindElements(By.TagName("td"));
                OptionParts[0].Click();
            }

            else
            {
                if (itemCount > 1)
                {
                    IList<IWebElement> OptionParts = allOptions[1].FindElements(By.TagName("td"));
                    OptionParts[0].Click();
                }
            }
        }

        public static int GetFrameIndex(IWebDriver driver, String id)
        {
            driver.SwitchTo().DefaultContent();
            int size = driver.FindElements(By.TagName("iframe")).Count;

            var i = 0;
            for (i = 0; i < size;)
            {
                driver.SwitchTo().Frame(i);
                //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromMilli‌​seconds(0));
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilli‌​seconds(0);
                int total = driver.FindElements(By.Id(id)).Count;
                if (total > 0)
                {
                    driver.SwitchTo().DefaultContent();
                    break;
                }
                else
                {
                    driver.SwitchTo().DefaultContent();
                    i++;
                }
            }

            return i;
        }

        public static bool SearchGridAndVerify(IWebDriver driver, IWebElement Table, String verify, int columnNumber)
        {
            var itemCount = Table.FindElements(By.TagName("tr")).Count();
            int flag = 0;
            IList<IWebElement> allOptions = Table.FindElements(By.TagName("tr"));
            IList<IWebElement> OptionParts = allOptions[1].FindElements(By.TagName("td"));

            if (OptionParts[0].Text.Contains("No data to display"))
            {
                return false;
            }

            else
            {
                for (int i = 1; i < itemCount;)
                {
                    IList<IWebElement> Columns = allOptions[i].FindElements(By.TagName("td"));
                    if (Columns[columnNumber].Text.Contains(verify))
                    {
                        flag = 1;
                        Columns[0].Click();
                        break;
                    }
                    else
                    {
                        i++;
                    }
                }

                if (flag == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
        }

        public static void ScrollToElement(IWebDriver Driver, IWebElement element)
        {
            var js = Driver as IJavaScriptExecutor;
            var y = element.Location.Y;
            js.ExecuteScript("javascript:window.scrollBy(0," + y + ")");
        }

        //public static bool EmailRecieved()
        //{
        //    var api = new WebApiServiceHelper();
        //    var requestString = string.Format("https://api.mailinator.com/api/inbox?to={0}&token={1}", ScenarioContext.Current["userName"], "7c5b21b2160b42269075e44ff3b7987f");
        //    return api.Get<MailinatorInbox>(requestString).messages.Count > 0;
        //}

        public static string TakeScreenshot()
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd-hhmm-ss");
            var fileName = "Exception-" + timestamp + ".png";
            var firingDriver = new EventFiringWebDriver(driver.WrappedDriver);
            //firingDriver.TakeScreenshot().SaveAsFile(fileName, ImageFormat.Png);
            firingDriver.TakeScreenshot().SaveAsFile(fileName, ScreenshotImageFormat.Png);
            return fileName;
        }

        public static void SendEmailOnErrorWithAttachment(string fileName)
        {
            //var timestamp = DateTime.Now.ToString("yyyy-MM-dd-hhmm-ss");
            //var mail = new MailMessage();
            //SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            //mail.From = new MailAddress("mahmudautomateerror@gmail.com");
            //mail.To.Add(emailOnError);
            //mail.Subject = timestamp + ": " + ScenarioContext.Current.ScenarioInfo.Title + " scenario has failed.";
            //mail.Body = ScenarioContext.Current.ScenarioInfo.Title
            //            + " scenario has failed. Please view attachment and attachment name for time when this occured. ";
            //var attachment = new Attachment(fileName);
            //mail.Attachments.Add(attachment);
            //attachment = new Attachment(ScenarioContext.Current["report_file_name"].ToString());
            //mail.Attachments.Add(attachment);
            ///*SmtpServer.Port = 587; //TLS
            ////SmtpServer.Port = 465; //SSL
            ////SmtpServer.Port = 25;
            //SmtpServer.Credentials = new NetworkCredential("test@itmagnet.com.au", "itm2015");
            ////SmtpServer.EnableSsl = true;
            //SmtpServer.EnableSsl = true;
            //SmtpServer.Send(mail);*/

            //SmtpClient smtp = new SmtpClient();
            //smtp.Host = "localhost";
            //smtp.Port = 465;
            //smtp.Send(mail);
        }

        public static void SendEmailonError(string FromEmail, string ToEmail, string Subject, string Body)
        {
            var timeStamp = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress(FromEmail);
            mail.To.Add(ToEmail);
            mail.Subject = Subject;
            mail.Body = Body;

            /*System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment("c:/textfile.txt");
            mail.Attachments.Add(attachment);*/

            SmtpServer.Port = 587;
            //SmtpServer.Port = 465;
            SmtpServer.Credentials = new System.Net.NetworkCredential(FromEmail, "facemm-12");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
        }

        //public static void SendEmailonError(string FromEmail, string ToEmail, string Subject, string Body)
        //{
        //    try
        //    {
        //        System.Web.Mail.MailMessage myMail = new System.Web.Mail.MailMessage();
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", "smtp.gmail.com");
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", "465");
        //        //myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", "587");
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", "2");
        //        //sendusing: cdoSendUsingPort, value 2, for sending the message using 
        //        //the network.

        //        //smtpauthenticate: Specifies the mechanism used when authenticating 
        //        //to an SMTP 
        //        //service over the network. Possible values are:
        //        //- cdoAnonymous, value 0. Do not authenticate.
        //        //- cdoBasic, value 1. Use basic clear-text authentication. 
        //        //When using this option you have to provide the user name and password 
        //        //through the sendusername and sendpassword fields.
        //        //- cdoNTLM, value 2. The current process security context is used to 
        //        // authenticate with the service.
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1");
        //        //Use 0 for anonymous
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", FromEmail);
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "facemm-12");
        //        myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", "true");
        //        myMail.From = FromEmail;
        //        myMail.To = ToEmail;
        //        myMail.Subject = Subject;
        //        //myMail.BodyFormat = pFormat;
        //        myMail.Body = Body;
        //        /*if (pAttachmentPath.Trim() != "")
        //        {
        //            MailAttachment MyAttachment =
        //                    new MailAttachment(pAttachmentPath);
        //            myMail.Attachments.Add(MyAttachment);
        //            myMail.Priority = System.Web.Mail.MailPriority.High;
        //        }*/

        //        System.Web.Mail.SmtpMail.SmtpServer = "smtp.gmail.com:465";
        //        //System.Web.Mail.SmtpMail.SmtpServer = "smtp.gmail.com:587";
        //        System.Web.Mail.SmtpMail.Send(myMail);
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e);
        //        throw;
        //    }
        //}
    }
}
