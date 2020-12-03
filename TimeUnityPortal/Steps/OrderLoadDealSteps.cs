using System;
using System.Globalization;
using System.Threading;
using NUnit.Framework;
using TechTalk.SpecFlow;
using TimeUnityPortal.Pages;

namespace TimeUnityPortal.Steps
{
    [Binding]
    public class OrderLoadDealSteps : TestBase
    {
        public static PortalLoginPage portalLoginPage;
        public static OrderPage orderPage;

        [Given(@"I go to portal login page")]
        public void GivenIGoToPortalLoginPage()
        {
            if (FeatureContext.Current.ContainsKey("error"))
            {
                return;
            }

            UIHelper.GoTo<PortalLoginPage>(HostUrl, false);
        }

        [Then(@"I provide username, password, deal Urn and login")]
        public void ThenIProvideUsernamePasswordDealUrnAndLogin()
        {
            portalLoginPage = UIHelper.PageInit<PortalLoginPage>(driver);
            UIHelper.EnterText(portalLoginPage.UserNameText, Username);
            UIHelper.EnterText(portalLoginPage.PasswordText, Password);
            UIHelper.EnterText(portalLoginPage.FctUrnText, IntegrationRequestsSteps.dealUrn);
            UIHelper.EnterText(portalLoginPage.BusinessRoleText, "LAWYER");
            UIHelper.EnterText(portalLoginPage.FctUsernameText, "");
            UIHelper.EnterText(portalLoginPage.FirstNameText, "Service");
            UIHelper.EnterText(portalLoginPage.LastNameText, "Tester");
            UIHelper.EnterText(portalLoginPage.PartnerUsernameText, "stester");
            UIHelper.ClickOnLink(portalLoginPage.LoginBtn);
            Thread.Sleep(3000);
        }

        [Then(@"I verify I am logged in")]
        public void ThenIVerifyIAmLoggedIn()
        {
            orderPage = UIHelper.PageInit<OrderPage>(driver);
            UIHelper.ElementIsClickable(orderPage.NeedHelpLink);
            UIHelper.ElementIsClickable(orderPage.LegalLink);
            UIHelper.ElementIsClickable(orderPage.PrivacyPolicyLink);
            //Assert.True(orderPage.PortalFieldIdentifierMessage.Text.Length>0);
            Assert.True(orderPage.PortalFieldIdentifierMessage.Text.Contains("Field(s) to be completed below"));
            Assert.True(orderPage.UnityFieldIdentifierMessage.Text.Contains("Field(s) to be completed in Unity"));
            Assert.True(orderPage.FooterText.Text.Contains("® Registered trademark of First American Financial Corporation"));
        }

        [Then(@"I verify the deal is loaded")]
        public void ThenIVerifyTheDealIsLoaded()
        {
            orderPage = UIHelper.PageInit<OrderPage>(driver);
            UIHelper.ElementIsClickable(orderPage.SubmittedMilestoneLbl);
            Assert.True(orderPage.SubmittedMilestoneLbl.Text.Contains("Submitted"));
            UIHelper.ElementIsClickable(orderPage.PolicyProcessingMilestoneLbl);
            Assert.True(orderPage.PolicyProcessingMilestoneLbl.Text.Contains("Policy Processing"));
            UIHelper.ElementIsClickable(orderPage.CompletedMilestoneLbl);
            Assert.True(orderPage.CompletedMilestoneLbl.Text.Contains("Completed"));
            UIHelper.ElementIsClickable(orderPage.TransactionDetailsHeader);
            Assert.True(orderPage.TransactionDetailsHeader.Text.Contains("Transaction Details"));
        }

        [Then(@"I verify deal data is accurate compared to Send Order request data")]
        public void ThenIVerifyDealDataIsAccurateComparedToSendOrderRequestData()
        {
            orderPage = UIHelper.PageInit<OrderPage>(driver);
            Assert.AreEqual(orderPage.ContactNameValue.Text, IntegrationRequestsSteps.ContactFirstName + " " + IntegrationRequestsSteps.ContactLastName, "Contact name value does not match with integration request data");
            //Assert.AreEqual(orderPage.MMSTransactionValue.Text.ToLower(), IntegrationRequestsSteps.MMSDeal, "MMS Transaction value does not match with integration request data");
            Assert.AreEqual(orderPage.FileNoValue.Text, IntegrationRequestsSteps.FileNumber, "File No. value does not match with integration request data");

            if ((IntegrationRequestsSteps.TransactionType != "Mortgage Only") && (IntegrationRequestsSteps.TransactionType != "Existing Owner With Mortgage"))
            {
                Assert.AreEqual(orderPage.TransactionTypeValue.Text, IntegrationRequestsSteps.TransactionType, "Transaction Type value does not match with integration request data");
                Assert.AreEqual(orderPage.PurchasePriceValue.Text.Replace("$", "").Replace(",", ""), IntegrationRequestsSteps.PurchasePrice + ".00");

                Assert.AreEqual(orderPage.ClosingDateValue_Purchase.Text, DateTime.ParseExact(IntegrationRequestsSteps.ClosingDate.Substring(0, 10), "yyyy-MM-dd", CultureInfo.InvariantCulture).ToString("MMM dd, yyyy"));
            }

            if ((IntegrationRequestsSteps.TransactionType != "Purchase New") && (IntegrationRequestsSteps.TransactionType != "Purchase Resale"))
            {
                Assert.AreEqual(orderPage.ClosingDateValue_Mortgage.Text, DateTime.ParseExact(IntegrationRequestsSteps.ClosingDate.Substring(0, 10), "yyyy-MM-dd", CultureInfo.InvariantCulture).ToString("MMM dd, yyyy"));
            }
        }
    }
}
