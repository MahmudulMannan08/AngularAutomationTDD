namespace TimeUnityPortal.Pages
{
    using System.Threading;
    using NUnit.Framework;
    using OpenQA.Selenium;
    using SeleniumExtras.PageObjects;
    using Protractor;

    public class OrderPage : IPage
    {
        [FindsBy(How = How.LinkText, Using = "Need Help?")]
        public IWebElement NeedHelpLink;

        [FindsBy(How = How.LinkText, Using = "Legal")]
        public IWebElement LegalLink;

        [FindsBy(How = How.LinkText, Using = "Privacy Policy")]
        public IWebElement PrivacyPolicyLink;

        [FindsBy(How = How.XPath, Using = "//div[@class='required-message-container']//div[1]//span[1]")]
        public IWebElement PortalFieldIdentifierMessage;

        [FindsBy(How = How.XPath, Using = "//span[contains(text(),'Field(s) to be completed in Unity')]")]
        public IWebElement UnityFieldIdentifierMessage;

        //[FindsBy(How = How.ClassName, Using = "copyright-black")]
        [FindsBy(How = How.XPath, Using = "//div[@class='copyright-black']")]
        public IWebElement FooterText;

        [FindsBy(How = How.XPath, Using = "//h2[contains(text(),'Transaction Details')]")]
        public IWebElement TransactionDetailsHeader;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-deal-milestones[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/p[1]")]
        public IWebElement SubmittedMilestoneLbl;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-deal-milestones[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/p[1]")]
        public IWebElement PolicyProcessingMilestoneLbl;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-deal-milestones[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[3]/div[1]/p[1]")]
        public IWebElement CompletedMilestoneLbl;

        [FindsBy(How = How.XPath, Using = "//span[@translate='TRANSACTIONDETAIL.LawyerNotaryName']")]
        public IWebElement LawyerNameLbl;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]")]
        public IWebElement ContactNameValue;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[6]/div[1]")]
        public IWebElement MMSTransactionValue;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[3]/div[1]")]
        public IWebElement FileNoValue;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[7]/div[1]")]
        public IWebElement TransactionTypeValue;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[4]/div[1]")]
        public IWebElement PurchasePriceValue;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[5]/div[1]")]
        public IWebElement ClosingDateValue_Purchase;

        [FindsBy(How = How.XPath, Using = "/html[1]/body[1]/app-root[1]/app-master[1]/div[1]/section[1]/div[1]/div[2]/section[1]/app-order[1]/div[1]/app-transaction-detail[1]/app-panel[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[4]/div[1]")]
        public IWebElement ClosingDateValue_Mortgage;

        public string Url
        {
            get
            {
                return "";
            }
        }
    }
}
