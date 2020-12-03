namespace TimeUnityPortal.Pages
{
    using System.Threading;
    using NUnit.Framework;
    using OpenQA.Selenium;
    using SeleniumExtras.PageObjects;
    using Protractor;

    public class PortalLoginPage : IPage
    {
        [FindsBy(How = How.XPath, Using = "//input[@placeholder='User name']")]
        public IWebElement UserNameText;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='Password']")]
        public IWebElement PasswordText;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='FCT URN']")]
        public IWebElement FctUrnText;

        [FindsBy(How = How.XPath, Using = "//div[@class='action-buttons']")]
        public IWebElement LoginBtn;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='business role']")]
        public IWebElement BusinessRoleText;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='FCT username']")]
        public IWebElement FctUsernameText;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='firstName']")]
        public IWebElement FirstNameText;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='lastName']")]
        public IWebElement LastNameText;

        [FindsBy(How = How.XPath, Using = "//input[@placeholder='partnerUsername']")]
        public IWebElement PartnerUsernameText;

        public string Url
        {
            get
            {
                return "";
            }
        }
    }
}
