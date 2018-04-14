// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;
using System.Linq;
using System.Security;
using System.Threading;

namespace Microsoft.Dynamics365.UIAutomation.Api
{
    public enum LoginResult
    {
        Success,
        Failure,
        Redirect
    }

    /// <summary>
    /// Login Page
    /// </summary>
    public class LoginPage
        : XrmPage
    {

        public static bool Online { get; set; }

        public string[] OnlineDomains { get; set; }

        public LoginPage(InteractiveBrowser browser)
            : base(browser)
        {
            this.OnlineDomains = Constants.Xrm.XrmDomains;
        }

        public LoginPage(InteractiveBrowser browser, params string[] onlineDomains)
            : this(browser)
        {
            this.OnlineDomains = onlineDomains;
        }

        /// <summary>
        /// Login Page
        /// </summary>
        /// <param name="uri">The Uri</param>
        /// <param name="username">The Username to login to CRM application</param>
        /// <param name="password">The Password to login to CRM application</param>
        /// <example>xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);</example>
        public BrowserCommandResult<LoginResult> Login(Uri uri, SecureString username, SecureString password)
        {
            Online = !(this.OnlineDomains != null && !this.OnlineDomains.Any(d => uri.Host.EndsWith(d)));

            if (Online)
            {
                return this.Execute(GetOptions("Login"), this.Login, uri, username, password, default(Action<LoginRedirectEventArgs>));
            }
            else
            {
                return this.Execute(GetOptions("Login"), this.Login, uri, username, password, (Action<LoginRedirectEventArgs>)ADFSLoginAction);
            }
        }

        public void ADFSLoginAction(LoginRedirectEventArgs args)
        {
            //Login Page details go here.  You will need to find out the id of the password field on the form as well as the submit button. 
            //You will also need to add a reference to the Selenium Webdriver to use the base driver. 

            //Example
            //--------------------------------------------------------------------------------------
            var d = args.Driver;
            d.FindElement(By.Id("passwordInput")).SendKeys(args.Password.ToUnsecureString());
            d.ClickWhenAvailable(By.Id("submitButton"), new TimeSpan(0, 0, 2));
            d.WaitForPageToLoad();
            //--------------------------------------------------------------------------------------
        }

        ///// <summary>
        ///// Login Page
        ///// </summary>
        ///// <param name="uri">The Uri</param>
        ///// <param name="username">The Username to login to CRM application</param>
        ///// <param name="password">The Password to login to CRM application</param>
        ///// <param name="redirectAction">The RedirectAction</param>
        ///// <example>xrmBrowser.LoginPage.Login(_xrmUri, _username, _password, ADFSLogin);</example>
        //public BrowserCommandResult<LoginResult> Login(Uri uri, SecureString username, SecureString password, Action<LoginRedirectEventArgs> redirectAction)
        //{
        //    return this.Execute(GetOptions("Login"), this.Login, uri, username, password, redirectAction);
        //}

        private LoginResult Login(IWebDriver driver, Uri uri, SecureString username, SecureString password, Action<LoginRedirectEventArgs> redirectAction)
        {
            var redirect = false;

            if (Online)
            {
                driver.Navigate().GoToUrl(uri);

                if (driver.IsVisible(By.Id("use_another_account_link")))
                    driver.ClickWhenAvailable(By.Id("use_another_account_link"));

                driver.WaitUntilAvailable(By.XPath(Elements.Xpath[Reference.Login.UserId]),
                    $"The Office 365 sign in page did not return the expected result and the user '{username}' could not be signed in.");

                driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.UserId])).SendKeys(username.ToUnsecureString());
                driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.UserId])).SendKeys(Keys.Tab);
                driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.UserId])).SendKeys(Keys.Enter);
                Thread.Sleep(2000);

                //If expecting redirect then wait for redirect to trigger
                if (redirectAction != null)
                {
                    //Wait for redirect to occur.
                    Thread.Sleep(3000);

                    redirectAction?.Invoke(new LoginRedirectEventArgs(username, password, driver));

                    redirect = true;
                }
                else
                {
                    try
                    {
                        driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.LoginPassword])).SendKeys(password.ToUnsecureString());
                        driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.LoginPassword])).SendKeys(Keys.Tab);
                        driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.LoginPassword])).Submit();

                        if (driver.IsVisible(By.XPath(Elements.Xpath[Reference.Login.StaySignedIn])))
                        {
                            driver.ClickWhenAvailable(By.XPath(Elements.Xpath[Reference.Login.StaySignedIn]));
                        }

                        driver.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.Login.CrmMainPage])
                            , new TimeSpan(0, 0, 60),
                            e => { e.WaitForPageToLoad(); },
                            f => { throw new Exception("Login page failed."); });
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

            else
            {
                var protocol = uri.Scheme;
                driver.Navigate().GoToUrl(protocol + "://" + Uri.EscapeDataString(username.ToUnsecureString()) + ":" + Uri.EscapeDataString(password.ToUnsecureString()) + "@" + uri.Authority + uri.AbsolutePath);

            }

            return redirect ? LoginResult.Redirect : LoginResult.Success;
        }
    }
}