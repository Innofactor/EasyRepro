// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;

namespace Microsoft.Dynamics365.UIAutomation.Api
{
    /// <summary>
    /// Xrm Guided Help Page
    /// </summary>
    public class XrmGuidedHelpPage
        : XrmPage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XrmGuidedHelpPage"/> class.
        /// </summary>
        /// <param name="browser">The browser.</param>
        public XrmGuidedHelpPage(InteractiveBrowser browser)
            : base(browser)
        {
            SwitchToDefault();
        }

        public bool IsEnabled
        {
            get
            {
                bool isGuidedHelpEnabled = false;
                string getIsEnabledGuideScript = string.Empty;
                if (LoginPage.Online)
                {
                    getIsEnabledGuideScript = "return Xrm.Internal.isFeatureEnabled('FCB.GuidedHelp') && Xrm.Internal.isGuidedHelpEnabledForUser();";
                }
                else
                {
                    getIsEnabledGuideScript = "return Xrm.Internal.isGuidedHelpEnabledForUser();";
                }
                bool.TryParse(this.Browser.Driver.ExecuteScript(getIsEnabledGuideScript).ToString(), out isGuidedHelpEnabled);
                return isGuidedHelpEnabled;
            }
        }

        /// <summary>
        /// Closes the Guided Help
        /// </summary>
        /// <example>xrmBrowser.GuidedHelp.CloseGuidedHelp();</example>
        public BrowserCommandResult<bool> CloseGuidedHelp()
        {
            return this.Execute(GetOptions("Close Guided Help"), driver =>
            {
                bool returnValue = false;

                if (IsEnabled)
                {
                    if (LoginPage.Online)
                    {
                        driver.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.GuidedHelp.MarsOverlay]), new TimeSpan(0, 0, 15), d =>
                        {
                            var allMarsElements = driver
                                .FindElement(By.XPath(Elements.Xpath[Reference.GuidedHelp.MarsOverlay]))
                                .FindElements(By.XPath(".//*"));

                            foreach (var element in allMarsElements)
                            {
                                var buttonId = driver.ExecuteScript("return arguments[0].id;", element).ToString();

                                if (buttonId.Equals(Elements.ElementId[Reference.GuidedHelp.Close], StringComparison.InvariantCultureIgnoreCase))
                                {
                                    driver.WaitUntilClickable(By.Id(buttonId), new TimeSpan(0, 0, 5));

                                    element.Click();
                                }
                            }

                            returnValue = true;
                        });
                    }
                    else
                    {
                        driver.WaitUntilVisible(By.Id(Reference.GuidedHelp.MarsOverlay), new TimeSpan(0, 0, 15), d =>
                        {
                            driver.SwitchTo().Frame(Reference.GuidedHelp.GuideIFrame);
                            driver.WaitUntilClickable(By.Id(Reference.GuidedHelp.ButtonClose), new TimeSpan(0, 0, 5));
                            var e = driver.FindElement(By.Id(Reference.GuidedHelp.ButtonClose));
                            e.Click(true);
                            driver.SwitchTo().DefaultContent();

                            returnValue = true;
                        });
                    }

                }

                return returnValue;
            });
        }
    }
}