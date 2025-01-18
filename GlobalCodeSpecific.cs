using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using Microsoft.Edge.SeleniumTools;

namespace SeleniumSpecificObject
{
    /*
    **************************

    Developed by Juraj Palusek

    **************************
    
    Due to annoying BluePrism instance limiation, it is impossible to
    "transfer" object instance from one BP object to another.

    It is therefore impossible by default to create a specific Selenium 
    action if necesarry without modifying the original Generic object.

    This solution is an attempty to circumvent this limitation by 
    attaching to running browser, which however must have been created
    by Selenium in the first place.

    It has only one method AttachToBrowser(), which requires two input strings
    executorURL and sessionID, which are extracted by the Selenium Generic 
    Object methods GetSessionID() and GetExecutorURLFromDriver().

    Unused usings are not a mistake. Just a preparation of all necesarry 
    Selenium-related references to start coding.
    */

    class Drivers
    {
        public static RemoteWebDriver Driver { get; set; }
    }

    class SeleniumSpecificActions : Drivers
    {
        public static void AttachToBrowser(string executorURL, string sessionID)
        {
            Driver = new ReuseRemoteWebDriver(new Uri(executorURL), sessionID);
        }
    }

    class ReuseRemoteWebDriver : RemoteWebDriver
    {
        private string _sessionId;

        public ReuseRemoteWebDriver(Uri remoteAddress, string sessionId)
            : base(remoteAddress, new DesiredCapabilities())
        {
            this._sessionId = sessionId;
            var sessionIdBase = this.GetType()
                .BaseType
                .GetField("sessionId",
                          BindingFlags.Instance |
                          BindingFlags.NonPublic);
            sessionIdBase.SetValue(this, new SessionId(sessionId));
        }

        protected override Response
            Execute(string driverCommandToExecute, Dictionary<string, object> parameters)
        {
            if (driverCommandToExecute == DriverCommand.NewSession)
            {
                var response = new Response
                {
                    Status = WebDriverResult.Success,
                    SessionId = this._sessionId,
                    Value = new Dictionary<string, object>()
                };
                return response;
            }
            var respBase = base.Execute(driverCommandToExecute, parameters);
            return respBase;
        }
    }
}
