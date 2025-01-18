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
using Ionic.Zip;

namespace SeleniumGenericObject
{
    /*
    **************************
        
    Developed by Juraj Palusek

    **************************
        
    This is a source code for BluePrism 
    wrapper implementing Selenium Web Automation.

    Current solution works both for MS Chromium Edge and Google Chrome browsers.

    Code consists of multiple classes, which group together logical blocks.
    Apart from Constants, all other classes are leveraging class inheritance
    mostly because of code readability and protection levels.
    */

    class Drivers
    {
        public static RemoteWebDriver Driver { get; set; }
    }

    static class Constants
    { 
        public const string DriverPath = @"C:\IBOTS\Selenium\";

        public const string EdgeDriverName = "msedgedriver.exe";

        public const string EdgeDriverZipName = @"/edgedriver_win32.zip";

        public const string DriverBaseUrl = @"https://msedgedriver.azureedge.net/";

        static readonly string[] rootPaths = new string[]
        {
            @"C:\Program Files (x86)",
            @"C:\Program Files"
        };

        const string edgePath = @"Microsoft\Edge\Application\msedge.exe";

        // Returns a valid Edge application location
        public static string EdgeLocation
        {
            get
            {
                string combinedPath;

                foreach (string path in rootPaths)
	            {
                    combinedPath = Path.Combine(path, edgePath);
                    if (File.Exists(combinedPath))
                            return combinedPath;
	            }
                throw new FileNotFoundException("Unable to find Edge location.");
            }
        }
    }

    class BrowserManipulation : Locators
    {
        public static void StartEdge()
        {
            UpdateEdgeDriver();

            EdgeOptions options = new EdgeOptions();
            {
                options.UseChromium = true;
                options.AddUserProfilePreference("download.default_directory", @"C:\IBOTS");
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("safebrowsing.enabled", false);
                options.AddArgument("--test-type");
                options.AddArgument("start-maximized");
                options.AddArgument("--js-flags=--expose-gc");
                options.AddArgument("--enable-precise-memory-info");
                options.AddArgument("--disable-popup-blocking");
                options.AddArgument("--disable-default-apps");
                options.AddArgument("--enable-automation");
                options.AddArgument("disable-extensions");
                options.AddAdditionalCapability("useAutomationExtension", false);
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--no-sandbox");
                options.AddArgument("disable-infobars");
            }

            Driver = new EdgeDriver(Constants.DriverPath, options);
        }

        public static void StartChrome()
        {
            ChromeOptions options = new ChromeOptions();
            {
                options.AddArgument("--test-type");
                options.AddArgument("start-maximized");
                options.AddArgument("--js-flags=--expose-gc");
                options.AddArgument("--enable-precise-memory-info");
                options.AddArgument("--disable-popup-blocking");
                options.AddArgument("--disable-default-apps");
                options.AddArgument("--enable-automation");
                options.AddArgument("disable-extensions");
                options.AddAdditionalCapability("useAutomationExtension", false);
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--no-sandbox");
                options.AddArgument("disable-infobars");
                options.AddUserProfilePreference("download.default_directory", @"C:\IBOTS");
            }

            Driver = new ChromeDriver(Constants.DriverPath, options);
        }

        private static void UpdateEdgeDriver()
        {
            int DriverMajorRelease;

            string DriverDir = Path.Combine(Constants.DriverPath, Constants.EdgeDriverName);

            if (File.Exists(DriverDir))
            {
                DriverMajorRelease = FileVersionInfo.GetVersionInfo(DriverDir).ProductMajorPart;
            }
            else
            {
                Directory.CreateDirectory(Constants.DriverPath);

                DriverMajorRelease = 0;
                
            }
            
            int EdgeMajorRelease = FileVersionInfo.GetVersionInfo(Constants.EdgeLocation).ProductMajorPart;

            if (DriverMajorRelease != EdgeMajorRelease)
            {
                string EdgeVersion = FileVersionInfo.GetVersionInfo(Constants.EdgeLocation).FileVersion;

                string EdgeDriverUrl = Constants.DriverBaseUrl + EdgeVersion + Constants.EdgeDriverZipName;

                string PackageName = Path.GetFileName(new Uri(EdgeDriverUrl).LocalPath);
                
                WebClient client = new WebClient();

                client.DownloadFile(EdgeDriverUrl, Constants.DriverPath + PackageName);

                using (ZipFile driverPackage = ZipFile.Read(Constants.DriverPath + PackageName))
                {
                    foreach (ZipEntry entry in driverPackage)
                    {
                        if (entry.FileName == Constants.EdgeDriverName)
                        {
                            entry.Extract(Constants.DriverPath, ExtractExistingFileAction.OverwriteSilently);
                            break;
                        }
                    }
                }

                File.Delete(Constants.DriverPath + PackageName);
            }
        }

        public static string GetSessionID()
        {
            return Driver.SessionId.ToString();
            //var sessionIdProperty = typeof(RemoteWebDriver).GetProperty("sessionId", BindingFlags.Instance | BindingFlags.NonPublic); //sessionId
            //if (sessionIdProperty != null)
            //{
            //    SessionId sessionId = sessionIdProperty.GetValue(Driver, null) as SessionId;
            //    return sessionId.ToString();//((RemoteWebDriver)driver).Capabilities.GetCapability("webdriver.remote.sessionid").ToString();
            //}
            //else return null;
        }

        public static string GetExecutorURLFromDriver()
        {
            var executorField = typeof(RemoteWebDriver)
                .GetField("executor",
                          BindingFlags.NonPublic
                          | BindingFlags.Instance);

            var executor = executorField.GetValue(Driver);

            var internalExecutorField = executor.GetType()
                .GetField("internalExecutor",
                          BindingFlags.NonPublic
                          | BindingFlags.Instance);
            var internalExecutor = internalExecutorField.GetValue(executor);

            var remoteServerUriField = internalExecutor.GetType()
                .GetField("remoteServerUri",
                          BindingFlags.NonPublic
                          | BindingFlags.Instance);
            var remoteServerUri = remoteServerUriField.GetValue(internalExecutor) as Uri;

            return remoteServerUri.ToString();
        }

        public static void NavigateToWebPage(string pageUrl)
        {
            Driver.Navigate().GoToUrl(@pageUrl);
        }

        public static void GoBack()
        {
            Driver.Navigate().Back();
        }

        public static void GoForward()
        {
            Driver.Navigate().Forward();
        }

        public static void RefreshWebPage()
        {
            Driver.Navigate().Refresh();
        }

        public static void MaximizeWindow()
        {
            Driver.Manage().Window.Maximize();
        }

        public static void MinimizeWindow()
        {
            Driver.Manage().Window.Minimize();
        }

        public static void CloseBrowser()
        {
            Driver.Quit();
            Driver = null;
        }

        //Scroll to the specified element on the Web Page
        public static void ScrollToElement(string elementType, string element, int waitTime = 0)
        {
            Actions actions = new Actions(Driver);
            actions.MoveToElement(Locators.GetElementByLocator(elementType, element, waitTime)).Perform();
        }

        //Scroll to the top of the Web Page
        public static void ScrollToTop()
        {
            ((IJavaScriptExecutor)Driver).ExecuteScript("window.scrollTo(document.body.scrollHeight, 0)");
        }

        //Take screenshot of WebPage
        public static void TakeScreenshot(string path)
        {
            ITakesScreenshot screenshotDriver = (ITakesScreenshot)Driver;
            Screenshot screenshot = screenshotDriver.GetScreenshot();
            screenshot.SaveAsFile(@path);
        }

        //Execute script without any return value
        public static void ExecuteJavaScript(string script)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript(script);
        }

        //Execute script with a specified type of return value
        public static T ExecuteJavaScriptwithReturnVal<T>(string script)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
            return (T)js.ExecuteScript(script);
        }

        //Switch to JS alert popup and accept or dismiss it based on the bool input provided
        public static void HandleJavaScriptAlert(bool accept)
        {
            IAlert alert = Driver.SwitchTo().Alert();
            if (accept)
                alert.Accept();
            else
                alert.Dismiss();
        }

        //Switch to JS alert popup, send keys to input field and accept
        public static void SetJavaScriptAlertText(string alertText)
        {
            IAlert alert = Driver.SwitchTo().Alert();
            alert.SendKeys(alertText);
            alert.Accept();
        }

        //Switch to JS alert popup and get text
        public static string GetJavaScriptAlertText()
        {
            return Driver.SwitchTo().Alert().Text;
        }

        //Switch to First Tab or browser windows
        public static void SwitchToFirstTab()
        {
            ReadOnlyCollection<string> windowHandles = Driver.WindowHandles;
            Driver.SwitchTo().Window(windowHandles.First());
        }

        //Switch between tabs or browser windows
        public static void SwitchToLastTab()
        {
            ReadOnlyCollection<string> windowHandles = Driver.WindowHandles;
            Driver.SwitchTo().Window(windowHandles.Last());
        }

        //Switch to specified tab or browser window
        public static void SwitchToSpecificTab(int tabNumber)
        {
            ReadOnlyCollection<string> windowHandles = Driver.WindowHandles;
            Driver.SwitchTo().Window(windowHandles[tabNumber]);
        }

        //Add new browser tab and activate it by default
        public static void AddBrowserTab(bool activateTab)
        {
            ((IJavaScriptExecutor)Driver).ExecuteScript("window.open();");
            if (activateTab)
                Driver.SwitchTo().Window(Driver.WindowHandles.Last());
        }

        // Close the currently active tab
        public static void CloseCurrentTab()
        {
            Driver.Close();
        }

        //Switch between Frames
        public static void SwitchToFrame()
        {
            Driver.SwitchTo().DefaultContent();
        }

        public static void SwitchToFrame(int frameNumber)
        {
            Driver.SwitchTo().Frame(frameNumber);
        }

        public static void SwitchToFrame(string frameName)
        {
            Driver.SwitchTo().Frame(frameName);
        }

        public static void SwitchToFrame(string elementType, string element, int waitTime = 0)
        {
            Driver.SwitchTo().Frame(Locators.GetElementByLocator(elementType, element, waitTime));
        }

        //Wait until a page is fully loaded via JavaScript
        public static void WaitUntilPageLoaded(int waitTime)
        {
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(waitTime));
            wait.Until((x) =>
            {
                return ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete");
            });
        }
    }

    class Locators : Drivers
    {
        protected static string CleanUserInput(string elementType)
        {
            return elementType.ToLower().Trim();
        }

        protected static IWebElement GetElementByLocator(string elementType, string element, int waitTime)
        {
            switch (CleanUserInput(elementType))
            {
                case "id":
                    if (waitTime == 0)
                        return Driver.FindElement(By.Id(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "classname":
                    if (waitTime == 0)
                        return Driver.FindElement(By.ClassName(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "cssselector":
                    if (waitTime == 0)
                        return Driver.FindElement(By.CssSelector(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "linktext":
                    if (waitTime == 0)
                        return Driver.FindElement(By.LinkText(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "name":
                    if (waitTime == 0)
                        return Driver.FindElement(By.Name(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "partiallinktext":
                    if (waitTime == 0)
                        return Driver.FindElement(By.PartialLinkText(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "tagname":
                    if (waitTime == 0)
                        return Driver.FindElement(By.TagName(element));
                    return GetElementWithWait(elementType, element, waitTime);
                case "xpath":
                    if (waitTime == 0)
                        return Driver.FindElement(By.XPath(element));
                    return GetElementWithWait(elementType, element, waitTime);
                default:
                    throw new ArgumentException($"Provided argument {elementType} is invalid.");
            }
        }

        protected static IList<IWebElement> GetElementsByLocator(string elementType, string element)
        {
            switch (CleanUserInput(elementType))
            {
                case "id":
                    return Driver.FindElements(By.Id(element));
                case "classname":
                    return Driver.FindElements(By.ClassName(element));
                case "cssselector":
                    return Driver.FindElements(By.CssSelector(element));
                case "linktext":
                    return Driver.FindElements(By.LinkText(element));
                case "name":
                    return Driver.FindElements(By.Name(element));
                case "partiallinktext":
                    return Driver.FindElements(By.PartialLinkText(element));
                case "tagname":
                    return Driver.FindElements(By.TagName(element));
                case "xpath":
                    return Driver.FindElements(By.XPath(element));
                default:
                    throw new ArgumentException($"Provided argument {elementType} is invalid.");
            }
        }

        private static IWebElement GetElementWithWait(string elementType, string element, int waitTime)
        {
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(waitTime));

            switch (elementType)
            {
                case "id":
                    return wait.Until(x => x.FindElement(By.Id(element)));
                case "classname":
                    return wait.Until(x => x.FindElement(By.ClassName(element)));
                case "cssselector":
                    return wait.Until(x => x.FindElement(By.CssSelector(element)));
                case "linktext":
                    return wait.Until(x => x.FindElement(By.LinkText(element)));
                case "name":
                    return wait.Until(x => x.FindElement(By.Name(element)));
                case "partiallinktext":
                    return wait.Until(x => x.FindElement(By.PartialLinkText(element)));
                case "tagname":
                    return wait.Until(x => x.FindElement(By.TagName(element)));
                case "xpath":
                    return wait.Until(x => x.FindElement(By.XPath(element)));
                default:
                    throw new ArgumentException($"Provided argument {elementType} is invalid.");
            }
        }
    }

    class SetMethods : Locators
    {
        public static void SendKeys(string elementType, string element, string value, int waitTime = 0)
        {
            Locators.GetElementByLocator(elementType, element, waitTime).SendKeys(value);
        }
        public static void SendSpecialKeys(string elementType, string element, string specialKey, int waitTime = 0)
        {
            var controledElement = Locators.GetElementByLocator(elementType, element, waitTime);

            switch (CleanUserInput(specialKey))
            {
                case "enter":
                    controledElement.SendKeys(Keys.Enter);
                    break;
                case "escape":
                    controledElement.SendKeys(Keys.Escape);
                    break;
                case "tab":
                    controledElement.SendKeys(Keys.Tab);
                    break;
                case "pageup":
                    controledElement.SendKeys(Keys.PageUp);
                    break;
                case "pagedown":
                    controledElement.SendKeys(Keys.PageDown);
                    break;
                case "return":
                    controledElement.SendKeys(Keys.Return);
                    break;
                case "delete":
                    controledElement.SendKeys(Keys.Delete);
                    break;
                case "up":
                    controledElement.SendKeys(Keys.ArrowUp);
                    break;
                case "down":
                    controledElement.SendKeys(Keys.ArrowDown);
                    break;
                case "left":
                    controledElement.SendKeys(Keys.ArrowLeft);
                    break;
                case "right":
                    controledElement.SendKeys(Keys.ArrowRight);
                    break;
                case "control":
                    controledElement.SendKeys(Keys.LeftControl);
                    break;
                case "alt":
                    controledElement.SendKeys(Keys.LeftAlt);
                    break;
                case "shift":
                    controledElement.SendKeys(Keys.LeftShift);
                    break;
                default:
                    throw new ArgumentException($"Provided argument {specialKey} is invalid.");
            }
        }
        public static void SendSpecialKeys(string elementType, string element, string specialKey, string standardKey, int waitTime = 0)
        {
            var controledElement = Locators.GetElementByLocator(elementType, element, waitTime);

            switch (CleanUserInput(specialKey))
            {
                case "control":
                    controledElement.SendKeys(Keys.LeftControl + standardKey);
                    break;
                case "alt":
                    controledElement.SendKeys(Keys.LeftAlt + standardKey);
                    break;
                case "shift":
                    controledElement.SendKeys(Keys.LeftShift + standardKey);
                    break;
                default:
                    throw new ArgumentException($"Provided argument {specialKey} is invalid.");
            }
        }
        public static void LeftClick(string elementType, string element, int waitTime = 0)
        {
            Locators.GetElementByLocator(elementType, element, waitTime).Click();
        }
        public static void RightClick(string elementType, string element, int waitTime = 0)
        {
            Actions actions = new Actions(Driver);
            actions.ContextClick(Locators.GetElementByLocator(elementType, element, waitTime)).Perform();
        }
        public static void DoubleClick(string elementType, string element, int waitTime = 0)
        {
            Actions actions = new Actions(Driver);
            actions.DoubleClick(Locators.GetElementByLocator(elementType, element, waitTime)).Perform();
        }
        public static void MouseHover(string elementType, string element, int waitTime = 0)
        {
            Actions actions = new Actions(Driver);
            actions.MoveToElement(Locators.GetElementByLocator(elementType, element, waitTime)).Perform();
        }
        public static void Clear(string elementType, string element, int waitTime = 0)
        {
            Locators.GetElementByLocator(elementType, element, waitTime).Clear();
        }
        public static void Submit(string elementType, string element, int waitTime = 0)
        {
            Locators.GetElementByLocator(elementType, element, waitTime).Submit();
        }
        public static void Select(string elementType, string element, string value, string selectOption, int waitTime = 0)
        {
            switch (CleanUserInput(selectOption))
            {
                case "text":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).SelectByText(value);
                    break;
                case "index":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).SelectByIndex(int.Parse(value));
                    break;
                case "value":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).SelectByValue(value);
                    break;
                default:
                    throw new ArgumentException($"Provided argument {selectOption} is invalid.");
            }
        }
        public static void Deselect(string elementType, string element, string value, string deselectOption, int waitTime = 0)
        {
            switch (CleanUserInput(deselectOption))
            {
                case "text":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).DeselectByText(value);
                    break;
                case "index":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).DeselectByIndex(int.Parse(value));
                    break;
                case "value":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).DeselectByValue(value);
                    break;
                case "all":
                    new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).DeselectAll();
                    break;
                default:
                    throw new ArgumentException($"Provided argument {deselectOption} is invalid.");
            }
        }
        //Drag and Drop to offset
        public static void DragAndDrop(string sourceElementType, string sourceElement, int offsetX, int offsetY, int waitTime)
        {
            Actions actions = new Actions(Driver);
            actions.DragAndDropToOffset(Locators.GetElementByLocator(sourceElementType, sourceElement, waitTime), offsetX, offsetY).Perform();
        }
        //Drag and Drop to element
        public static void DragAndDrop(string sourceElementType, string sourceElement, string targetElementType, string targetElement, int waitTime)
        {
            Actions actions = new Actions(Driver);
            actions.DragAndDrop(Locators.GetElementByLocator(sourceElementType, sourceElement, waitTime), Locators.GetElementByLocator(targetElementType, targetElement, waitTime)).Perform();
        }
    }

    class GetMethods : Locators
    {
        public static string GetInnerText(string elementType, string element, int waitTime = 0)
        {
            return Locators.GetElementByLocator(elementType, element, waitTime).Text;
        }
        public static string GetTextFromDDL(string elementType, string element, int waitTime = 0)
        {
            return new SelectElement(Locators.GetElementByLocator(elementType, element, waitTime)).AllSelectedOptions.SingleOrDefault().Text;
        }
        public static string GetAttribute(string elementType, string element, string attribute, int waitTime = 0)
        {
            return Locators.GetElementByLocator(elementType, element, waitTime).GetAttribute(attribute);
        }
        public static bool IsDisplayed(string elementType, string element, int waitTime = 0)
        {
            return Locators.GetElementByLocator(elementType, element, waitTime).Displayed;
        }
        public static bool IsEnabled(string elementType, string element, int waitTime = 0)
        {
            return Locators.GetElementByLocator(elementType, element, waitTime).Enabled;
        }
        public static bool IsSelected(string elementType, string element, int waitTime = 0)
        {
            return Locators.GetElementByLocator(elementType, element, waitTime).Selected;
        }
        public static int GetElementCount(string elementType, string element)
        {
            return Locators.GetElementsByLocator(elementType, element).Count();
        }
        //IsMultipleSelect

        // Get the title of the page, current URL and page HTML source
        public static string[] GetPageInfo()
        {
            return new string[]
            {
                Driver.Title,
                Driver.Url,
                Driver.PageSource
            };
        }
    }
}
