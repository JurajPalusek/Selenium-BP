# Selenium-BP

Selenium-BP is a C# based source code for BluePrism wrapper to utilize Selenium Web Automation Framework. Current solution works both for MS Chromium Edge and Google Chrome browsers.

# Usage

Blue Prism object structure allows to have all the source code on "global page" and every action then consists of Method call with input / output parameters. In order to easily change / update the code, the solution is divided into three .cs files:

- [GlobalCode.cs](https://github.com/JurajPalusek/Selenium-BP/blob/main/GlobalCode.cs)
- [GlobalCodeSpecific.cs](https://github.com/JurajPalusek/Selenium-BP/blob/main/GlobalCodeSpecific.cs)
- [Program.cs](https://github.com/JurajPalusek/Selenium-BP/blob/main/Program.cs)

## GlobalCode.cs

Includes the main bulk of the solution, with all the automation methods, including browser start and auto update of Selenium WebDriver (currently implemented for MS Chromium Edge only).

## GlobalCodeSpecific.cs

Due to annoying BluePrism instance limiation, it is impossible to "transfer" object instance from one BP object to another. It is therefore impossible by default to create a specific Selenium action if necesarry without modifying the original Generic object. This solution is an attempt to circumvent this limitation by attaching to running browser, which however must have been created by Selenium in the first place. 

It has only one method AttachToBrowser(), which requires two input strings executorURL and sessionID, which are extracted by the Selenium Generic Object methods GetSessionID() and GetExecutorURLFromDriver().

## Program.cs

Created just for testing purposes inside IDE, it has no real ussage on the BluePrism object level.

# Testing

Basic functionality of the solution can by tested in Program.cs in Main() for example:

```cs
static void Main()
{
    Console.WriteLine("Implementation of Selenium Web Automation for BP Object");

    BrowserManipulation.StartEdge();

    BrowserManipulation.NavigateToWebPage("https://www.github.com/");

    SetMethods.SendKeys("name", "q", "selenium");
}
```

# Contributing

This code is for demonstration purposes only to showcase the use of Selenium framework inside Blue Prism object. No pull or clone requests are possible.
