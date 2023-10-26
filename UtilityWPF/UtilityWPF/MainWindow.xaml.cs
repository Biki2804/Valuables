using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Microsoft.SharePoint.Client;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
//using Condition = System.Windows.Automation.Condition;
using System.Security;
using System.Security.Cryptography;
using System.Threading;
using System.Web;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Shapes;
using Exception = System.Exception;
using List = Microsoft.SharePoint.Client.List;

namespace UtilityWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        System.Timers.Timer popUpTimer = new System.Timers.Timer();

        public MainWindow()
        {
            InitializeComponent();


            //DownloadAllDocumentsfromLibrary();

            // HandleAuthenticationDialogForIE("965539");


            popUpTimer.Elapsed += PopUpTimer_Elapsed;
            popUpTimer.Interval = 2000;


        }
        bool checkPopUpOpen = false;

        private void PopUpTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {

            //popUpTimer.Enabled = false;
            //Console.WriteLine(driver.Title.ToString());
            //0x2012E
            string password = "280491";
            //Thread.Sleep(1500);
            try
            {
                if (checkPopUpOpen)
                {
                    var MainWindow = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                        new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                            new PropertyCondition(AutomationElement.ClassNameProperty, "PasswordBox"),
                            new PropertyCondition(AutomationElement.NameProperty, "PIN")));
                    Thread.Sleep(200);
                    popUpTimer.Enabled = false;
                    if (MainWindow != null)
                    {
                        ValuePattern userNamePattern = (ValuePattern)MainWindow.GetCurrentPattern(ValuePattern.Pattern);
                        userNamePattern.SetValue(password);

                        Thread.Sleep(200);
                        var MainWindow2 = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                                 new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                                     new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                                     new PropertyCondition(AutomationElement.NameProperty, "OK")));

                        Thread.Sleep(500);

                        InvokePattern buttonPattern = (InvokePattern)MainWindow2.GetCurrentPattern(InvokePattern.Pattern);
                        buttonPattern.Invoke();
                        popUpTimer.Enabled = false;
                    }
                    else
                    {

                        popUpTimer.Enabled = true;
                        Thread.Sleep(1000);
                    }

                    var errorWindow = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                         new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                             new PropertyCondition(AutomationElement.ClassNameProperty, "TextBlock"),
                             new PropertyCondition(AutomationElement.NameProperty, "An incorrect PIN was presented to the smart card")));
                    if (errorWindow != null)
                    {
                        popUpTimer.Enabled = true;
                        Thread.Sleep(2000);
                    }
                    else { popUpTimer.Enabled = false; }

                }

            }
            catch (Exception ex)
            {


            }


        }

        IWebDriver driver;


        public void GetFolders(MAPIFolder folder)
        {
            if (folder.Folders.Count == 0)
            {
                if (folder.Name == "HPS3 Tickets")
                {
                    //Console.WriteLine(m.FullFolderPath);
                    mailsFromThisFolder = folder;
                }
            }
            else
            {
                foreach (MAPIFolder subFolder in folder.Folders)
                {
                    GetFolders(subFolder);
                }
            }
        }

        Microsoft.Office.Interop.Outlook.MAPIFolder mailsFromThisFolder;

        string HTMLBODY = string.Empty;
        string url = string.Empty;
        string formattedDate = string.Empty;
        string titleVal= string.Empty;
        string priorityVal = string.Empty;

        string dateString = string.Empty;

        string userName = string.Empty;
        bool checkSC3Mail = false;

        string shortTime = string.Empty;
        internal string ReadOutlook()
        {
            string URL = string.Empty;
            TicketVal = string.Empty;
            titleVal = string.Empty;
            priorityVal = string.Empty;
            dateString = string.Empty;
            userName = string.Empty;
            HTMLBODY = string.Empty;
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = null;
                Microsoft.Office.Interop.Outlook._NameSpace ns = null;
                //Microsoft.Office.Interop.Outlook.MailItem item = null;
                //Microsoft.Office.Interop.Outlook.MAPIFolder rootFolder = null;
                //Microsoft.Office.Interop.Outlook.MAPIFolder subFolder = null;
                string subject  = string.Empty;
                app = new Microsoft.Office.Interop.Outlook.Application();

                ns = app.GetNamespace("MAPI");


                MAPIFolder mainFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

                foreach (MAPIFolder folder in mainFolder.Folders)
                {
                    GetFolders(folder);
                }

                Microsoft.Office.Interop.Outlook.Items items = mailsFromThisFolder.Items;
                for (int counter = 1; counter <= items.Count; counter++)
                {

                    Console.Write(items.Count + " " + counter);
                    dynamic item;
                    item = items[counter];
                    if (item.Unread == true)
                    {
                        checkSC3Mail = true;
                        string counterVal =  counter.ToString();
                        subject = item.Subject;
                        string longDate = item.SentOn.ToString();
                        string senderName = item.ReceivedByName.ToString();
                        shortTime = item.SentOn.ToShortTimeString();
                        shortTime = shortTime.Trim().Substring(0, 2);
                        string[] splitString = senderName.Split(new[] { "(" }, StringSplitOptions.RemoveEmptyEntries);
                        string[] Namestring = splitString[0].Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                        
                        string userName =Namestring[1].Trim().ToString();
                       
                        dateString = longDate;
                        DateTime parsedDateTime = DateTime.Parse(longDate);

                        dateString = parsedDateTime.ToString("ddd MM/dd/yyyy HH:mm:ss tt");

                        HTMLBODY = item.HTMLBody;
                        item.Unread = false;
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(HTMLBODY))
                {
                    var doc = new HtmlDocument();
                    var encoded = HttpUtility.HtmlDecode(HTMLBODY);
                    doc.LoadHtml(encoded);
                    var nodes = doc.DocumentNode.SelectNodes("//a[@href]");
                    foreach (var node in nodes)
                    {
                        url = node.Attributes["href"].Value;
                        URL = url.ToString();

                        break;
                    }
                    var tdNodes = doc.DocumentNode.SelectNodes("//td");
                    foreach (var node in tdNodes)
                    {
                        if(!string.IsNullOrEmpty(node.InnerText) && node.InnerText.Equals("Title"))
                        {
                             titleVal = node.NextSibling.NextSibling.InnerText;
                        }

                        if (!string.IsNullOrEmpty(node.InnerText) && (node.InnerText.Equals("Priority") || node.InnerText.Equals("Complexity")))
                        {
                            priorityVal = node.NextSibling.NextSibling.InnerText;
                        }
                        if(!string.IsNullOrEmpty(priorityVal))
                        {
                            priorityVal = priorityVal.Substring(1);
                        }
                        
                    }


                }

                Marshal.ReleaseComObject(ns);
                Marshal.ReleaseComObject(app);


                //strm.Close();
            }

            catch (System.Exception ex)
            {

                
            }
            return URL;
        }

        string statusVal = string.Empty;
        string nameField = string.Empty;
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            string testURL = string.Empty;
            ReadOutlook();

            if (checkSC3Mail)
            {
                checkSC3Mail = false;
                TicketVal = url.Substring(url.Length - 10);
                testURL = url;

                var options = new EdgeOptions();
                options.BinaryLocation = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";
                driver = new EdgeDriver(@"C:\Bikram\edgedriver_win64", options);
                driver.Manage().Window.Maximize();
                popUpTimer.Enabled = true;
                checkPopUpOpen = true;

                driver.Navigate().GoToUrl(testURL);

                Thread.Sleep(6000);
                string checkKeyword = string.Empty;
                string[] WorkStatusValue = ClickOnIncidentElement(TicketVal);
                statusVal = WorkStatusValue[1].ToString();
                string[] nameArr = WorkStatusValue[4].ToString().Split(',');
                if(nameArr.Length > 0)
                {
                    nameField = nameArr[1];
                }
                
                
                if (WorkStatusValue != null && (WorkStatusValue[1].ToLower().Equals("Working".ToLower()) 
                    || WorkStatusValue[1].ToLower().Equals("Updated".ToLower()) 
                    || WorkStatusValue[1].ToLower().Equals("Open".ToLower())
                    || WorkStatusValue[1].ToLower().Equals("Transferred".ToLower())) 
                    && (WorkStatusValue[2].ToLower().Equals("K2NG VWITS Support SKODA".ToLower()) || 
                    WorkStatusValue[2].ToLower().Equals("RetailCarConfig  VWITS Support SKODA".ToLower())))
                {
                    ClickOnTakOverBtn();
                    string assgnGrp = string.Empty;
                    if (WorkStatusValue[0] != null && (WorkStatusValue[0].ToLower().Contains("Milkyway".ToLower()) ||
                                                        WorkStatusValue[0].ToLower().Contains("GTM".ToLower()) ||
                                                        WorkStatusValue[0].ToLower().Contains("GA4".ToLower()) ||
                                                        WorkStatusValue[0].ToLower().Contains("PH INTEGRATION".ToLower())||
                                                        WorkStatusValue[0].ToLower().Contains("Rule CMS".ToLower())))
                    {
                        if(WorkStatusValue[0].ToLower().Contains("GTM".ToLower())||
                            WorkStatusValue[0].ToLower().Contains("GA4".ToLower()))
                        {
                            assgnGrp = "SDRIVE Support SKODA";
                            if(WorkStatusValue[0].ToLower().Contains("GTM".ToLower()))
                            {
                                checkKeyword = "GTM";
                            }
                            else
                            {
                                checkKeyword = "GA4";
                            }
                        }
                        else if (WorkStatusValue[0].ToLower().Contains("Milkyway".ToLower()))
                        {
                            assgnGrp = "DOTNET CORE Advanced Support SKODA";
                            checkKeyword = "Milkyway";
                        }
                        else
                        {
                            assgnGrp = "K2NG Advanced Support SKODA";
                            if (WorkStatusValue[0].ToLower().Contains("PH INTEGRATION".ToLower()))
                            {
                                checkKeyword = "PH INTEGRATION";
                            }
                            else
                            {
                                checkKeyword = "Rule CMS";
                            }
                        }

                        ChangeAssignmentGroup(TicketVal, assgnGrp);

                        ClickOnActivityTab(TicketVal);

                        

                        SelectDropDownText(TicketVal, checkKeyword);

                        //ClickOnSaveExitBtn();

                    }
                    else if(WorkStatusValue[0] != null && WorkStatusValue[0].ToLower().Contains("StarGate".ToLower()))
                    {

                        ClickOnMoreBtn();
                        ClickOnWaitOnUser();

                        //ClickOnActivityTab(TicketVal);

                        checkKeyword = "StarGate";
                        EnterWaitOnUserText(checkKeyword);
                        

                        //ClickOnSaveExitBtn();
                    }
                    else
                    {

                    }

                    // ClickOnActivityTab();

                    //To update file and upload into sharepoint folder
                    //DownloadAllDocumentsfromLibrary();
                }
            }
            else
            {
                testURL = "https://sc3.vwgroup.com/sc3/index.do?lang=en";

                var options = new EdgeOptions();
                options.BinaryLocation = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";
                driver = new EdgeDriver(@"C:\Bikram\edgedriver_win64", options);
                driver.Manage().Window.Maximize();
                popUpTimer.Enabled = true;
                checkPopUpOpen = true;

                driver.Navigate().GoToUrl(testURL);

                Thread.Sleep(2000);

                ClickIncident();
                ClickOnIncidentWorkingTicket();

                ClickRequest();
                ClickOnIncidentWorkingTicket();

                //DownloadAllDocumentsfromLibrary();
            }
            
            driver.Close();
            driver.Quit();


            checkPopUpOpen = false;
            
        }


        internal void SelectDropDownText(string tktVal, string Keyword)
        {
            try
            {
                IWebElement frame = driver.FindElement(By.XPath(".//iframe[@title='" + tktVal + "']"));

                driver.SwitchTo().Frame(frame);

                IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                
                js.ExecuteScript("var divVal = document.getElementsByTagName(\"div\");let innerVal;if(divVal){for(var i = 0; i < divVal.length; i++)" +
                    "{if(divVal[i] && (divVal[i].innerText.indexOf(\"Analysis/Research\") >-1 || divVal[i].innerText.indexOf(\"Activity\") >-1))" +
                    "{innerVal = divVal[i].innerText;}}}if(innerVal && (innerVal.indexOf(\"Analysis/Research\") >-1 || innerVal.indexOf(\"Activity\") >-1))" +
                    "{var inp = document.getElementsByTagName(\"input\");if(inp) " +
                    "{for(var j = 0; j < inp.length; j++){if(inp[j] && inp[j].id === \"X163\")" +
                    "{if(innerVal.indexOf(\"Analysis/Research\") >-1)" +
                    "{inp[j].value = \"Analysis/Research\";}" +
                    "else{inp[j].value = \"Activity\";}break;}}}}");


                Thread.Sleep(1000);


                //driver.FindElement(By.XPath(".//label[@id = 'X165_Label']")).Click();
                string checkText = string.Empty;
                checkText = Keyword;
                OpenQA.Selenium.Interactions.Actions act = new OpenQA.Selenium.Interactions.Actions(driver);
                act.MoveToElement(driver.FindElement(By.XPath(".//div[@id = 'X165View']"))).DoubleClick().Build().Perform();
                switch (checkText)
                {
                    case "Milkyway":
                        string milkywayText = string.Empty;
                        milkywayText = "Hello Dear colleagues," + Environment.NewLine +
                            "This ticket is related to Milkyway so transferring to your group." + Environment.NewLine +
                            "Could you please look into it." + Environment.NewLine + Environment.NewLine +
                            "Thank you," + Environment.NewLine +
                            "AMS Team";
                        act.SendKeys(milkywayText).Build().Perform();
                        break;
                    case "GA4":
                        string GA4Text = string.Empty;
                        GA4Text = "Hello dear colleagues, " + Environment.NewLine +
                            "This ticket is related to Access to GA4, " + Environment.NewLine +
                            "so assigning it to your group." + Environment.NewLine +
                            "Could you please look into this." + Environment.NewLine +
                            "Have a nice day!" + Environment.NewLine + Environment.NewLine +
                            "Thank You," + Environment.NewLine +
                            "AMS Team";
                        act.SendKeys(GA4Text).Build().Perform();
                        break;
                    case "GTM":
                        string ga4Text = string.Empty;
                        ga4Text = "Hello dear colleagues," + Environment.NewLine +
                            "This ticket is regarding GTM and Tagging, so transferring this ticket to your group. " + Environment.NewLine +
                            "Could you please look into it." + Environment.NewLine + Environment.NewLine +
                            "Thank you," + Environment.NewLine +
                            "AMS Team";
                        act.SendKeys(ga4Text).Build().Perform();
                        break;
                    case "PH INTEGRATION":
                        string phIntegrationText = string.Empty;
                        phIntegrationText = "Hello Dear Colleagues," +
                            "This ticket is related to monitoring issue." + Environment.NewLine +
                            "Could you please look into this." + Environment.NewLine +
                            "\r\nHave a nice day!" + Environment.NewLine + Environment.NewLine +
                            "Thank you," + Environment.NewLine +
                            "\r\nAMS Team";
                        act.SendKeys(phIntegrationText).Build().Perform();
                        break;
                    case "Rule CMS":
                        string ruleCMSText = string.Empty;
                        ruleCMSText = "Hello Dear Colleagues," +
                            "This ticket is related to monitoring issue." + Environment.NewLine +
                            "Could you please look into this."+ Environment.NewLine +
                            "\r\nHave a nice day!" + Environment.NewLine + Environment.NewLine +
                            "Thank you," + Environment.NewLine +
                            "\r\nAMS Team";
                        act.SendKeys(ruleCMSText).Build().Perform();
                        break;

                    default:
                        break;

                }

                
                

                //var element1 = driver.FindElement(By.Id("X165View"));
                //element1.SendKeys("test");

                driver.SwitchTo().DefaultContent();
                
            }
            catch (Exception ex)
            {

                
            }
        }

        internal void EnterWaitOnUserText(string Keyword)
        {
            try
            {
                IWebElement frame = driver.FindElement(By.XPath(".//iframe[@title='Wizard: Reason']"));

                driver.SwitchTo().Frame(frame);

                
                //driver.FindElement(By.XPath(".//label[@id = 'X165_Label']")).Click();
                string checkText = string.Empty;
                checkText = Keyword;
                OpenQA.Selenium.Interactions.Actions act = new OpenQA.Selenium.Interactions.Actions(driver);
                act.MoveToElement(driver.FindElement(By.XPath(".//div[@id = 'X8View']"))).Click().Build().Perform();

                switch (checkText)
                {
                    case "StarGate":
                        string starGateText = string.Empty;
                        starGateText = "Hi" + nameField + " ," + Environment.NewLine +
                            "As the Star gate Application is still not in production." + Environment.NewLine +
                            "Therefore, we are closing this ticket." + Environment.NewLine + Environment.NewLine +
                            "Thanks," + Environment.NewLine +
                            "AMS Team";
                        act.SendKeys(starGateText).Build().Perform();
                        Thread.Sleep(1000);
                        IWebElement btn = driver.FindElement(By.Id("X14Btn"));
                        btn.Click();
                        break;
                }




                //var element1 = driver.FindElement(By.Id("X165View"));
                //element1.SendKeys("test");

                driver.SwitchTo().DefaultContent();

            }
            catch (Exception ex)
            {


            }
        }

        string TicketVal = string.Empty;
        string statusValue = string.Empty;

        string[] assignmentValue = new string[10];
        internal string[] ClickOnIncidentElement(string tcktVal)
        {
            assignmentValue = new string[10];
            try
            {
                Thread.Sleep(3000);
                TicketVal = tcktVal;

                IWebElement frame = driver.FindElement(By.XPath(".//iframe[@title='" + TicketVal + "']"));

                driver.SwitchTo().Frame(frame);


                var html = driver.FindElements(By.TagName("div"));
                int count = html.Count();

                bool statusCheck = true;

                if (html != null && html.Count >= 1)
                {
                    for (int i = 0; i < html.Count; i++)
                    {
                        count--;

                        var body = html[i].FindElements(By.TagName("form"));
                        if (body != null && body.Count >= 1)
                        {
                            for (int j = 0; j < body.Count; j++)
                            {
                                var iframe = body[j].FindElements(By.TagName("div"));
                                if (iframe != null && iframe.Count >= 1)
                                {

                                    for (int k = 0; k < iframe.Count; k++)
                                    {
                                        if (statusCheck)
                                        {
                                            var td = iframe[k].FindElements(By.TagName("input"));
                                            if (td != null && td.Count >= 1)
                                            {

                                                for (int l = 0; l < td.Count; l++)
                                                {
                                                    if (td[l] != null && td[l].GetAttribute("id").Equals("X16"))
                                                    {

                                                        assignmentValue[0] = td[l].GetAttribute("value");
                                                       
                                                    }


                                                    if (td[l] != null && td[l].GetAttribute("id").Equals("X18"))
                                                    {

                                                        assignmentValue[1] = td[l].GetAttribute("value");
                                                        statusCheck = false;

                                                        break;
                                                    }
                                                }

                                                if (!statusCheck)
                                                { break; }
                                            }


                                        }
                                    }
                                    if (!statusCheck)
                                    { break; }
                                }


                            }

                            if (!statusCheck)
                            { break; }
                        }

                    }
                }
                statusCheck = true;

                if (html != null && html.Count >= 1)
                {
                    for (int i = 0; i < html.Count; i++)
                    {
                        count--;

                        var body = html[i].FindElements(By.TagName("form"));
                        if (body != null && body.Count >= 1)
                        {
                            for (int j = 0; j < body.Count; j++)
                            {
                                var iframe = body[j].FindElements(By.TagName("div"));
                                if (iframe != null && iframe.Count >= 1)
                                {

                                    for (int k = 0; k < iframe.Count; k++)
                                    {
                                        if (statusCheck)
                                        {
                                            var td = iframe[k].FindElements(By.TagName("input"));
                                            if (td != null && td.Count >= 1)
                                            {

                                                for (int l = 0; l < td.Count; l++)
                                                {
                                                    if (td[l] != null && (td[l].GetAttribute("id").Equals("X77") || td[l].GetAttribute("id").Equals("X58")))
                                                    {

                                                        assignmentValue[2] = td[l].GetAttribute("value");

                                                    }

                                                    if (td[l] != null && (td[l].GetAttribute("id").Equals("X33") || td[l].GetAttribute("id").Equals("X50")))
                                                    {

                                                        assignmentValue[4] = td[l].GetAttribute("value");

                                                    }

                                                    if (td[l] != null && (td[l].GetAttribute("id").Equals("X130") || td[l].GetAttribute("id").Equals("X88")))
                                                    {

                                                        assignmentValue[3] = td[l].GetAttribute("value");
                                                        statusCheck = false;

                                                        break;
                                                    }
                                                }

                                                if (!statusCheck)
                                                {
                                                    break;
                                                }
                                            }

                                        }
                                    }

                                    if (!statusCheck)
                                    {
                                        break;
                                    }
                                }

                            }
                            if (!statusCheck)
                            {
                                break;
                            }
                        }

                    }
                }


                driver.SwitchTo().DefaultContent();

                //IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                //IWebElement ClkPaiementByCard = driver.FindElement(By.XPath("//iframe[@title='IR Queue']"));
                //js.ExecuteScript("arguments[0].click();", ClkPaiementByCard);

                //driver.FindElement(By.XPath("//iframe[@title='IR Queue']")).SendKeys(Keys.F5);


                //IJavaScriptExecutor js1 = driver as IJavaScriptExecutor;
                //js1.ExecuteScript("var targetTDs = document.querySelectorAll('table[class=\"x-grid3-row-table\"]');" +
                //    "for (var i = 0; i < targetTDs.length; i++)" +
                //    "{var td = targetTDs[i];" +
                //    "if(td.innerText.indexOf('Working') > -1)" +
                //    "{var innertd = td.getElementsByTagName('a');" +
                //    "for (var j = 0; j < innertd.length; j++)" +
                //    "{if(innertd[j].innerText.startsWith(\"IR\"))" +
                //    "{innertd[j].click();}}}}");


            }
            catch (System.Exception ex)
            {

                throw ex;
            }
            return assignmentValue;
        }


        internal string[] ChangeAssignmentGroup(string tcktVal, string asgngrp)
        {
            assignmentValue = new string[10];
            try
            {
                //TicketVal = tcktVal;

                IWebElement frame = driver.FindElement(By.XPath(".//iframe[@title='" + tcktVal + "']"));

                driver.SwitchTo().Frame(frame);

                if (tcktVal.StartsWith("IR"))
                {
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    Thread.Sleep(2000);
                    js.ExecuteScript("var inp = document.getElementsByTagName(\"input\");if(inp){for(var i = 0; i<inp.length; i++){if(inp[i] && inp[i].id === \"X77\"){inp[i].value = \" \";}}}");
                    Thread.Sleep(1000);


                    IJavaScriptExecutor js1 = driver as IJavaScriptExecutor;
                    Thread.Sleep(2000);
                    js1.ExecuteScript("var inp = document.getElementsByTagName(\"input\");if(inp){for(var i = 0; i<inp.length; i++){if(inp[i] && inp[i].id === \"X77\"){inp[i].value = '"+asgngrp+"';}}}");
                    Thread.Sleep(1000);
                }
                else if (tcktVal.StartsWith("RR"))
                {
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    Thread.Sleep(2000);
                    js.ExecuteScript("var inp = document.getElementsByTagName(\"input\");if(inp){for(var i = 0; i<inp.length; i++){if(inp[i] && inp[i].id === \"X58\"){inp[i].value = \" \";}}}");
                    Thread.Sleep(1000);


                    IJavaScriptExecutor js1 = driver as IJavaScriptExecutor;
                    Thread.Sleep(2000);
                    js1.ExecuteScript("var inp = document.getElementsByTagName(\"input\");if(inp){for(var i = 0; i<inp.length; i++){if(inp[i] && inp[i].id === \"X58\"){inp[i].value = '"+asgngrp+"';}}}");
                    Thread.Sleep(1000);
                }


                    driver.SwitchTo().DefaultContent();


            }
            catch (System.Exception ex)
            {

                throw ex;
            }
            return assignmentValue;
        }

        bool driverSwitched = false;

        public void CloseTab(string id)
        {
            try
            {

                driver.SwitchTo().DefaultContent();

                Thread.Sleep(2000); // changed from 2000 to 200

                driverSwitched = true;
                IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                Thread.Sleep(1000); // changed from 1500 to 1000
                js.ExecuteScript("var a = document.getElementsByTagName('a');if(a){for(var i = 0; i< a.length ; i++){if(a[i] && a[i].innerText == '" + id + "'){a[i].previousSibling.click();break;}}}");
                Thread.Sleep(1000); // changed from 1000 to 500
            }
            catch (Exception ex)
            {


            }
        }

        internal void ClickIRorRRTab()
        {
            try
            {

                driver.SwitchTo().DefaultContent();
                IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                js.ExecuteScript("var a = document.getElementsByTagName('a');if(a){for(var i = 0; i< a.length ; i++){if(a[i] && (a[i].innerText == 'IR Queue' || a[i].innerText == 'RR Queue' )) {a[i].click();}}}");
            }
            catch (Exception ex)
            {


            }
        }

        internal void ClickOnTakOverBtn()
        {
            try
            {
                var btnElement = driver.FindElements(By.TagName("button"));
                if (btnElement != null && btnElement.Count > 0)
                {
                    for (int k = 0; k < btnElement.Count; k++)
                    {
                        if (btnElement[k].GetAttribute("innerText").Equals("Take Over"))
                        {
                            btnElement[k].Click();
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {


            }
        }


        internal void ClickOnSaveExitBtn()
        {
            try
            {
                var btnElement = driver.FindElements(By.TagName("button"));
                if (btnElement != null && btnElement.Count > 0)
                {
                    for (int k = 0; k < btnElement.Count; k++)
                    {
                        if (btnElement[k].GetAttribute("innerText").Equals("Save & Exit"))
                        {
                            btnElement[k].Click();
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {


            }
        }



        internal void ClickOnMoreBtn()
        {
            try
            {
                var btnElement = driver.FindElements(By.TagName("button"));
                if (btnElement != null && btnElement.Count > 0)
                {
                    for (int k = 0; k < btnElement.Count; k++)
                    {
                        if (btnElement[k].GetAttribute("innerText").Equals("More"))
                        {
                            btnElement[k].Click();
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {


            }
        }

        internal void ClickOnWaitOnUser()
        {
            try
            {
                var btnElement = driver.FindElements(By.TagName("a"));
                if (btnElement != null && btnElement.Count > 0)
                {
                    for (int k = 0; k < btnElement.Count; k++)
                    {
                        if (btnElement[k].GetAttribute("innerText").Equals("Wait on User"))
                        {
                            btnElement[k].Click();
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {


            }
        }

        internal void ClickOnActivityTab(string tktVal)
        {
            try
            {

                IWebElement frame = driver.FindElement(By.XPath(".//iframe[@title='" + tktVal + "']"));

                driver.SwitchTo().Frame(frame);


                var html = driver.FindElements(By.TagName("div"));
                int count = html.Count();

                bool statusCheck = true;

                if (html != null && html.Count >= 1)
                {
                    for (int i = 0; i < html.Count; i++)
                    {
                        count--;

                        var body = html[i].FindElements(By.TagName("form"));
                        if (body != null && body.Count >= 1)
                        {
                            for (int j = 0; j < body.Count; j++)
                            {
                                var iframe = body[j].FindElements(By.TagName("div"));
                                if (iframe != null && iframe.Count >= 1)
                                {

                                    for (int k = 0; k < iframe.Count; k++)
                                    {
                                        if (statusCheck)
                                        {
                                            var td = iframe[k].FindElements(By.TagName("table"));
                                            if (td != null && td.Count >= 1)
                                            {

                                                for (int l = 0; l < td.Count; l++)
                                                {
                                                    var tds = td[l].FindElements(By.TagName("td"));
                                                    if (tds != null && tds.Count >= 1)
                                                    {
                                                        for (int m = 0; m < tds.Count; m++)
                                                        {
                                                            var anchor = tds[m].FindElements(By.TagName("a"));
                                                            if (anchor != null && anchor.Count >= 1)
                                                            {
                                                                for (int z = 0; z < anchor.Count; z++)
                                                                {
                                                                    if (anchor[z] != null && anchor[z].GetAttribute("id").Equals("X160_t") &&
                                                                       anchor[z].GetAttribute("innerText").Contains("Activities"))
                                                                    {

                                                                        anchor[z].Click();
                                                                        statusCheck = false;

                                                                        break;
                                                                    }
                                                                }
                                                                if (!statusCheck)
                                                                {
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        if (!statusCheck)
                                                        {
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (!statusCheck)
                                                {
                                                    break;
                                                }

                                            }

                                        }
                                    }
                                    if (!statusCheck)
                                    {
                                        break;
                                    }
                                }
                            }
                            if (!statusCheck)
                            {
                                break;
                            }
                        }

                    }
                }

                driver.SwitchTo().DefaultContent();

            }
            catch (Exception ex)
            {


            }
        }

        public void DownloadAllDocumentsfromLibrary()
        {
            try
            {
                ClientContext ctxSite = GetSPOContext();
                string libraryname = "Documents";
                var list = ctxSite.Web.Lists.GetByTitle(libraryname);
                var rootFolder = list.RootFolder;
                string pathString = string.Format(@"{0}{1}\", @"C:\", libraryname);
                if (!Directory.Exists(pathString))
                {
                    System.IO.Directory.CreateDirectory(pathString);
                }

                GetFoldersAndFiles(rootFolder, ctxSite, pathString);

                CreateSPFolder(ctxSite);
            }
            catch (Exception)
            {

                
            }
            
        }


        private static void GetFoldersAndFiles(Microsoft.SharePoint.Client.Folder mainFolder, ClientContext clientContext, string pathString)
        {
            try
            {
                clientContext.Load(mainFolder, k => k.Name, k => k.Files, k => k.Folders);
                clientContext.ExecuteQuery();
                foreach (var folder in mainFolder.Folders)
                {
                    string subfolderPath = string.Format(@"{0}{1}\", pathString, folder.Name);
                    if (!Directory.Exists(subfolderPath))
                    {
                        System.IO.Directory.CreateDirectory(subfolderPath);
                    }

                    GetFoldersAndFiles(folder, clientContext, subfolderPath);
                }

                foreach (var file in mainFolder.Files)
                {
                    var fileName = System.IO.Path.Combine(pathString, file.Name);
                    if (System.IO.File.Exists(fileName))
                    {
                        System.IO.File.Delete(fileName);
                    }
                    if (fileName != null)
                    {
                        if (fileName.ToLower().Contains("Updated_Tickets_From_Utillity.xlsx".ToLower()))
                        {
                            var fileRef = file.ServerRelativeUrl;
                            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileRef);

                            using (var fileStream = System.IO.File.Create(fileName))
                            {
                                fileInfo.Stream.CopyTo(fileStream);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {

            }
        }


        static string UserName = string.Empty;
        static string Pwd = string.Empty;

        private static ClientContext GetSPOContext()
        {
            ClientContext spoContext = null;
            try
            {
                UserName = "bikram.jena@volkswagen.co.in";
                string spsiteurl = "https://volkswagengroup.sharepoint.com/sites/SKODAFIMAMS/";
                Pwd = "Biki@@280491";
                var secure = new SecureString();

                foreach (char c in Pwd)
                {
                    secure.AppendChar(c);
                }


                spoContext = new ClientContext(spsiteurl);
                spoContext.Credentials = new SharePointOnlineCredentials(UserName, secure);
                
            }
            catch (Exception ex)
            {

                
            }
            return spoContext;


        }


        public static void UploadDocument(ClientContext clientContext, string sourceFilePath, string serverRelativeDestinationPath)
        {
            try
            {
                using (var fs = new FileStream(sourceFilePath, FileMode.Open))
                {
                    var fi = new FileInfo(sourceFilePath);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, serverRelativeDestinationPath, fs, true);
                }
            }
            catch (Exception ex)
            {

            }

        }


        public static void UploadFolder(ClientContext clientContext, System.IO.DirectoryInfo folderInfo, Microsoft.SharePoint.Client.Folder folder)
        {
            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;

            try
            {
                files = folderInfo.GetFiles("K2NG Daily Ticket Resolution_Updated.xlsx");



                if (files != null)
                {
                    foreach (System.IO.FileInfo fi in files)
                    {
                        Console.WriteLine(fi.FullName);
                        clientContext.Load(folder);
                        clientContext.ExecuteQuery();
                        string relativeURl = folder.ServerRelativeUrl + "/" + "General/";
                        string[] relativeArr = relativeURl.ToString().Split(new string[] { "Shared Documents" }, StringSplitOptions.RemoveEmptyEntries);
                        relativeArr[0] = relativeArr[0].Trim() + "Shared Documents/General";
                        relativeArr[1] = relativeArr[1].Replace("/General", "");
                        relativeURl = relativeArr[0] + relativeArr[1];

                        UploadDocument(clientContext, fi.FullName, relativeURl + fi.Name);
                    }


                }
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine(e.Message);
            }

            catch (System.IO.DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }


        }



        public void CreateSPFolder(ClientContext ctx)
        {
            try
            {

                Web web = ctx.Web;
                ctx.Load(web, w => w.Lists, w => w.ServerRelativeUrl);
                ctx.ExecuteQuery();

                var query = ctx.LoadQuery(web.Lists.Where(p => p.Title == "Documents"));
                ctx.ExecuteQuery();

                List documentLib = query.FirstOrDefault();
                var folder = documentLib.RootFolder;
                DirectoryInfo di = new DirectoryInfo("C:\\Documents\\General\\Ticket Resolution");
                ctx.Load(documentLib.RootFolder);
                ctx.ExecuteQuery();

                folder = documentLib.RootFolder.Folders.Add(di.Name);
                ctx.ExecuteQuery();

                //DirectoryInfo fl = new DirectoryInfo("C:\\Documents\\General\\Ticket Resolution\\K2NG Daily Ticket Resolution_Updated.xlsx");


                UpdateExcel();

                UploadFolder(ctx, di, folder);

            }
            catch (Exception ex)
            {


            }
        }

        static Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = null;
        static Microsoft.Office.Interop.Excel.Range rng = null;
        static Microsoft.Office.Interop.Excel.Application app = null;
        static Microsoft.Office.Interop.Excel.Workbook wbk = null;

        bool requestCheck = false;

        public void UpdateExcel()
        {
            try
            {
                string path = string.Empty;
                path = @"C:\Documents\General\Ticket Resolution\Updated_Tickets_From_Utillity.xlsx";
                //GetFile();
                xlWorkSheet = null;
                rng = null;
                app = null;
                wbk = null;

                app = new Microsoft.Office.Interop.Excel.Application();
                wbk = app.Workbooks.Open(path, 0, false, 2, string.Empty, string.Empty, true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, true, 0, true, true, false);
                app.Visible = false;

                int count = wbk.Worksheets.Count;
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)wbk.Worksheets.get_Item(2);
                rng = xlWorkSheet.UsedRange;

                string sheet = xlWorkSheet.Name;
                int totalrowcount = xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;

                int totalcellCount = rng.Columns.Count;

                string[] morningshift = new string[] { "06", "07", "08", "09", "10", "11", "12", "13"};

                string[] afternoonShift = new string[] { "14", "15", "16", "17", "18", "19", "20", "21"};

                string[] nightShift = new string[] { "22", "23", "00", "01", "02", "03", "04", "05"};

                if (totalrowcount < 1)
                {

                }
                else
                {
                    totalrowcount = totalrowcount + 1;
                    for (int i = totalrowcount; i <= totalrowcount; i++)
                    {
                        (rng.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2 = TicketVal;
                        //(rng.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value2 = priorityVal;
                        (rng.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value2 = dateString;
                        string sihiftValue = string.Empty;
                        if (morningshift.Contains(shortTime))
                        {
                            sihiftValue = "Morning Shift";
                        }
                        else if (afternoonShift.Contains(shortTime))
                        {
                            sihiftValue = "Afternoon Shift";
                        }
                        else if(nightShift.Contains(shortTime))
                        {
                            sihiftValue = "Night Shift";
                        }
                        (rng.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value2 = sihiftValue;
                        (rng.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value2 = titleVal;
                        (rng.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value2 = statusVal;
                        DateTime time = DateTime.Now;
                        string timeVal = time.ToString("MM/dd/yy");
                        (rng.Cells[i, 10] as Microsoft.Office.Interop.Excel.Range).Value2 = timeVal;
                        (rng.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value2 = "Bikram";
                        (rng.Cells[i, 12] as Microsoft.Office.Interop.Excel.Range).Value2 = "1";
                        //(rng.Cells[i, 13] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        //(rng.Cells[i, 14] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";

                    }
                }
                app.DisplayAlerts = false;
                wbk.Save();
                wbk.Close();
                app.Quit();

            }
            catch (Exception ex)
            {
                app.DisplayAlerts = false;
                wbk.Save();
                wbk.Close();
                app.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(wbk);
                Marshal.ReleaseComObject(app);
            }
            finally
            {
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(wbk);
                Marshal.ReleaseComObject(app);
            }
        }

        internal void ClickIncident()
        {
            try
            {
                Thread.Sleep(3000);
                requestCheck = false;
                //var element = driver.FindElement(By.XPath("//div[@class ='x-panel-header x-panel-header-noborder x-unselectable icon-problem-mgnt icon-hp x-accordion-hd']"));

                IJavaScriptExecutor executor = driver as IJavaScriptExecutor;
                executor.ExecuteScript("document.getElementsByClassName(\"x-panel-header x-panel-header-noborder x-unselectable icon-problem-mgnt icon-hp x-accordion-hd\")[0].click();");


                IJavaScriptExecutor executor1 = driver as IJavaScriptExecutor;
                executor1.ExecuteScript("document.getElementById(\"ROOT/Incident Management/Incident Queue\").click();");

                Thread.Sleep(1000);

            }
            catch (System.Exception ex)
            {
                throw ex;

            }
        }

        List<string> lstTicketID = new List<string>();

        internal void ClickOnIncidentWorkingTicket()
        {

            try
            {
                Thread.Sleep(2000);
                string ticketID = string.Empty;
                IWebElement frame;
                
                if (requestCheck)
                {
                    
                    frame = driver.FindElement(By.XPath(".//iframe[@title='RR Queue']"));
                    
                }
                else
                {
                    frame = driver.FindElement(By.XPath(".//iframe[@title='IR Queue']"));
                    
                }

                
                driver.SwitchTo().Frame(frame);
                

                var html = driver.FindElements(By.TagName("div"));
                int count = html.Count();

                bool statusCheck = true;



                if (html != null && html.Count >= 1)
                {
                    for (int i = 0; i < html.Count; i++)
                    {
                        count--;

                        var body = html[i].FindElements(By.TagName("form"));
                        if (body != null && body.Count >= 1)
                        {
                            for (int j = 0; j < body.Count; j++)
                            {
                                var iframe = body[j].FindElements(By.TagName("div"));
                                if (iframe != null && iframe.Count >= 1)
                                {

                                    for (int k = 0; k < iframe.Count; k++)
                                    {
                                        if (iframe[k] != null && iframe[k].GetAttribute("className").Equals("x-grid3-hd-inner x-grid3-hd-5")
                                            && iframe[k].GetAttribute("innerText").Contains("Status"))
                                        {
                                            iframe[k].Click();
                                            statusCheck = false;
                                            break;
                                        }
                                    }
                                    if (!statusCheck)
                                    { break; }
                                }


                            }

                            if (!statusCheck)
                            { break; }
                        }

                    }
                }
            repeat:


                driver.SwitchTo().DefaultContent();

                Thread.Sleep(2000); // changed from 2000 to 500
                if (requestCheck)
                {

                    frame = driver.FindElement(By.XPath(".//iframe[@title='RR Queue']"));

                }
                else
                {
                    frame = driver.FindElement(By.XPath(".//iframe[@title='IR Queue']"));

                }

                
                statusCheck = true;
                int checkCount = 0;

                Thread.Sleep(2000); // changed from 2000 to 500
                driver.SwitchTo().Frame(frame);
                Thread.Sleep(2000); // changed from 2000 to 500

                var html1 = driver.FindElements(By.TagName("div"));
                if (html1 != null && html1.Count >= 1)
                {
                    for (int i = 0; i < html1.Count; i++)
                    {
                        var body1 = html1[i].FindElements(By.TagName("form"));
                        if (body1 != null && body1.Count >= 1)
                        {
                            for (int j = 0; j < body1.Count; j++)
                            {
                                if (body1[j].GetAttribute("id").Equals("topaz"))
                                {
                                    var iframe1 = body1[j].FindElements(By.TagName("div"));
                                    if (iframe1 != null && iframe1.Count >= 1)
                                    {
                                        for (int k = 0; k < iframe1.Count; k++)
                                        {
                                            if (iframe1[k].GetAttribute("id").Equals("recordListGrid"))
                                            {
                                                if (statusCheck)
                                                {
                                                    var td1 = iframe1[k].FindElements(By.TagName("table"));
                                                    if (td1 != null && td1.Count >= 1)
                                                    {
                                                        //checkCount = td1.Count;
                                                        for (int l = 0; l < td1.Count; l++)
                                                        {
                                                           
                                                                var tdVal1 = td1[l].FindElements(By.TagName("td"));
                                                                if (tdVal1 != null && tdVal1.Count >= 1)
                                                                {
                                                                    for (int m = 0; m < tdVal1.Count; m++)
                                                                    {

                                                                        if (tdVal1[m] != null && tdVal1[m].GetAttribute("innerText").Equals("Working"))
                                                                        {
                                                                            
                                                                            if (!lstTicketID.Contains(tdVal1[m - 4].GetAttribute("innerText")))
                                                                            {
                                                                                ticketID = tdVal1[m - 4].GetAttribute("innerText");                                                                                
                                                                                lstTicketID.Add(ticketID);
                                                                                //Thread.Sleep(500);
                                                                                tdVal1[m - 4].Click();

                                                                                //Thread.Sleep(2000);

                                                                                string[] WorkStatusValue = ClickOnIncidentElement(ticketID);
                                                                                if (WorkStatusValue != null && (WorkStatusValue[0].Equals("Working") || WorkStatusValue[0].Equals("Transferred"))
                                                                                    && WorkStatusValue[1].Equals("K2NG VWITS Support SKODA"))
                                                                                {
                                                                                    ClickOnTakOverBtn();

                                                                                    // ClickOnActivityTab();

                                                                                    //To update file and upload into sharepoint folder
                                                                                    DownloadAllDocumentsfromLibrary();

                                                                                }


                                                                                CloseTab(ticketID);

                                                                                //ClickIRorRRTab();

                                                                                statusCheck = false;
                                                                                break;

                                                                            }
                                                                            
                                                                        }

                                                                    }

                                                                    if (!statusCheck)
                                                                    {
                                                                        break;
                                                                    }
                                                                }
                                                            
                                                           
                                                        }

                                                        if (!statusCheck)
                                                        {
                                                            break;
                                                        }
                                                    }

                                                }
                                            }
                                        }

                                        if (!statusCheck)
                                        {
                                            break;
                                        }
                                    }
                                }
                            }
                            if (!statusCheck)
                            {
                                break;
                            }
                        }


                    }              
                }
                
                if(driverSwitched)
                {
                    driverSwitched = false;
                    goto repeat;
                }
                 

                driver.SwitchTo().DefaultContent();


            }
            catch (System.Exception ex)
            {

                driver.SwitchTo().DefaultContent();

            }
            finally { driver.SwitchTo().DefaultContent(); }

        }

        internal void ClickRequest()
        {
            try
            {
                Thread.Sleep(2000);
                //var element = driver.FindElement(By.XPath("//div[@class ='x-panel-header x-panel-header-noborder x-unselectable icon-problem-mgnt icon-hp x-accordion-hd']"));
                requestCheck = true;
                
                IJavaScriptExecutor executor = driver as IJavaScriptExecutor;
                executor.ExecuteScript("document.getElementsByClassName(\"x-panel-header x-panel-header-noborder x-unselectable icon-service-catalog icon-hp x-accordion-hd\")[0].click();");


                IJavaScriptExecutor executor1 = driver as IJavaScriptExecutor;
                executor1.ExecuteScript("document.getElementById(\"ROOT/Request Fulfillment/Request Queue\").click();");

                Thread.Sleep(1000);

            }
            catch (System.Exception ex)
            {
               

            }
        }


        internal void ClickOnRequestWorkingTicket()
        {

            try
            {
                Thread.Sleep(2000);

                IWebElement reqFrame = driver.FindElement(By.XPath(".//iframe[@title='RR Queue']"));

                driver.SwitchTo().Frame(reqFrame);


                var htmlReq = driver.FindElements(By.TagName("div"));
                int count = htmlReq.Count();

                bool statusCheck = true;



                if (htmlReq != null && htmlReq.Count >= 1)
                {
                    for (int i = 0; i < htmlReq.Count; i++)
                    {
                        count--;

                        var bodyReq = htmlReq[i].FindElements(By.TagName("form"));
                        if (bodyReq != null && bodyReq.Count >= 1)
                        {
                            for (int j = 0; j < bodyReq.Count; j++)
                            {
                                var iframeReq = bodyReq[j].FindElements(By.TagName("div"));
                                if (iframeReq != null && iframeReq.Count >= 1)
                                {

                                    for (int k = 0; k < iframeReq.Count; k++)
                                    {
                                        if (iframeReq[k] != null && iframeReq[k].GetAttribute("className").Equals("x-grid3-hd-inner x-grid3-hd-5")
                                            && iframeReq[k].GetAttribute("innerText").Contains("Status"))
                                        {
                                            iframeReq[k].Click();
                                            statusCheck = false;
                                            break;
                                        }
                                    }
                                    if (!statusCheck)
                                    { break; }
                                }


                            }

                            if (!statusCheck)
                            { break; }
                        }

                    }
                }
                statusCheck = true;

                if (htmlReq != null && htmlReq.Count >= 1)
                {
                    for (int i = 0; i < htmlReq.Count; i++)
                    {

                        var bodyReq = htmlReq[i].FindElements(By.TagName("form"));
                        if (bodyReq != null && bodyReq.Count >= 1)
                        {
                            for (int j = 0; j < bodyReq.Count; j++)
                            {
                                var iframeReq = bodyReq[j].FindElements(By.TagName("div"));
                                if (iframeReq != null && iframeReq.Count >= 1)
                                {

                                    for (int k = 0; k < iframeReq.Count; k++)
                                    {
                                        if (statusCheck)
                                        {
                                            var tdReq = iframeReq[k].FindElements(By.TagName("table"));
                                            if (tdReq != null && tdReq.Count >= 1)
                                            {

                                                for (int l = 0; l < tdReq.Count; l++)
                                                {
                                                    var tdValReq = tdReq[l].FindElements(By.TagName("td"));
                                                    if (tdValReq != null && tdValReq.Count >= 1)
                                                    {
                                                        for (int m = 0; m < tdValReq.Count; m++)
                                                        {

                                                            if (tdValReq[m] != null && tdValReq[m].GetAttribute("innerText").Equals("Working"))
                                                            {

                                                                string val = tdValReq[m].GetAttribute("innerText");

                                                                tdValReq[m - 4].Click();
                                                                statusCheck = false;

                                                                break;
                                                            }

                                                        }

                                                        if (!statusCheck)
                                                        {
                                                            break;
                                                        }
                                                    }

                                                }

                                                if (!statusCheck)
                                                {
                                                    break;
                                                }
                                            }

                                        }
                                    }

                                    if (!statusCheck)
                                    {
                                        break;
                                    }
                                }

                            }
                            if (!statusCheck)
                            {
                                break;
                            }
                        }

                    }
                }

                driver.SwitchTo().DefaultContent();


            }
            catch (System.Exception ex)
            {

                driver.SwitchTo().DefaultContent();

            }
            finally { driver.SwitchTo().DefaultContent(); }

        }

    }



}
