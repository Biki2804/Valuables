using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Microsoft.SharePoint.Client;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
//using Condition = System.Windows.Automation.Condition;
using System.Security;
using System.Threading;
using System.Web;
using System.Windows;
using System.Windows.Automation;
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
            popUpTimer.Interval = 5000;
        }

        private void PopUpTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            popUpTimer.Enabled = false;
            //Console.WriteLine(driver.Title.ToString());
            //0x2012E
            string password = "965539";
            //Thread.Sleep(1500);
            try
            {
                var MainWindow = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                    new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.ClassNameProperty, "PasswordBox"),
                        new PropertyCondition(AutomationElement.NameProperty, "PIN")));


                if (MainWindow != null)
                {
                    ValuePattern userNamePattern = (ValuePattern)MainWindow.GetCurrentPattern(ValuePattern.Pattern);
                    userNamePattern.SetValue(password);


                    var MainWindow2 = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                             new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                                 new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                                 new PropertyCondition(AutomationElement.NameProperty, "OK")));



                    InvokePattern buttonPattern = (InvokePattern)MainWindow2.GetCurrentPattern(InvokePattern.Pattern);
                    buttonPattern.Invoke();
                    popUpTimer.Enabled = false;
                }
                else { popUpTimer.Enabled = true; }

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
        internal string ReadOutlook()
        {
            string URL = string.Empty;

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
                        string counterVal =  counter.ToString();
                        subject = item.Subject;
                        string longDate = item.SentOn.ToString();
                        string senderName = item.ReceivedByName.ToString();
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


        private void Button_Click(object sender, RoutedEventArgs e)
        {

            ReadOutlook();

            var options = new EdgeOptions();
            options.BinaryLocation = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";
            driver = new EdgeDriver(@"C:\Bikram\edgedriver_win64", options);
            popUpTimer.Enabled = true;

            TicketVal = url.Substring(url.Length - 10);

            driver.Navigate().GoToUrl(url);

            Thread.Sleep(7000);

            string[] WorkStatusValue = ClickOnIncidentElement();
            if (WorkStatusValue != null && WorkStatusValue[0].Equals("Working") && WorkStatusValue[1].Equals("K2NG VWITS Support SKODA"))
            {
                ClickOnTakOverBtn();

                // ClickOnActivityTab();

                //To update file and upload into sharepoint folder
                DownloadAllDocumentsfromLibrary();
            }
            driver.Close();
            driver.Quit();
            //driver.Dispose();
            //Marshal.ReleaseComObject(driver); 
        }

        internal void ClickIncident()
        {
            try
            {
                //var element = driver.FindElement(By.XPath("//div[@class ='x-panel-header x-panel-header-noborder x-unselectable icon-problem-mgnt icon-hp x-accordion-hd']"));

                IJavaScriptExecutor executor = driver as IJavaScriptExecutor;
                executor.ExecuteScript("document.getElementsByClassName(\"x-panel-header x-panel-header-noborder x-unselectable icon-problem-mgnt icon-hp x-accordion-hd\")[0].click();");


                IJavaScriptExecutor executor1 = driver as IJavaScriptExecutor;
                executor1.ExecuteScript("document.getElementById(\"ROOT/Incident Management/Incident Queue\").click();");

                Thread.Sleep(500);

            }
            catch (System.Exception ex)
            {
                throw ex;

            }
        }
        string TicketVal = string.Empty;
        string statusValue = string.Empty;

        string[] assignmentValue = new string[5];
        internal string[] ClickOnIncidentElement()
        {
            assignmentValue = new string[5];
            try
            {


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
                                                    if (td[l] != null && td[l].GetAttribute("id").Equals("X18"))
                                                    {

                                                        assignmentValue[0] = td[l].GetAttribute("value");
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
                                                    if (td[l] != null && td[l].GetAttribute("id").Equals("X77Readonly"))
                                                    {

                                                        assignmentValue[1] = td[l].GetAttribute("value");
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

        internal void ClickOnTakOverBtn()
        {
            try
            {
                var btnElement = driver.FindElements(By.TagName("button"));
                if (btnElement != null && btnElement.Count > 0)
                {
                    for (int k = 0; k < btnElement.Count; k++)
                    {
                        if (btnElement[k].GetAttribute("innerText").Equals("Exit"))
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

        internal void ClickOnActivityTab()
        {
            try
            {

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
                        if (fileName.ToLower().Contains("K2NG Daily Ticket Resolution_Updated.xlsx".ToLower()))
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



        public static void CreateSPFolder(ClientContext ctx)
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

        public static void UpdateExcel()
        {
            try
            {
                string path = string.Empty;
                path = @"C:\Documents\General\Ticket Resolution\K2NG Daily Ticket Resolution_Updated.xlsx";
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

                if (totalrowcount < 1)
                {

                }
                else
                {
                    totalrowcount = totalrowcount + 1;
                    for (int i = totalrowcount; i <= totalrowcount; i++)
                    {
                        (rng.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 9] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 10] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 12] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 13] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";
                        (rng.Cells[i, 14] as Microsoft.Office.Interop.Excel.Range).Value2 = "TicketNo";

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

    }



}
