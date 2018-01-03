using Microsoft.SharePoint.Client;
using NLog;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace sp_online_wikipagewebpart
{
    class Program
    {
        private static Logger _log = LogManager.GetCurrentClassLogger();

        internal static string email { get; set; }
        internal static string password { get; set; }
        internal static string sharepointUrl { get; set; }
        internal static string wikiListName { get; set; }
        internal static string otherListName { get; set; }

        static void Main(string[] args)
        {
            Console.Write("Please enter your sharpoint url: ");
            sharepointUrl = Console.ReadLine().Trim();

            Console.Write("Please enter your email adress: ");
            email = Console.ReadLine().Trim();

            Console.Write("Please enter your password: ");
            email = Console.ReadLine().Trim();

            ConnectToSP(email, password);
        }

        private static void ConnectToSP(string email, string password)
        {
            Console.Clear();

            using (ClientContext context = new ClientContext(sharepointUrl))
            {
                Console.WriteLine($"Conneting to SharePoint ('sharepointUrl')...");

                SecureString securePassword = new SecureString();
                foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                context.Credentials = new SharePointOnlineCredentials(email, securePassword);

                Console.Write("Please enter the name of your wiki library: ");
                wikiListName = Console.ReadLine().Trim();

                Console.Write("Please enter the name of a second list / library: ");
                otherListName = Console.ReadLine().Trim();

                CreateWikiPageAndAddWebparts(context);

            }
        }

        private static void CreateWikiPageAndAddWebparts(ClientContext context)
        {
            //Prepare FileNames
            List<string> listFileNames = new List<string>();
            for (int i = 0; i < 1000; i++)
            {
                listFileNames.Add("TestFile_" + i.ToString("D5") + ".aspx");
            }

            //Get the lists
            var wikiLib = context.Web.Lists.GetByTitle(wikiListName);
            var otherList = context.Web.Lists.GetByTitle(otherListName);

            context.Load(wikiLib);
            context.Load(otherList);
            context.ExecuteQuery();

            
            foreach (var fileName in listFileNames)
            {
                //Create wikipage
                _log.Debug($"Creating wiki page {fileName}....");
                context.Site.RootWeb.AddWikiPage(wikiListName, fileName);
                _log.Debug($"SUCCESSFULLY: Creating wiki page {fileName}....");

                for (int i = 1; i < 9; i++)
                {
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();

                    _log.Debug($"Adding webpart to wikipage {fileName}....");

                    try
                    {
                        var listID = "";
                        var listName = "";
                        var title = "Webpart_" + i;

                        if (i % 2 == 0)
                        {
                            listID = wikiLib.Id.ToString();
                            listName = wikiLib.Title;
                        }
                        else
                        {
                            listID = otherList.Id.ToString();
                        }

                        WebPartEntity wp = new WebPartEntity();
                        wp.WebPartXml = CreateXML(listID, listName, title);
                        wp.WebPartIndex = i;
                        wp.WebPartTitle = title;

                        context.Web.AddWebPartToWikiPage(wikiLib.RootFolder + "/" + fileName, wp, 1, 1, true);
                        stopwatch.Stop();

                        _log.Debug($"SUCCESS: Adding webpart to wikipage {fileName}....");
                        _log.Debug($"SUCCESS: It tooks: {stopwatch.Elapsed}");
                    }
                    catch (Exception e)
                    {
                        _log.Error(e, $"FAILURE: Adding webpart to wikipage {fileName}....");
                    }
                }               

            }
        }

        private static string CreateXML(string listId, string listName, string title)
        {
            string xmlListViewWebPart = "<webParts>" +
                                        "<webPart xmlns='http://schemas.microsoft.com/WebPart/v3'>" +
                                        "<metaData>" +
                                        "<type name='Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />" +
                                        "<importErrorMessage>Cannot import this Web Part.</importErrorMessage>" +
                                        "</metaData>" +
                                        "<data>" +
                                        "<properties>" +
                                        "<property name='ShowWithSampleData' type='bool'>False</property>" +
                                        "<property name='Default' type='string' />" +
                                        "<property name='NoDefaultStyle' type='string' null='true' />" +
                                        "<property name='CacheXslStorage' type='bool'>True</property>" +
                                        "<property name='ViewContentTypeId' type='string' />" +
                                        "<property name='XmlDefinitionLink' type='string' />" +
                                        "<property name='ManualRefresh' type='bool'>False</property>" +
                                        "<property name='ListUrl' type='string' />" +
                                        $"<property name='ListId' type='System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'>{listId}</property>" +
                                        "<property name='TitleUrl' type='string'></property>" +
                                        "<property name='EnableOriginalValue' type='bool'>False</property>" +
                                        "<property name='Direction' type='direction'>NotSet</property>" +
                                        "<property name='ServerRender' type='bool'>False</property>" +
                                        "<property name='ViewFlags' type='Microsoft.SharePoint.SPViewFlags, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'>Html, TabularView, Hidden, Mobile</property>" +
                                        "<property name='AllowConnect' type='bool'>True</property>" +
                                        $"<property name='ListName' type='string'>{listName}</property>" +
                                        "<property name='ListDisplayName' type='string' />" +
                                        $"<property name='Title' type='string'>{title}</property>" +
                                        "<property name='ShowToolbarWithRibbon' type='bool'>False</property>";

            xmlListViewWebPart += "<property name='ChromeType' type='chrometype'>TitleOnly</property>";

            
            xmlListViewWebPart += "<property name='InplaceSearchEnabled' type='bool'>False</property>" +
                                  "</properties></data></webPart></webParts>";

            return xmlListViewWebPart;
        }


    }
}
