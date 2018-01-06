using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using NLog;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

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
        internal static List<string> existingPages { get; set; }

        static void Main(string[] args)
        {
            Console.Write("Please enter your sharpoint url: ");
            sharepointUrl = Console.ReadLine().Trim();

            Console.Write("Please enter your email adress: ");
            email = Console.ReadLine().Trim();

            Console.Write("Please enter your password: ");
            password = Console.ReadLine().Trim();

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

                existingPages = new List<string>();

                ReadExistingWikiPages(context, existingPages);

                CreateWikiPageAndAddWebparts(context);

            }
        }

        private static void ReadExistingWikiPages(ClientContext context, List<string> existingPages)
        {
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml =  @"<View><Query><Where><IsNotNull><FieldRef Name='FileLeafRef' /></IsNotNull></Where></Query></View>";

            ListItemCollectionPosition position = null;

            do
            {
                camlQuery.ListItemCollectionPosition = position;

                ListItemCollection listItems = context.Web.Lists.GetByTitle(wikiListName).GetItems(camlQuery);
                context.Load(listItems);
                context.ExecuteQuery();

                context.Load(listItems, lic => lic.Include(
                    item => item.Id,
                    item => item["FileLeafRef"]
                    ));
                context.ExecuteQuery();

                position = listItems.ListItemCollectionPosition;

                foreach (ListItem listItem in listItems)
                {
                    existingPages.Add(listItem["FileLeafRef"].ToString());
                }


            } while (position != null);
        }

        private static void CreateWikiPageAndAddWebparts(ClientContext context)
        {
            //Prepare FileNames
            List<string> listFileNames = new List<string>();
            for (int i = 0; i < 5000; i++)
            {
                listFileNames.Add("TestFile_" + i.ToString("D5") + ".aspx");
            }

            //Get the lists
            var wikiLib = context.Web.Lists.GetByTitle(wikiListName);
            var otherList = context.Web.Lists.GetByTitle(otherListName);

            context.Load(wikiLib,
                w => w.Id,
                w => w.Title,
                w => w.RootFolder.ServerRelativeUrl);
            context.Load(otherList,
                o => o.Id,
                o => o.Title,
                o => o.RootFolder.ServerRelativeUrl);
            context.ExecuteQuery();


            foreach (var fileName in listFileNames)
            {

                if (existingPages.Contains(fileName))
                {
                    _log.Debug($"Wiki page {fileName} already exist....");
                    continue;
                }

                //Create wikipage
                _log.Debug($"Creating wiki page {fileName}....");
                context.Site.RootWeb.AddWikiPage(wikiListName, fileName);
                _log.Debug($"SUCCESSFULLY: Creating wiki page {fileName}....");

                context.Web.AddLayoutToWikiPage(WikiPageLayout.TwoColumnsHeaderFooter, wikiLib.RootFolder.ServerRelativeUrl + "/" + fileName);

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                var result = true;

                _log.Debug($"Adding webpart to wikipage {fileName}....");

                for (int i = 1; i < 9; i++)
                {

                    try
                    {
                        var listID = "";
                        var listName = "";
                        var title = "Webpart_" + i;
                        List list = null;

                        if (i % 2 == 0)
                        {
                            listID = wikiLib.Id.ToString();
                            listName = wikiLib.Title;
                            list = wikiLib;
                        }
                        else
                        {
                            listID = otherList.Id.ToString();
                            listName = otherList.Title;
                            list = otherList;
                        }

                        WebPartEntity wp = new WebPartEntity();
                        wp.WebPartXml = CreateXML(listID, listName, title);
                        wp.WebPartIndex = i;
                        wp.WebPartTitle = title;

                        //original
                        //context.Web.AddWebPartToWikiPage(wikiLib.RootFolder.ServerRelativeUrl + "/" + fileName, wp, 1, 1, true);

                        //For debugging
                        AddWebPartToWikiPage(context, wikiLib.RootFolder.ServerRelativeUrl + "/" + fileName, wp, 1, 1, true);

                        SetViewForWebpart(wikiLib.RootFolder.ServerRelativeUrl + "/" + fileName, context, list, wp.WebPartTitle);

                    }
                    catch (Exception e)
                    {
                        _log.Error(e, $"FAILURE: Adding webpart to wikipage {fileName}....");
                        result = false;
                    }
                }

                stopwatch.Stop();

                if (result)
                {
                    _log.Debug($"Success: It took {stopwatch.Elapsed}");
                }
                else
                {
                    _log.Debug($"Failure: It took {stopwatch.Elapsed}");
                }
            }
        }

        private static void AddWebPartToWikiPage(ClientContext ctx, string fileUrl, WebPartEntity webPart, int row, int col, bool addSpace)
        {
            File webPartPage = ctx.Web.GetFileByServerRelativeUrl(fileUrl);

            if (webPartPage == null)
            {
                return;
            }

            ctx.Load(webPartPage);
            ctx.Load(webPartPage.ListItemAllFields);
            ctx.ExecuteQuery();

            string wikiField = (string)webPartPage.ListItemAllFields["WikiField"];

            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(webPart.WebPartXml);
            WebPartDefinition wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "wpz", 0);
            ctx.Load(wpdNew);
            ctx.ExecuteQuery();


            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div? 
            XmlElement layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            XmlElement layoutsZoneInner = layoutsTable.SelectSingleNode(string.Format("tbody/tr[{0}]/td[{1}]/div/div", row, col)) as XmlElement;
            // - space element
            XmlElement space = xd.CreateElement("p");
            XmlText text = xd.CreateTextNode(" ");
            space.AppendChild(text);

            // - wpBoxDiv
            XmlElement wpBoxDiv = xd.CreateElement("div");
            layoutsZoneInner.AppendChild(wpBoxDiv);

            if (addSpace)
            {
                layoutsZoneInner.AppendChild(space);
            }

            XmlAttribute attribute = xd.CreateAttribute("class");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read ms-rte-wpbox";
            attribute = xd.CreateAttribute("contentEditable");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "false";
            // - div1
            XmlElement div1 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div1);
            div1.IsEmpty = false;
            attribute = xd.CreateAttribute("class");
            div1.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read " + wpdNew.Id.ToString("D");
            attribute = xd.CreateAttribute("id");
            div1.Attributes.Append(attribute);
            attribute.Value = "div_" + wpdNew.Id.ToString("D");
            // - div2
            XmlElement div2 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div2);
            div2.IsEmpty = false;
            attribute = xd.CreateAttribute("style");
            div2.Attributes.Append(attribute);
            attribute.Value = "display:none";
            attribute = xd.CreateAttribute("id");
            div2.Attributes.Append(attribute);
            attribute.Value = "vid_" + wpdNew.Id.ToString("D");

            ListItem listItem = webPartPage.ListItemAllFields;
            listItem["WikiField"] = xd.OuterXml;
            listItem.Update();
            ctx.ExecuteQuery();

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

        private static bool SetViewForWebpart(string fileUrl, ClientContext context, List list, string webpartName)
        {
            //Setup
            bool result = false;
            _log.Debug("Inside ListViewWebPartHelper.SetViewForWebpart");

            try
            {
                File file = context.Web.GetFileByServerRelativeUrl(fileUrl);

                LimitedWebPartManager webPartManager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                context.Load(webPartManager,
                    w => w.WebParts,
                    w => w.WebParts.Include(wp => wp.WebPart),
                    w => w.WebParts.Include(wp => wp.WebPart.Properties));
                context.ExecuteQueryRetry();

                WebPartDefinitionCollection webPartDefinitionCollection = webPartManager.WebParts;

                foreach (WebPartDefinition webPartDefinition in webPartDefinitionCollection)
                {
                    WebPart webpart = webPartDefinition.WebPart;

                    Dictionary<string, object> propertyValue = webpart.Properties.FieldValues;

                    if (propertyValue["Title"].ToString().Equals(webpartName))
                    {
                        View listView = list.Views.GetById(webPartDefinition.Id);
                        var viewFields = listView.ViewFields;

                        context.Load(listView, l => l.ListViewXml);
                        context.Load(viewFields);
                        context.ExecuteQueryRetry();

                        if (true)
                        {
                            int styleId = 0;

                            //parse xml
                            XmlDocument doc = new XmlDocument();
                            doc.LoadXml(listView.ListViewXml);

                            XmlElement element = (XmlElement)doc.SelectSingleNode("//View//ViewStyle");
                            if (element == null)
                            {
                                element = doc.CreateElement("ViewStyle");
                                element.SetAttribute("ID", styleId.ToString());
                                doc.DocumentElement.AppendChild(element);
                            }
                            else
                            {
                                element.SetAttribute("ID", styleId.ToString());
                            }

                            listView.ListViewXml = doc.FirstChild.InnerXml;

                            //Performance
                            listView.Update();
                            context.ExecuteQueryRetry();
                        }

                        webpart.Properties["InplaceSearchEnabled"] = false;
                        webpart.Properties["DisableViewSelectorMenu"] = true;
                        webPartDefinition.SaveWebPartChanges();
                        context.ExecuteQueryRetry();

                        if (true)
                        {
                            listView.ViewFields.RemoveAll();
                            listView.ViewFields.Add("ID");
                            listView.ViewFields.Add("FileLeafRef");

                        }

                        listView.RowLimit = 30;
                        listView.TabularView = false;
                        listView.Update();
                        context.ExecuteQueryRetry();

                        result = true;
                        break;
                    }
                }

                return result;
            }
            catch (Exception e)
            {
                _log.Error(e, "Error inside ListViewWebPartHelper.SetViewForWebpart");
                return false;
            }
        }
    }
}
