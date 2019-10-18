using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;
using ClientSidePage = OfficeDevPnP.Core.Pages.ClientSidePage;
using OfficeDevPnP.Core.Pages;
using Newtonsoft.Json.Linq;
using System.Security;
using System.Net;

namespace ProvisionQuicklinksWebPart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var ctx = new ClientContext("https://contoso.sharepoint.com/"))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                SecureString theSecureString = new NetworkCredential("", "Password").SecurePassword;
                ctx.Credentials = new SharePointOnlineCredentials("username@contoso.onmicrosoft.com", theSecureString);

                ctx.Load(ctx.Site, s => s.Id);
                ctx.Load(ctx.Web, w => w.Id, w => w.Url);
                ctx.ExecuteQueryRetry();
                var webInfo = new WebInfo
                {
                    SiteId = ctx.Site.Id.ToString(),
                    WebId = ctx.Web.Id.ToString(),
                    WebUrl = ctx.Web.Url
                };
                // quick links to be added
                List<QuickLinkItem> quickLinkItems = new List<QuickLinkItem>
                {
                    // internal link
                    new QuickLinkItem
                    {
                        UniqueId = Guid.NewGuid().ToString(),
                        Url = "/teams/cl-00103/win/Forms/AllItems.aspx",
                        Title = "All Items",
                        ImageUrl= "/sites/CodeStore/SiteAssets/CC-Logo-only-red.png",
                        ThumbnailType = ThumbnailType.Image
                    },
                    // external link
                    new QuickLinkItem
                    {
                        UniqueId = Guid.NewGuid().ToString(),
                        Url = "https://www.bing.com/search?q=Google+Chrome",
                        Title = "Google Chrome",
                        ThumbnailType = ThumbnailType.Icon,
                        IconName="website"
                    }
                };
                
                // Get home page (welcome page)
                var welcomePage = GetWelcomePage(ctx);

                // Create an empty quick links web part
                var webpart = CreateEmptyQuickLinkWebPart(welcomePage, ctx);
                welcomePage.AddControl(webpart, 0);

                // Create quick links web part with some links
                webpart = CreateQuickLinksWebPart(webInfo, welcomePage, quickLinkItems);
                welcomePage.AddControl(webpart, 0);

                // Update an existing quick links web part
                webpart = GetQuickLinkWebPart(welcomePage);
                UpdateQuicklinkWebPart(webpart, quickLinkItems,webInfo);

               welcomePage.Save();
            }
            Console.WriteLine("Done!");
            Console.ReadKey();
        }

        // Return first quick links web part
        static ClientSideWebPart GetQuickLinkWebPart(ClientSidePage page)
        {
            var clientSideControls = page.Controls.Where(c => c.Type.Name == "ClientSideWebPart").ToList();
            var webParts = clientSideControls.ConvertAll(w => w as ClientSideWebPart);
            var quickLinks = webParts.Where(w => w.Title == "Quick links").ToList();
            if (quickLinks.Count > 0)
                return quickLinks.First();
            else
                return null;
        }

        // Return home page
        static ClientSidePage GetWelcomePage(ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.RootFolder, w => w.ServerRelativeUrl);
            ctx.ExecuteQuery();
            File oFile = ctx.Web.GetFileByServerRelativeUrl(ctx.Web.ServerRelativeUrl + ctx.Web.RootFolder.WelcomePage);
            ctx.Load(oFile);
            ctx.ExecuteQuery();
            var page = ClientSidePage.Load(ctx, oFile.Name);
            return page;
        }

        // Find an existing web part by web part Id
        static ClientSideWebPart FindWebPartById(ClientSidePage page, ClientContext ctx, string webPartId)
        {
            var clientSideWebParts = page.Controls.Where(c => c.Type.Name == "ClientSideWebPart").ToList();
            var webParts = clientSideWebParts.ConvertAll(w => w as ClientSideWebPart);
            return webParts.FirstOrDefault(w => w.WebPartId == webPartId);
        }

        // Add new link to quicklink web part
        static void UpdateQuicklinkWebPart(ClientSideWebPart webpart, List<QuickLinkItem> quickLinks,WebInfo webInfo)
        {
            var propertiesJson = JObject.Parse(webpart.PropertiesJson);
            var objServerProcessedContent = propertiesJson.Property("serverProcessedContent");
            var serverProcessedContent = webpart.ServerProcessedContent;
            var searchablePlainTexts = (JObject)serverProcessedContent["searchablePlainTexts"];
            var imageSources = (JObject)serverProcessedContent["imageSources"];
            var links = (JObject)serverProcessedContent["links"];
            var items = (JArray)propertiesJson["items"];
            var linkIndex = GetLinkIndex(searchablePlainTexts);

            foreach(var qlink in quickLinks)
            {
                var title = new JProperty($"items[{ linkIndex}].title", qlink.Title);
                var description = new JProperty($"items[{ linkIndex}].description", qlink.Description);
                var altText = new JProperty($"items[{ linkIndex}].altText", qlink.AltText);
                var link = new JProperty($"items[{ linkIndex}].sourceItem.url", qlink.Url);
                
                // update searchablePlainTexts     
                searchablePlainTexts.Add(title);
                searchablePlainTexts.Add(description);
                searchablePlainTexts.Add(altText);

                // update images
                if (!string.IsNullOrEmpty(qlink.ImageUrl))
                    imageSources.Add($"items[{linkIndex}].image.url", qlink.ImageUrl);

                // update links
                links.Add(link);

                // update items
                var item = GetQuickLinkItem(linkIndex + 2, webInfo, qlink);
                items.Add(item);
                linkIndex += 1;
            }
            
            // update PropertiesJson
            if (objServerProcessedContent != null)
            {
                propertiesJson.Property("serverProcessedContent").Remove();
            }
            propertiesJson.AddFirst(new JProperty("serverProcessedContent", serverProcessedContent));
            webpart.PropertiesJson = Newtonsoft.Json.JsonConvert.SerializeObject(propertiesJson);
        }

        // Return last index
        static int GetLinkIndex(JObject properties)
        {
            // return 0 if web part has no links
            var index = 0;

            if (properties?.Count > 0)
            {
                // get the last link index, and return the index+1
                var property = (JProperty)properties.Last;
                var name = property.Name;
                var firstIndex = name.IndexOf('[');
                name = name.Substring(firstIndex + 1);
                var lastIndex = name.IndexOf(']');
                index = int.Parse(name.Substring(0, lastIndex)) + 1;
            }

            return index;
        }

        // Return new quicklinks web part instance
        static ClientSideWebPart CreateQuickLinksWebPart(WebInfo webInfo, ClientSidePage page, List<QuickLinkItem> quickLinks)
        {
            JObject serverProcessedContent = GetQuickLinksServerProcessedContent(quickLinks);
            var quickLinksWebPart = page.InstantiateDefaultWebPart(DefaultClientSideWebParts.QuickLinks);
            quickLinksWebPart.PropertiesJson = JObject.FromObject(new
            {
                controlType = 3,
                displayMode = 2,
                instanceId = quickLinksWebPart.InstanceId,
                id = "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
                position = new
                {
                    zoneIndex = 1,
                    sectionIndex = 1,
                    controlIndex = 1
                },
                webPartId = quickLinksWebPart.WebPartId,
                webPartData = new
                {
                    id = quickLinksWebPart.WebPartId,
                    instanceId = Guid.NewGuid().ToString(),
                    title = "Quick links",
                    description = "Add links to important documents and pages.",
                    serverProcessedContent,
                    dataVersion = "2.2",
                    properties = new
                    {
                        items = quickLinks.Select((QuickLinkItem quickLink, int index) => GetQuickLinkItem(index + 1, webInfo, quickLink)),
                        isMigrated = true,
                        layoutId = "CompactCard",
                        shouldShowThumbnail = true,
                        hideWebPartWhenEmpty = true,
                        dataProviderId = "QuickLinks",
                        webId= webInfo.WebId,
                        siteId = webInfo.SiteId,
                        baseUrl = webInfo.WebUrl
                    }
                }
            }).ToString();

            return quickLinksWebPart;
        }
        // Return quick link item
        static JObject GetQuickLinkItem(int quickLinkItemId, WebInfo webInfo, QuickLinkItem quickLinkItem)
        {
            var siteUri = new Uri(webInfo.WebUrl);
            bool external = quickLinkItem.Url[0] != '/' && !quickLinkItem.Url.StartsWith($"https://{siteUri.Host}", StringComparison.CurrentCultureIgnoreCase);
            string blankGuid = new Guid().ToString();
            JObject item = JObject.FromObject(new
            {
                id = quickLinkItemId,
                itemType = 2,
                thumbnailType= quickLinkItem.ThumbnailType,
                siteId = external ? blankGuid : webInfo.SiteId,
                webId = external ? blankGuid : webInfo.WebId,
                uniqueId = quickLinkItem.UniqueId,
                fileExtension = string.Empty,
                progId = string.Empty
            });

            if(quickLinkItem.ThumbnailType == ThumbnailType.Image)
            {
                var imageObject = JObject.FromObject(new
                {
                    guids= new
                    {
                        listId = Guid.NewGuid(),
                        siteId = webInfo.SiteId,
                        webId = webInfo.WebId,
                        uniqueId = quickLinkItem.UniqueId,
                    },
                    imageFit=2
                });
                var image = new JProperty("image", imageObject);
                item.Add(image);
            }

            if(quickLinkItem.ThumbnailType == ThumbnailType.Icon)
            {
                var iconProperty = JObject.FromObject(new
                {
                     iconName = quickLinkItem.IconName
                });
                var fabricReactIcon = new JProperty("fabricReactIcon", iconProperty);
                item.Add(fabricReactIcon);
            }

            return item;
        }

        // Return an empty ServerProcessedContent object
        static JObject GetQuickLinksServerProcessedContent(List<QuickLinkItem> quickLinkItems)
        {
            JObject searchablePlainTexts = new JObject();
            JObject imageSources = new JObject();
            JObject links = new JObject();
            JObject componentDependencies = new JObject();
            
            componentDependencies.Add("layoutComponentId", "706e33c8-af37-4e7b-9d22-6e5694d92a6f");

            for (var index = 0; index < quickLinkItems.Count; index++)
            {
                var quickLink = quickLinkItems[index];
                searchablePlainTexts.Add($"items[{index}].title", quickLink.Title);
                searchablePlainTexts.Add($"items[{index}].description", string.Empty);
                if (!string.IsNullOrEmpty(quickLink.ImageUrl))
                    imageSources.Add($"items[{index}].image.url", quickLink.ImageUrl);              
                links.Add($"items[{index}].sourceItem.url", quickLink.Url);
            }

            JObject serverProcessedContent = JObject.FromObject(new
            {
                htmlStrings = new JObject(),
                searchablePlainTexts,
                imageSources,
                links,
                componentDependencies
            });

            return serverProcessedContent;
        }

        // Return new quick link web part instance
        static ClientSideWebPart CreateEmptyQuickLinkWebPart(ClientSidePage page, ClientContext clientContext)
        {

            clientContext.Load(clientContext.Site, s => s.Id);
            clientContext.Load(clientContext.Web, w => w.Id, w => w.Url);
            clientContext.Load(page.PageListItem, p => p.Id);
            clientContext.ExecuteQueryRetry();

            string siteId = clientContext.Site.Id.ToString();
            string webId = clientContext.Web.Id.ToString();
            string webUrl = clientContext.Web.Url;

            var quickLinksWebPart = page.InstantiateDefaultWebPart(DefaultClientSideWebParts.QuickLinks);

            quickLinksWebPart.PropertiesJson = JObject.FromObject(new
            {
                controlType = 3,
                displayMode = 2,
                id = "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
                position = new
                {
                    zoneIndex = 1,
                    sectionIndex = 1,
                    controlIndex = 1
                },
                webPartId = quickLinksWebPart.WebPartId,
                webPartData = new
                {
                    id = quickLinksWebPart.WebPartId,
                    instanceId = Guid.NewGuid().ToString(),
                    title = "Quick links",
                    description = "Add links to important documents and pages.",
                    serverProcessedContent = new
                    {
                        searchablePlainTexts = new { },
                        imageSources = new { },
                        links = new
                        {
                            baseUrl = webUrl
                        },
                        componentDependencies = new
                        {
                            layoutComponentId = "706e33c8-af37-4e7b-9d22-6e5694d92a6f"
                        }
                    },
                    dataVersion = "2.2",
                    properties = new
                    {
                        items = new JArray(),
                        isMigrated = true,
                        layoutId = "CompactCard",
                        shouldShowThumbnail = true,
                        hideWebPartWhenEmpty = true,
                        dataProviderId = "QuickLinks",
                        webId,
                        siteId,
                        baseUrl = clientContext.Web.Url
                    }
                }
            }).ToString();

            return quickLinksWebPart;
        }
    }
}
