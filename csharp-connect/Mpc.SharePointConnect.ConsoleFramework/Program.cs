using System;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Security;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;

namespace Mpc.SharePointConnect.ConsoleFramework
{
    internal class Program
    {
        private static void GetFicheiro()
        {
            var url = "https://mpcjellycode.sharepoint.com/sites/Site";
            var username = "mcosta@mpcjellycode.onmicrosoft.com";
            var password = "JellyCode2018";
            var securedPassword = new SecureString();
            password.ToList().ForEach(securedPassword.AppendChar);
            var credentials = new SharePointOnlineCredentials(username, securedPassword);

            using (var clientContext = new ClientContext(url))
            {
                clientContext.Credentials = credentials;

                var web = clientContext.Web;

                clientContext.Load(web, website => website.ServerRelativeUrl);
                clientContext.ExecuteQuery();

                Console.WriteLine(clientContext.ServerVersion.ToString(4));

                clientContext.Load(web.Lists, GetListQuery());
                clientContext.ExecuteQuery();

                foreach (List list in web.Lists)
                {
                    if (list.Title != "Documents")
                    {
                        continue;
                    }

                    Console.WriteLine("List title is: " + list.Title);
                    Console.WriteLine(list.Id);
                    Console.WriteLine("RootFolder: " + list.RootFolder.ServerRelativePath.DecodedUrl);
                    //Console.WriteLine(string.Join(", ", list.Fields.Select(x => x.InternalName)));
                    var ids = list.RootFolder.Files.Select(x => x.UniqueId).ToList();

                    Console.WriteLine("Files: " + string.Join(Environment.NewLine, list.RootFolder.Files.Select(x => x.ServerRelativePath.DecodedUrl + " " + x.UniqueId)));
                    Console.WriteLine("Subfolders: " + string.Join(", ", list.RootFolder.Folders.Select(x => x.ServerRelativePath.DecodedUrl)));
                    Console.WriteLine(list.EnableAttachments);

                    foreach (var id in ids)
                    {
                        var file = web.GetFileById(id);
                        clientContext.Load(file);
                        clientContext.ExecuteQuery();

                        var fullPath = $"{file.ServerRelativeUrl}";

                        var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fullPath);

                        var memoryStream = new MemoryStream();

                        fileInfo.Stream.CopyTo(memoryStream);

                        var bytes = memoryStream.ToArray();
                    }

                    Console.WriteLine(Environment.NewLine);
                }

                Console.ReadLine();
            }
        }

        private static Expression<Func<ListCollection, object>> GetListQuery()
        {
            return lists => lists.Include(
                list => list.Title,
                list => list.Id,
                list => list.RootFolder,
                list => list.RootFolder.ServerRelativePath,
                list => list.RootFolder.Files,
                list => list.RootFolder.Files.Include(f => f.ServerRelativePath, f => f.UniqueId),
                list => list.RootFolder.Folders,
                list => list.RootFolder.Folders.Include(f => f.ServerRelativePath),
                list => list.RootFolder.Folders.Include(f => f.Files),
                list => list.Fields,
                list => list.EnableAttachments);
        }

        private static void GetSiteCollections()
        {
            //Make sure to update the admin url and the credentials 
            string adminCenterUrl = "https://mpcjellycode-admin.sharepoint.com/";
            string userName = "mcosta@mpcjellycode.onmicrosoft.com";
            string password = "JellyCode2018";

            ClientContext adminCtx = new ClientContext(adminCenterUrl);

            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);

            adminCtx.Credentials = new SharePointOnlineCredentials(userName, secureString);

            Tenant tenant = new Tenant(adminCtx);
            SPOSitePropertiesEnumerable props = tenant.GetSiteProperties(0, true);
            adminCtx.Load(props);
            adminCtx.ExecuteQuery();

            if (props != null && props.Count > 0)
            {
                Console.WriteLine("TITLE \t URL \t COMPATIBILITY LEVEL");
                foreach (SiteProperties prop in props)
                {
                    Console.WriteLine(prop.Title + "\t" + prop.Url + "\t" + prop.CompatibilityLevel.ToString());
                }
            }
            Console.ReadLine();
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            GetFicheiro();
            GetSiteCollections();
        }
    }
}
