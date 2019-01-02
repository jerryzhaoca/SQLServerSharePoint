using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Data.SqlClient;
using System.Data;

namespace SQLServerSharePoint
{
    class Program
    {
        static void Main(string[] args)
        {
            string constr = "server=???;database=???;integrated security=SSPI"; // database connection string
            string currDate = DateTime.Now.ToString("yyyy-MM-dd");
            SqlConnection con = new SqlConnection(constr);
            string sql = "???"; // sql query

            var url = "???"; // SharePoint URL
            try
            {
                con.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(sql, con);
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0) // if sql query returns no new jobs
                { Console.WriteLine("No new jobs added."); }
                else // if sql query returns new job(s)
                {


                    Console.WriteLine("Connects to database successfully");
              
                    string[] s = new string[dt.Rows.Count];
                    using (var context = new ClientContext(url))
                    {
                        var listTitle = "Documents"; // SharePoint List
                        var username = "???"; // SharePoint username
                        var template = "Work Orders/template"; // template folder
                        SecureString passWord = new SecureString();
                        foreach (char c in "???".ToCharArray()) passWord.AppendChar(c); // user password
                        context.Credentials = new SharePointOnlineCredentials(username, passWord);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            s[i] = dt.Rows[i][0].ToString();
                            var folderName = s[i];

                            if (!FolderExists(context.Web, listTitle, folderName))
                            {
                                CreateFolder(context.Web, listTitle, folderName);
                                CopyFiles(context.Web, listTitle, template, folderName);
                            }
                            else
                            {
                                Console.WriteLine(folderName + " exists. Failed to create.");
                                continue;
                            }

                        }
                    }
                }
               
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally
            {
                con.Close();
                Console.WriteLine("Task Finished!");
                Console.ReadLine();
            }

        }
        public static bool FolderExists(Web web, string listTitle, string folderUrl)
        {
            var list = web.Lists.GetByTitle(listTitle);
            var folders = list.GetItems(CamlQuery.CreateAllFoldersQuery());
            web.Context.Load(list.RootFolder);
            web.Context.Load(folders);
            web.Context.ExecuteQuery();
            var folderRelativeUrl = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderUrl);
            //Console.WriteLine(folderRelativeUrl);
            return Enumerable.Any(folders, folderItem => (string)folderItem["FileRef"] == folderRelativeUrl);
        }
        private static void CreateFolder(Web web, string listTitle, string folderName)
        {
            var list = web.Lists.GetByTitle(listTitle);
            var folderCreateInfo = new ListItemCreationInformation
            {
                UnderlyingObjectType = FileSystemObjectType.Folder,
                LeafName = folderName
            };
            var folderItem = list.AddItem(folderCreateInfo);
            folderItem.Update();
            Console.WriteLine("Create Folder " + folderName);
            web.Context.ExecuteQuery();
        }

        public static void CopyFiles(Web web, string listTitle, string srcFolder, string destFolder)
        {
            var srcList = web.Lists.GetByTitle(listTitle);
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope='RecursiveAll'></View>";
            qry.FolderServerRelativeUrl = string.Format("/sites/FileSharing/Shared Documents/{0}", srcFolder);
            var srcItems = srcList.GetItems(qry);
            web.Context.Load(srcItems, icol => icol.Include(i => i.FileSystemObjectType, i => i["FileRef"], i => i.File));
            web.Context.ExecuteQuery();

            foreach (var item in srcItems)
            {
                switch (item.FileSystemObjectType)
                {
                    case FileSystemObjectType.Folder:
                        var destFolderUrl = ((string)item["FileRef"]).Replace(srcFolder, destFolder);
                        CreateFolder2(web, destFolderUrl);
                        Console.WriteLine("Create Folder " + destFolderUrl);
                        break;
                    case FileSystemObjectType.File:
                        var destFileUrl = item.File.ServerRelativeUrl.Replace(srcFolder, destFolder);
                        item.File.CopyTo(destFileUrl, true);
                        web.Context.ExecuteQuery();
                        Console.WriteLine("Copy File " + destFileUrl);
                        break;
                }
            }
        }

        private static Folder CreateFolder2(Web web, string folderUrl)
        {
            if (string.IsNullOrEmpty(folderUrl))
                throw new ArgumentNullException("Folder Url could not be empty");

            var folder = web.Folders.Add(folderUrl);
            web.Context.Load(folder);
            web.Context.ExecuteQuery();
            return folder;
        }
    }
}
