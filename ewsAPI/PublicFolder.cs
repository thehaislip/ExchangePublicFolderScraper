using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using ewsAPI.Models;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using System.IO;
using MoreLinq;

namespace ewsAPI
{
    public class PublicFolder
    {
        public Folder GetFolderByPath(string path,string username, string password, string email) {
            path = path.Replace('/','\\').Replace("\\\\","");
            var ar = path.Split('\\');
            var service = GetEWSService(username,password,email);
            var fv = new FolderView(1000);
            fv.PropertySet = GetPropertySet();
            fv.PropertySet.Add(EWSProperties.PR_PF_Proxy);
            var folders = service.FindFolders(WellKnownFolderName.PublicFoldersRoot, fv).Folders.FirstOrDefault(e => {
                e.TryGetProperty(EWSProperties.PR_Display_name, out string displayName);
                return displayName == ar[0];
            });
            var i = 0;
            foreach (var p in ar)
            {
                if (i > 0)
                {
                    folders = folders.FindFolders(fv).FirstOrDefault(e => {
                        e.TryGetProperty(EWSProperties.PR_Display_name, out string displayName);
                        return displayName == p;
                    });
                }
                i++;
            }
            return folders;
        }

        public IEnumerable<PublicFolderModel> GetAllFolders(string username,string password,string email,IEnumerable<KeyValuePair<string,string>> toDelete = null)
        {
            ExchangeService service = GetEWSService(username,password,email);

            var fv = new FolderView(2500);


            fv.PropertySet = GetPropertySet();
            //fv.Traversal = FolderTraversal.Deep;
            var w = new MAPI();


            //fv.Traversal = FolderTraversal.Shallow;
            //new FolderId("AQEuAAADGkRzkKpmEc2byACqAC/EWgMAbN4R9X4ytEC8CYU/2LnSxQACFO0hvgAAAA==")
            //WellKnownFolderName.PublicFoldersRoot
            var fList = new List<PublicFolderModel>();

            var folders = service.FindFolders(WellKnownFolderName.PublicFoldersRoot, fv);

            foreach (var item in folders.Folders)
            {
                fList.Add(ParseFolder(item,toDelete));
                //item.TryGetProperty(EWSProperties.Pr_Folder_Path, out string path);
                fList.AddRange(GetFolder(item, fv,toDelete).ToList());

            }

            // e.TryGetProperty(EWSProperties.PR_Display_name, out string displayName);
            var t = fList.Count(e => e.NumberOfRules > 0);

            return fList;
        }
        private IEnumerable<PublicFolderModel> GetFolder(Folder baseFolder, FolderView fv,IEnumerable<KeyValuePair<string,string>> toDelete = null)
        {

            var fList = new List<PublicFolderModel>();
            baseFolder.TryGetProperty(EWSProperties.Pr_Folder_Path, out string path);
            //if (path.ToLower().Contains("weather forecasts"))
            //{
            //    System.Diagnostics.Debugger.Break();
            //}

            if (baseFolder.ChildFolderCount == 0)
            {
                fList.Add(ParseFolder(baseFolder,toDelete));
                return fList;
            }
            var q = baseFolder.FindFolders(fv);
            loopthroughFolders(fv, fList, q);
            var moreAvailable = q.MoreAvailable;
            while (moreAvailable)
            {
                //loopthroughFolders(fv, fList, q);
                var folderV = new FolderView(1000);
                folderV.PropertySet = GetPropertySet();
                folderV.Offset = q.NextPageOffset.GetValueOrDefault();
                q = baseFolder.FindFolders(folderV);
                moreAvailable = q.MoreAvailable;
                loopthroughFolders(folderV, fList, q,toDelete);
            }

            return fList;
        }

        private void loopthroughFolders(FolderView fv, List<PublicFolderModel> fList, FindFoldersResults q, IEnumerable<KeyValuePair<string,string>> toDelete = null)
        {
            foreach (var item in q)
            {

                fList.Add(ParseFolder(item,toDelete));
                fList.AddRange(GetFolder(item, fv,toDelete));
            }
        }

        private PublicFolderModel ParseFolder(Folder folder, IEnumerable<KeyValuePair<string,string>> toDelete = null)
        {

            folder.TryGetProperty(EWSProperties.Pr_Folder_Path, out string path);
            folder.TryGetProperty(EWSProperties.HasRules, out bool hasRule);
            folder.TryGetProperty(EWSProperties.PR_FolderSize, out long folderSize);
            Tuple<int, int, int> rules = null;
            Tuple<int, DateTime?, string, DateTime?, int, string, string> messageDetails = null;
            string error = "";
            //if (path.ToLower().Contains("ewspfoldertest"))
            //{
            //    System.Diagnostics.Debugger.Break();
            //}
            if (toDelete != null)
            {
                var delete = toDelete.Where(e => e.Key.ToLower() == path.ToLower()).ToList();
                if (delete.Count() > 0)
                {
                    DeleteMessages(folder, delete);
                }
                             
            }

            if (folder.EffectiveRights.HasFlag(EffectiveRights.Read))
            {
                try
                {
                    rules = GetRules(folder);
                    messageDetails = GetMessages(folder);
                }
                catch (Exception ex)
                {
                    error = ex.Message;

                }

            }
            else
            {
                error = $"Do not have permission to read contents of {path}";
            }
            return new PublicFolderModel()
            {
                FolderID = folder.Id.ToString(),
                //FolderSize = folderSize,
                FolderPath = path,
                NumberOfRules = rules?.Item1,
                NumberOfActieRules = rules?.Item2,
                NumberOfDisabledRules = rules?.Item3,
                //  TotalSize = messageDetails?.Item1,
                LatestDateReceived = messageDetails?.Item2,
                LastModifiedName = messageDetails?.Item3,
                LastModifiedTime = messageDetails?.Item4,
                Subject = messageDetails?.Item6,
                From = messageDetails?.Item7,
                //   ItemCount = messageDetails?.Item5,

                Error = error
            };
        }

        private void DeleteMessages(Folder folder, IEnumerable<KeyValuePair<string, string>> toDelete)
        {
            if (folder.EffectiveRights.HasFlag(EffectiveRights.Read))
            {
                var m = GetAllMessages(folder);
                foreach (Item item in m)
                {
                    if (toDelete.Any(e => e.Value == item.Id.ToString()))
                    {
                        item.Delete(DeleteMode.SoftDelete);
                    }

                }
            }
        }

       

        private Tuple<int, int, int> GetRules(Folder folder)
        {
            var iv = new ItemView(1000);
            iv.Traversal = ItemTraversal.Associated;

            var ruleCount = 0;
            var enableCount = 0;
            var disabledCount = 0;

            foreach (var item in folder.FindItems(iv).ToList())
            {
                if (item.ItemClass == "IPM.Rule.Version2.Message")
                {
                    PropertySet propset;

                    propset = new PropertySet(BasePropertySet.FirstClassProperties);

                    propset.Add(EWSProperties.PR_RULE_MSG_STATE);

                    item.Load(propset);

                    item.TryGetProperty(EWSProperties.PR_RULE_MSG_STATE, out int state);
                    ruleCount++;
                    if (state == 1)
                    {
                        enableCount++;
                    }
                    else
                    {
                        disabledCount++;
                    }

                }

            }
            return Tuple.Create(ruleCount, enableCount, disabledCount);
        }

        private ExchangeService GetEWSService(string username,string password,string email)
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            //service.UseDefaultCredentials = true;
            service.Credentials = new WebCredentials(username, password);
            service.AutodiscoverUrl(email, RedirectionUrlValidationCallback);
           // service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;
            
            return service;
        }

        private PropertySet GetPropertySet()
        {
            PropertySet propset;
            propset = new PropertySet(BasePropertySet.FirstClassProperties);
            propset.Add(EWSProperties.Pr_Folder_Path);
            propset.Add(EWSProperties.HasRules);
            propset.Add(EWSProperties.PR_FolderSize);
            propset.Add(FolderSchema.EffectiveRights);
            propset.Add(FolderSchema.ChildFolderCount);

            return propset;
        }

        private Tuple<int, DateTime?, string, DateTime?, int, string, string> GetMessages(Folder folder)
        {

            int messageSize = 0;
            int count = 0;
            DateTime? datetimeReceived = null;

            string lastModName = "";
            string subject = "";
            var from = "";
            DateTime? lastModTime = null;
            var headers = Enumerable.Empty<dynamic>().Select(e => new { Name = "", Value = "" });
            var iv = new ItemView(1000);
            var ps = new PropertySet(BasePropertySet.FirstClassProperties);
            ps.Add(EmailMessageSchema.From);
            iv.OrderBy.Add(EWSProperties.PR_Last_Modification_Time, SortDirection.Descending);
            FindItemsResults<Item> f;

            f = folder.FindItems(iv);

            if (f.Count() > 0)
            {
                var latest = f.MaxBy(i => i.LastModifiedTime);
                datetimeReceived = latest.DateTimeReceived;
                lastModName = latest.LastModifiedName;
                lastModTime = latest.LastModifiedTime;
                subject = latest.Subject;

                latest.TryGetProperty(EmailMessageSchema.From, out EmailAddress emailAddress);
                from = emailAddress?.Address;
            }
            //var moreAvailable = f.MoreAvailable;
            //while (count < f.TotalCount )
            //{

            //    if (f.Count() > 0)
            //    {
            //        count += f.Count();
            //        messageSize += f.Sum(i => i.Size);
            //        messageSize += f.Sum(i => i.Attachments.Sum(s => s.Size));
            //    }
            //    if (f.MoreAvailable)
            //    {
            //        f = folder.FindItems(new ItemView(1000, f.NextPageOffset.GetValueOrDefault()));
            //    }
            //    moreAvailable = f.MoreAvailable;
            //}


            return Tuple.Create(messageSize, datetimeReceived, lastModName, lastModTime, count, subject, from);

        }
        private IEnumerable<Item> GetAllMessages(Folder folder)
        {

            DateTime? datetimeReceived = null;
            var items = new List<Item>();
            string lastModName = "";
            string subject = "";
            var from = "";
            DateTime? lastModTime = null;
            var headers = Enumerable.Empty<dynamic>().Select(e => new { Name = "", Value = "" });
            var iv = new ItemView(1000);
            var ps = new PropertySet(BasePropertySet.FirstClassProperties);
            ps.Add(EmailMessageSchema.From);
            iv.OrderBy.Add(EWSProperties.PR_Last_Modification_Time, SortDirection.Descending);
            FindItemsResults<Item> f;

            f = folder.FindItems(iv);

            if (f.Count() > 0)
            {
                var latest = f.MaxBy(i => i.LastModifiedTime);
                datetimeReceived = latest.DateTimeReceived;
                lastModName = latest.LastModifiedName;
                lastModTime = latest.LastModifiedTime;
                subject = latest.Subject;

                latest.TryGetProperty(EmailMessageSchema.From, out EmailAddress emailAddress);
                from = emailAddress?.Address;
            }
            var moreAvailable = f.MoreAvailable;
            while (moreAvailable )
            {
                moreAvailable = f.MoreAvailable;
                items.AddRange(folder.FindItems(new ItemView(1000, f.NextPageOffset.GetValueOrDefault())));
            }
            return f;
            
            

        }
        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

       

        private bool CertificateValidationCallBack(
                                                        object sender,
                                                        X509Certificate certificate,
                                                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                                                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }


        private void SplitFolder(string folder, out string first, out string rest)
        {
            if (string.IsNullOrEmpty(folder))
            {
                first = folder;
                rest = null;
                return;
            }

            if (!folder.Contains("\\"))
            {
                first = folder;
                rest = null;
                return;
            }

            first = folder.Substring(0, folder.IndexOf("\\"));
            rest = folder.Substring(folder.IndexOf("\\") + 1, folder.Length - folder.IndexOf("\\") - 1);
        }


    }



}
