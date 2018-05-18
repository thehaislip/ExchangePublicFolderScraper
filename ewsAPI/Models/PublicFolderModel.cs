using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ewsAPI.Models
{
    public class PublicFolderModel
    {
        public string FolderID { get; set; }
                       
        public string FolderPath { get; set; }
        public int? NumberOfRules { get;  set; }
        public int? NumberOfActieRules { get;  set; }
        public int? NumberOfDisabledRules { get;  set; }
        public string Error { get;  set; }
       // public int? TotalSize { get;  set; }
       // public int? ItemCount { get; set; }
        public DateTime? LatestDateReceived { get;  set; }
        public string LastModifiedName { get;  set; }
        public DateTime? LastModifiedTime { get;  set; }
        public string Subject { get; internal set; }
        public string From { get; internal set; }
        //public long FolderSize { get; internal set; }
    }
}
