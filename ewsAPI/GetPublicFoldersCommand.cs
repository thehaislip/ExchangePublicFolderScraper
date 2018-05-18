using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace ewsAPI
{
    [Cmdlet(VerbsCommon.Get,"PublicFolders")]
    class GetPublicFoldersCommand :Cmdlet
    {
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
        }
        protected override void ProcessRecord()
        {
            base.ProcessRecord();
        }
        protected override void EndProcessing()
        {
            //WriteObject(new PublicFolder().GetAllFolders().Distinct(), true);
        }
    }
}
