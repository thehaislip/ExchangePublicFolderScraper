using System;
using ewsAPI;
using System.Management.Automation;
using System.Linq;

namespace GetPublicFolderDetails
{
    [Cmdlet(VerbsCommon.Get, "PublicFolders")]
    public class GetPublicFoldersCommand : Cmdlet
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
            WriteObject(new PublicFolder().GetAllFolders().Distinct(), true);
        }
    }
}
