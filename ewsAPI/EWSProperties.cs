using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ewsAPI
{
    public static class EWSProperties
    {

        public static ExtendedPropertyDefinition HasRules
        {
            get
            {
                return new ExtendedPropertyDefinition(0x663A, MapiPropertyType.Boolean);
            }
        }

        public static ExtendedPropertyDefinition PidTagMessageSizeExtended
        {
            get
            {
                return new ExtendedPropertyDefinition(3592, MapiPropertyType.Long);
            }
        }
        public static ExtendedPropertyDefinition PidTagLocalCommitTimeMax
        {
            get
            {
                return new ExtendedPropertyDefinition(0x670A, MapiPropertyType.SystemTime);
            }
        }
        public static ExtendedPropertyDefinition Pr_Folder_Path
        {
            get
            {
                return new ExtendedPropertyDefinition(26293, MapiPropertyType.String);
            }
        }
        public static ExtendedPropertyDefinition PR_Display_name
        {
            get
            {
                return new ExtendedPropertyDefinition(0x3001, MapiPropertyType.String);
            }
        }

        public static ExtendedPropertyDefinition PR_RULE_MSG_STATE
        {
            get {
                return new ExtendedPropertyDefinition(0x65E9, MapiPropertyType.Integer);
            }
        }
        public static ExtendedPropertyDefinition PR_EXTENDED_RULE_ACTIONS
        {
            get
            {
                return new ExtendedPropertyDefinition(0x0E99, MapiPropertyType.Binary);
            }
        }
        public static ExtendedPropertyDefinition PR_EXTENDED_RULE_CONDITION
        {
            get
            {
                return new ExtendedPropertyDefinition(0x0E9A, MapiPropertyType.Binary);
            }
        }
        public static ExtendedPropertyDefinition PR_Last_Modification_Time
        {
            get
            {
                return new ExtendedPropertyDefinition(0x0E9A, MapiPropertyType.Binary);
            }
        }
        public static ExtendedPropertyDefinition PR_FolderSize
        {
            get
            {
                return new ExtendedPropertyDefinition(3592, MapiPropertyType.Long);
            }
        }
        public static ExtendedPropertyDefinition PR_PF_Proxy
        {
            get
            {
                return new ExtendedPropertyDefinition(0x671D, MapiPropertyType.Binary);
            }
        }
        
    }
}
