using System.Configuration;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using Spire.Doc;
using System.Text;
using System.Threading;

namespace Generate_Outlook_Signatures
{
    public class UserInfo
    {
        public static string Name { get; set; }
        public static string UserName { get; set; }
        public static string Title { get; set; }
        public static string Department { get; set; }
        public static string Email { get; set; }
        public static string PhoneOffice { get; set; }
        public static string PhoneDirect { get; set; }
        public static string PhoneCell { get; set; }
        public static string Fax { get; set; }
    }
}
