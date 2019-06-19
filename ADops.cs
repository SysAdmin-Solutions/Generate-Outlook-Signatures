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
    class ADops
    {

        public static string DomainAddress = ("LDAP://" + (ConfigurationManager.AppSettings.Get("DomainAddress")));

        public static Array List { get; set; }

        public static void GetAdUserList()
        {
            DirectoryEntry enTry = new DirectoryEntry(DomainAddress);
            DirectorySearcher mySearcher = new DirectorySearcher(enTry);
            mySearcher.Filter = "(objectClass=user)";
            var userNames = new List<string>();
            string user = null;
            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                user = (resEnt.GetDirectoryEntry().Properties["samaccountname"].Value.ToString());
                userNames.Add(user);
            }
            var list = userNames.ToArray();
            List = list;
        }
        public static string GetAdUserInfo(string user, string property)
        {
            DirectoryEntry enTry = new DirectoryEntry(DomainAddress);
            DirectorySearcher mySearcher = new DirectorySearcher(enTry);
            mySearcher.Filter = "(&(objectClass=user)(anr=" + user + "))";
            string output = null;
            foreach (SearchResult resEnt in mySearcher.FindAll())
            {

                PropertyValueCollection aduser = resEnt.GetDirectoryEntry().Properties[property];
                if ((aduser == null) | (aduser.Value == null))
                {
                    continue;
                }
                if (aduser != null)
                {
                    output = aduser.Value.ToString();
                }



            }

            return output;
        }
        public static void SetADUserInfo(string user)
        {
            string name = GetAdUserInfo(user, "displayname");
            string UserName = GetAdUserInfo(user, "samaccountname");
            string Title = GetAdUserInfo(user, "title");
            string Department = GetAdUserInfo(user, "department");
            string Email = GetAdUserInfo(user, "mail");
            string PhoneOffice = GetAdUserInfo(user, "telephonenumber");
            string PhoneDirect = GetAdUserInfo(user, "ipPhone");
            string PhoneCell = GetAdUserInfo(user, "mobile");
            string Fax = GetAdUserInfo(user, "facsimileTelephoneNumber");

            UserInfo.Name = name;
            UserInfo.UserName = UserName;
            UserInfo.Title = Title;
            UserInfo.Department = Department;
            UserInfo.Email = Email;
            UserInfo.PhoneOffice = PhoneOffice;
            UserInfo.PhoneDirect = PhoneDirect;
            UserInfo.PhoneCell = PhoneCell;
            UserInfo.Fax = Fax;

        }






    }
}
