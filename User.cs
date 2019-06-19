using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Configuration;

namespace Generate_Outlook_Signatures
{
    class User
    {
        public static string MasterList = ConfigurationManager.AppSettings.Get("MasterListLocation");
        public static string Username = Environment.UserName;
        public static string UserFileLocation = (MasterList + @"\User List\" + Username);
        public static string AppData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        public static string OutlookSignatureLocation = (AppData + @"\Microsoft\Signatures");


        public static void RetrieveSignatures()
        {
            foreach (var file in Directory.GetFiles(UserFileLocation))
            {
                File.Copy(file, Path.Combine(OutlookSignatureLocation, Path.GetFileName(file)), true);
            }
            foreach (var folder in Directory.GetDirectories(UserFileLocation))
            {
                string FolderName = new DirectoryInfo(folder).Name;
                DirCopy.DirectoryCopy(folder, OutlookSignatureLocation + @"\" + FolderName, true);
            }

        }
        public static void RetrieveSignaturesLocal()
        {
            foreach (var file in Directory.GetFiles(Signatures.TempLocation))
            {
                File.Copy(file, Path.Combine(OutlookSignatureLocation, Path.GetFileName(file)), true);
            }
            foreach (var folder in Directory.GetDirectories(Signatures.TempLocation))
            {
                string FolderName = new DirectoryInfo(folder).Name;
                DirCopy.DirectoryCopy(folder, OutlookSignatureLocation + @"\" + FolderName, true);
            }

        }
        public static void TestMasterFileLocation()
        {
            if (Directory.Exists(User.UserFileLocation) == true)
            {
                return;
            }
            else
            {
                Console.WriteLine("Error. Could not find " + User.Username + @"'s signature files at the current Master File Location.");
                Console.WriteLine("Please ensure that the Master File variable is set properly, and if so, be sure that all user signatures are currently available, or that " + User.Username + " has read access to the Master File Location.");
                Console.WriteLine("Please contact your IT Systems Administrator, and provide them with this Error Message");
                Thread.Sleep(30000);
                Environment.Exit(2);
            }
        }
        public static void TestOutlookSignatureLocation()
        {
            if (Directory.Exists(User.OutlookSignatureLocation) == true)
            {
                return;
            }
            else
            {
                Console.WriteLine("Error. Could not locate Outlook Signatures Folder. Please ensure that Outlook is installed on this system.");
                Console.WriteLine("Please contact your IT Systems Administrator, and provide them with this Error Message");
                Thread.Sleep(30000);
                Environment.Exit(2);
            }
        }

    }
}

