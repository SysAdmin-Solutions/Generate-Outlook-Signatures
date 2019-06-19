using System.Configuration;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using Spire.Doc;
using System.Text;
using System.Threading;
using System.Security.Principal;
using System.Diagnostics;
using System.Reflection;
using System.Windows;

namespace Generate_Outlook_Signatures
{
    class Program
    {
        static void Main(string[] args)
        {
            // mode 1 is Generate for all
            // mode 2 is Generate for signed in and import
            // mode 3 is Recieve for signed in from mode 1 list.

            var Mode = System.Configuration.ConfigurationManager.AppSettings.Get("Mode");
            bool valid = false;

            if (Mode == "0")
            {
                valid = true;
                try
                {
                    InitializeProgram();
                } catch (Exception e)
                {
                    Console.WriteLine("There was an error changing the config file. This may be because this setup process needs to be run as administrator.");
                    Thread.Sleep(5000);
                    Console.WriteLine("Relaunching program as administrator...");
                    Thread.Sleep(2000);
                    AdminRelauncher();

                 }
                Thread.Sleep(5000);
                Environment.Exit(0);

            }

            if (Mode == "1")
            {
                valid = true;
                Signatures.InitializeSignatures();
                ADops.GetAdUserList();
                foreach (string userName in ADops.List)
                {
                    ADops.SetADUserInfo(userName);
                    if (UserInfo.Email != null)
                    {
                        Console.WriteLine("Generating Signatures for " + UserInfo.Name);
                        Signatures.CreateSignature(userName);
                    }
                }
                Console.WriteLine("Operation Complete.");
                Thread.Sleep(10000);
                Environment.Exit(0);
            }

            if (Mode == "2")
            {
                valid = true;
                var currentuser = Environment.UserName;
                try
                {
                    ADops.SetADUserInfo(currentuser);
                    Console.WriteLine("Creating Signatures for " + currentuser);
                    Signatures.InitializeSignaturesLocal();
                    Signatures.CreateSignatureLocal();
                    User.TestOutlookSignatureLocation();
                    Console.WriteLine("Importing Created Signatures...");
                    User.RetrieveSignaturesLocal();
                    Thread.Sleep(2000);
                    Signatures.CleanTemp();
                    

                } catch (Exception e) {
                    Console.WriteLine(e);
                    Thread.Sleep(10000);
                    Environment.Exit(1);
                }
                Console.WriteLine("Signatures have been successfully imported!");
                Thread.Sleep(3000);
                Environment.Exit(0);

                

            }
            if (Mode == "3")
            {
                valid = true;
                User.TestMasterFileLocation();
                User.TestOutlookSignatureLocation();
                Console.WriteLine("Retrieving Signatures for " + User.Username);
                try
                {
                    Thread.Sleep(2000);
                    User.RetrieveSignatures();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    Thread.Sleep(10000);
                    Environment.Exit(1);
                }
                Console.WriteLine("Signatures have been successfully imported!");
                Thread.Sleep(5000);
                Environment.Exit(0);

            }
            if (valid == false)
            {
                Console.WriteLine("Invalid Mode configured. Please set your mode to either \"0\", \"1\", \"2\", or \"3\" ");
                Thread.Sleep(30000);
                Environment.Exit(1);
            }
        }
        public static void InitializeProgram()
        {
            Console.WriteLine("Welcome to the Outlook Signature Generator by Cameron Russell");
            Thread.Sleep(3000);
            Console.WriteLine("It appears that you have not used this program yet. Lets configure the program for your first use.");
            Thread.Sleep(5000);
            Console.Clear();
            Console.WriteLine("It appears that your Master List Location is currently set to " + ConfigurationManager.AppSettings["MasterListLocation"]);
            string choice1 = null;
            Console.WriteLine("Is this correct?  y/n");
            choice1 = Console.ReadLine().ToString().ToLower();
            if (choice1 == "n")
            {
                Console.WriteLine("Please enter your new Master List Location in the form of \\server\\folder\\location  ");
                var input = Console.ReadLine();
                UpdateSetting("MasterListLocation", input);
            }
            string choice2 = null;
            Console.Clear();
            Console.WriteLine("It appears that your Domain Address is currently set to " + ConfigurationManager.AppSettings["DomainAddress"]);
            Console.WriteLine("Is this correct?  y/n");
            choice2 = Console.ReadLine().ToString().ToLower();
            string choice3 = null;
            if (choice2 == "n")
            {
                Console.WriteLine("Please enter your new Domain Address in the form of: OU=ou,DC=contoso,DC=com");
                var input = Console.ReadLine();
                UpdateSetting("DomainAddress", input);
            }
            bool validmode = false;
            do {
                Console.WriteLine("Please Enter which mode you would like to set the program to.");
                Thread.Sleep(2000);
                Console.WriteLine("Please keep in mind, in order to change modes after setting it here, you must change it manually in the .config file located with the .exe file.");
                Thread.Sleep(3000);
                Console.WriteLine("If you would like to run this setup process again, manually change the Mode in the .config file to \"0\" ");
                Thread.Sleep(5000);
                Console.WriteLine("");
                Console.WriteLine("Press \"0\" to run this setup again.");
                Console.WriteLine("Press \"1\" to Enter Server Generate Mode, where signatures are created for everyone with an email in AD, and saved to their userlist folder in the master list location.");
                Console.WriteLine("Press \"2\" to Enter Local Standalone Mode, where signatures are created for the current logged on user, and imported directly to their outlook signatures.");
                Console.WriteLine("Press \"3\" to Enter Recieve Mode, where signatures are retrieved from the master file userlist for the current logged on user. Used to retrieve list from mode 1 programs.");
                Console.WriteLine("For more information, refer to the ReadMe.txt located with the .exe file.");
                choice3 = Console.ReadLine().ToString();
                if (choice3 == "0") { validmode = true; }
                if (choice3 == "1") { validmode = true; }
                if (choice3 == "2") { validmode = true; }
                if (choice3 == "3") { validmode = true; }

            } while (validmode == false);
            UpdateSetting("Mode", choice3);
            Console.WriteLine("Configuration Complete. Rerun the program to start in selected mode.");
            

        }

        public static void UpdateSetting(string key, string value)
        {
            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            configuration.AppSettings.Settings[key].Value = value;
            configuration.Save();
            ConfigurationManager.RefreshSection("appSettings");
        }
        public static bool IsRunAsAdmin()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(id);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        public static void AdminRelauncher()
        {
            if (!IsRunAsAdmin())
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = Assembly.GetEntryAssembly().CodeBase;

                proc.Verb = "runas";

                try
                {
                    Process.Start(proc);
                    Environment.Exit(1);
                }
                catch (Exception e)
                {
                    Console.WriteLine("This program must be run as administrator to set the variables!");
                }
            }
        }

    }
    
    
    
}
