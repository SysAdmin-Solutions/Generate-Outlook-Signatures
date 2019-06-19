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
    public class Signatures
    {
        public static string MasterList = (ConfigurationManager.AppSettings.Get("MasterListLocation"));
        public static string TempLocation = @"C:\GOSTemp";
        public static string TemplateLocation { get; set; }
        public static string UserFileLocation { get; set; }
        public static void InitializeSignatures()
        {
            var MasterList = ConfigurationManager.AppSettings.Get("MasterListLocation");
            if (Directory.Exists(MasterList + @"\Templates") == false)
            {
                System.IO.Directory.CreateDirectory(MasterList + "\\Templates");
                Console.WriteLine("Error, The Templates folder was not detected at the Master List Location: " + MasterList);
                Console.WriteLine("The Templates folder has been created for you. Please Put your .Docx templates in the folder, using the proper variable subsititues provided in the readme.txt");
                Console.WriteLine("Press any key to Exit the Program.");
                Console.ReadKey();
                Environment.Exit(1);
            }
            if (Directory.Exists(MasterList + "\\User List") == false)
            {
                System.IO.Directory.CreateDirectory(MasterList + "\\User List");
            }
            TemplateLocation = (MasterList + "\\Templates");
            UserFileLocation = (MasterList + "\\User List");
            foreach (var folder in Directory.GetDirectories(UserFileLocation))
            {
                Directory.Delete(folder, true);
            }

        }
        public static void InitializeSignaturesLocal()
        {
            var MasterList = ConfigurationManager.AppSettings.Get("MasterListLocation");
            if (Directory.Exists(MasterList + @"\Templates") == false)
            {
                try
                {
                    System.IO.Directory.CreateDirectory(MasterList + "\\Templates");
                    Console.WriteLine("Error, The Templates folder was not detected at the Master List Location: " + MasterList);
                    Console.WriteLine("The Templates folder has been created for you. Please Put your .Docx templates in the folder, using the proper variable subsititues provided in the readme.txt");
                }
                catch (Exception e) { Console.WriteLine(e); }
                Thread.Sleep(30000);
                Environment.Exit(1);
            }
            TemplateLocation = (MasterList + @"\Templates");
        }
        public static void CleanTemp()
        {
            Directory.Delete(TempLocation, true);
        }
        public static void CreateSignatureLocal()
        {
            System.IO.Directory.CreateDirectory(TempLocation);
            foreach (var file in Directory.GetFiles(TemplateLocation))
            {
                try
                {
                    File.Copy(file, Path.Combine(TempLocation, Path.GetFileNameWithoutExtension(file)));
                }
                catch (Exception e) { Console.WriteLine(e); }
            }
            foreach (var file in Directory.GetFiles(TempLocation))
            {
                try
                {
                    Document doc = new Document();
                    doc.LoadFromFile(file);
                    if (UserInfo.Name != null)
                    {
                        doc.Replace("[[ADDisplayName]]", UserInfo.Name, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDisplayName]]", "", true, true);
                    }
                    if (UserInfo.Title != null)
                    {
                        doc.Replace("[[ADTitle]]", UserInfo.Title, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADTitle]]", "", true, true);
                    }
                    if (UserInfo.Department != null)
                    {
                        doc.Replace("[[ADDepartment]]", UserInfo.Department, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDepartment]]", "", true, true);
                    }
                    if (UserInfo.Name != null)
                    {
                        doc.Replace("[[ADDisplayName]]", UserInfo.Name, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDisplayName]]", "", true, true);
                    }
                    if (UserInfo.Email != null)
                    {
                        doc.Replace("[[Email]]", UserInfo.Email, true, true);
                    }
                    else
                    {
                        doc.Replace("[[Email]]", "", true, true);
                    }
                    if (UserInfo.PhoneOffice != null)
                    {
                        doc.Replace("[[ADTelePhoneNumber]]", "Office: " + UserInfo.PhoneOffice, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADTelePhoneNumber]]", "", true, true);
                    }
                    if (UserInfo.PhoneDirect != null)
                    {
                        doc.Replace("[[ADDID]]", "Direct: " + UserInfo.PhoneDirect, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDID]]", "", true, true);
                    }
                    if (UserInfo.PhoneCell != null)
                    {
                        doc.Replace("[[ADMobile]]", "Cell: " + UserInfo.PhoneCell, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADMobile]]", "", true, true);
                    }
                    if (UserInfo.Fax != null)
                    {
                        doc.Replace("[[ADFax]]", "Fax: " + UserInfo.Fax, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADFax]]", "", true, true);
                    }

                    doc.SaveToFile(file + ".docx", FileFormat.Docx2013);
                    doc.SaveToFile(file + ".txt", FileFormat.Txt);
                    doc.SaveToFile(file + ".rtf", FileFormat.Rtf);
                    doc.SaveToFile(file + ".htm", FileFormat.Html);
                    string txtfile = (file + ".txt");
                    EncodingOps.ConvertFileEncoding(txtfile, txtfile, Encoding.UTF8, Encoding.Default);
                    File.Delete(file);






                }
                catch (Exception e) { Console.WriteLine(e); };
            }

        }


        public static void CreateSignature(string username)
        {
            System.IO.Directory.CreateDirectory(UserFileLocation + "\\" + username);
            var userfile = (UserFileLocation + "\\" + username);


            foreach (var file in Directory.GetFiles(TemplateLocation))
            {
                var ufile = (userfile + "\\" + file);
                if (File.Exists(ufile) == false)
                {
                    try
                    {
                        File.Copy(file, Path.Combine(userfile, Path.GetFileNameWithoutExtension(file)));
                    }
                    catch (Exception e) { Console.WriteLine(e); }
                }
            }
            foreach (var file in Directory.GetFiles(userfile))
            {
                try
                {
                    Document doc = new Document();
                    doc.LoadFromFile(file);
                    if (UserInfo.Name != null)
                    {
                        doc.Replace("[[ADDisplayName]]", UserInfo.Name, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDisplayName]]", "", true, true);
                    }
                    if (UserInfo.Title != null)
                    {
                        doc.Replace("[[ADTitle]]", UserInfo.Title, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADTitle]]", "", true, true);
                    }
                    if (UserInfo.Department != null)
                    {
                        doc.Replace("[[ADDepartment]]", UserInfo.Department, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDepartment]]", "", true, true);
                    }
                    if (UserInfo.Name != null)
                    {
                        doc.Replace("[[ADDisplayName]]", UserInfo.Name, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDisplayName]]", "", true, true);
                    }
                    if (UserInfo.Email != null)
                    {
                        doc.Replace("[[Email]]", UserInfo.Email, true, true);
                    }
                    else
                    {
                        doc.Replace("[[Email]]", "", true, true);
                    }
                    if (UserInfo.PhoneOffice != null)
                    {
                        doc.Replace("[[ADTelePhoneNumber]]", "Office: " + UserInfo.PhoneOffice, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADTelePhoneNumber]]", "", true, true);
                    }
                    if (UserInfo.PhoneDirect != null)
                    {
                        doc.Replace("[[ADDID]]", "Direct: " + UserInfo.PhoneDirect, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADDID]]", "", true, true);
                    }
                    if (UserInfo.PhoneCell != null)
                    {
                        doc.Replace("[[ADMobile]]", "Cell: " + UserInfo.PhoneCell, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADMobile]]", "", true, true);
                    }
                    if (UserInfo.Fax != null)
                    {
                        doc.Replace("[[ADFax]]", "Fax: " + UserInfo.Fax, true, true);
                    }
                    else
                    {
                        doc.Replace("[[ADFax]]", "", true, true);
                    }

                    doc.SaveToFile(file + ".docx", FileFormat.Docx2013);
                    doc.SaveToFile(file + ".txt", FileFormat.Txt);
                    doc.SaveToFile(file + ".rtf", FileFormat.Rtf);
                    doc.SaveToFile(file + ".htm", FileFormat.Html);
                    string txtfile = (file + ".txt");
                    EncodingOps.ConvertFileEncoding(txtfile, txtfile, Encoding.UTF8, Encoding.Default);
                    File.Delete(file);






                }
                catch (Exception e) { Console.WriteLine(e); };
            }

        }
    }
    class EncodingOps {
        public static void ConvertFileEncoding(String sourcePath, String destPath,
                                           Encoding sourceEncoding, Encoding destEncoding)
        {
            // If the destination's parent doesn't exist, create it.
            String parent = Path.GetDirectoryName(Path.GetFullPath(destPath));
            if (!Directory.Exists(parent))
            {
                Directory.CreateDirectory(parent);
            }

            // If the source and destination encodings are the same, just copy the file.
            if (sourceEncoding == destEncoding)
            {
                File.Copy(sourcePath, destPath, true);
                return;
            }

            // Convert the file.
            String tempName = null;
            try
            {
                tempName = Path.GetTempFileName();
                using (StreamReader sr = new StreamReader(sourcePath, sourceEncoding, false))
                {
                    using (StreamWriter sw = new StreamWriter(tempName, false, destEncoding))
                    {
                        int charsRead;
                        char[] buffer = new char[128 * 1024];
                        while ((charsRead = sr.ReadBlock(buffer, 0, buffer.Length)) > 0)
                        {
                            sw.Write(buffer, 0, charsRead);
                        }
                    }
                }
                File.Delete(destPath);
                File.Move(tempName, destPath);
            }
            finally
            {
                File.Delete(tempName);
            }
        }
    }
}

