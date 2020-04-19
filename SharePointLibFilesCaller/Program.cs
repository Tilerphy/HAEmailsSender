using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointLibFilesCaller
{
    class Program
    {
        static void Main(string[] args)
        {

            string lib = ConfigurationManager.AppSettings["SharePointSite"];
            string relativeFolder = ConfigurationManager.AppSettings["RelativeFolder"];
            if (string.IsNullOrEmpty(lib)) 
            {
                Console.WriteLine("Failed: SharePointSite is not set in .config file.");
                return;
            }
            
            using (SPSite site = new SPSite(lib))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPFolder folder = web.GetFolder(relativeFolder);

                    if (folder != null && folder.Exists)
                    {
                        SPFileCollection files = folder.Files;
                        foreach (SPFile file in files)
                        {
                            if (!Directory.Exists("emails"))
                            {
                                Directory.CreateDirectory("emails");
                            }
                            File.WriteAllBytes("emails/" + file.Name, file.OpenBinary());
                        }
                    }
                    else
                    {
                        Console.WriteLine("Sub Folder not exists.");
                    }
                }
            }
            Console.WriteLine("Download csv files Completed");
            ProcessFiles();
            Console.WriteLine("Completed.");
            Console.ReadLine();
            
        }

        static void ProcessFiles() 
        {
            if (Directory.Exists("emails")) 
            {
                FileInfo[] files = new DirectoryInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "emails")).GetFiles();
                string sender = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["Sender"]);
                foreach (FileInfo file in files) 
                {

                    ProcessStartInfo processStartInfo = new ProcessStartInfo(sender);
                    processStartInfo.Arguments = string.Format("\"{0}\"", file.FullName);
                    processStartInfo.CreateNoWindow = true;
                    Process.Start(processStartInfo);
              
                    
                }
            }
        }
    }
}
