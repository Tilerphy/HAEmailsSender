using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SharePointLibFilesCaller
{
    class Program
    {
        public static List<string> CurrentNeedtoResolve = new List<string>();
        private static List<Process> SubProcesses = new List<Process>();
        private static object locker = new object();
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
                            if (!Directory.Exists("origin-emails"))
                            {
                                Directory.CreateDirectory("origin-emails");
                            }
                            File.WriteAllBytes("origin-emails/" + file.Name, file.OpenBinary());
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
            
        }

        static void ProcessFiles() 
        {
            string workFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "emails");
            if (!Directory.Exists("emails")) 
            {
                Directory.CreateDirectory("emails");
            }
            if (Directory.Exists("origin-emails")) 
            {
                FileInfo[] files = new DirectoryInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "origin-emails")).GetFiles();
                Console.WriteLine(files.Length);
                
                foreach (FileInfo file in files) 
                {
                    int fileIndex = 0;
                    string[] alllines = File.ReadAllLines(file.FullName);
                    string headerLine = alllines[0];
                    IEnumerable<string> dataLines = alllines.Skip(1);
                    int maxlinesNumber =  int.Parse(ConfigurationManager.AppSettings["MaxLineEachCSV"]);
                    CurrentNeedtoResolve.Add(string.Format("{0}/{1}.{2}.csv", workFolder,file.Name,  fileIndex));
                    StreamWriter sw = new StreamWriter(string.Format("{0}/{1}.{2}.csv", workFolder, file.Name, fileIndex));
                    sw.WriteLine(headerLine);
                    int lineNumber = 0;
                    foreach (string line in dataLines) 
                    {
                        lineNumber++;
                        if (lineNumber > maxlinesNumber) 
                        {
                            
                            sw.Close();
                            fileIndex++;
                            CurrentNeedtoResolve.Add(string.Format("{0}/{1}.{2}.csv", workFolder, file.Name, fileIndex));
                            sw = new StreamWriter(string.Format("{0}/{1}.{2}.csv", workFolder, file.Name, fileIndex));
                            sw.WriteLine(headerLine);
                            lineNumber = 0;

                        }
                        sw.WriteLine(line);
                    }
                    sw.Close();
                    File.Delete(file.FullName);


                }
            }

            if (Directory.Exists("emails")) 
            {
                string sender = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["Sender"]);
                List<Task> tasks = new List<Task>();
                foreach (string file in CurrentNeedtoResolve) 
                {
                    Task t = new Task((f)=> 
                    {
                        ProcessStartInfo processStartInfo = new ProcessStartInfo(sender);
                        processStartInfo.Arguments = string.Format("\"{0}\"", file);
                        processStartInfo.CreateNoWindow = true;
                        processStartInfo.RedirectStandardOutput = true;
                        processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        processStartInfo.UseShellExecute = false;
                        Process p = new Process();
                        p.StartInfo = processStartInfo;
                        p.OutputDataReceived += P_OutputDataReceived;
                        SubProcesses.Add(p);
                        p.Start();
                        p.BeginOutputReadLine();
                        p.WaitForExit();
                    }, file);

                    t.Start();
                    tasks.Add(t);

                }
                Task.WaitAll(tasks.ToArray());
                Console.WriteLine("All jobs are finished.");

                lock (locker)
                {
                    using (StreamWriter sw = new StreamWriter("log.log", true))
                    {
                        sw.WriteLine("All jobs are finished.");
                    }
                }

            }
        }

        private static void P_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (e.Data != null) 
            {
                string print = e.Data.Trim();
                if (bool.Parse(ConfigurationManager.AppSettings["DisplaySubProcessLog"]))
                {
                    Console.WriteLine(print);
                }
                lock (locker)
                {
                    using (StreamWriter sw = new StreamWriter("log.log", true))
                    {
                        sw.WriteLine(print);
                    }
                }

            }
            
            
        }
    }
}
