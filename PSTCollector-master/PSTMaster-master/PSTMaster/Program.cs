using ArgumentsParser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PSTMaster
{
    class Program
    {
        public static void Main(string[] args)
        {
            Arguments CommandLine = new Arguments(args);

            if (args.Length == 0)
            {
                PrintInfo("Run PSTMaster.exe /help to see usage.");
            }

            else if (args.Length == 1 && HelpRequired(args[0]))
            {
                DisplayHelp();
            }

            else if (CommandLine["mode"] == null || CommandLine["jobname"] == null || CommandLine["location"] == null || CommandLine["collectpath"] == null)
            {
                PrintError("Looks like one or more paramters missed.");
                
            }

            else
            {
                string mode = CommandLine["mode"].ToLower();
                string jobname = CommandLine["jobname"].ToLower();
                string location = CommandLine["location"].ToLower();
                string collectpath = CommandLine["collectpath"].ToLower();
                string configpath = string.Format("{0}PSTCollect\\", Path.GetPathRoot(Environment.SystemDirectory));
                string computername = System.Environment.MachineName;
                string outputfilename = string.Format("{0}.{1}.{2}.csv", computername, jobname, mode);
                string logfile = string.Format("{0}.{1}.log", computername, jobname);
                string fullogpath = string.Format("{0}\\{1}", configpath, logfile);
                string excemessage;

                /*dropfile(fullogpath, out excemessage);
                if (excemessage != null)
                {
                    PrintError(excemessage);
                    WriteLog(configpath, logfile, excemessage);
                }*/


                try
                {
                    if (Directory.Exists(configpath))
                    {
                        Console.WriteLine("The config directory:{0} exists already.", configpath);
                        WriteLog(configpath, logfile, string.Format("The config directory:{0} exists already.", configpath));

                    }

                    else
                    {
                        DirectoryInfo di = Directory.CreateDirectory(configpath);
                        Console.WriteLine("The config directory was created successfully at {0}.", Directory.GetCreationTime(configpath));
                        WriteLog(configpath, logfile, string.Format("The config directory was created successfully at {0}.", Directory.GetCreationTime(configpath)));
                    }


                    if (!collectpath.EndsWith("\\"))
                    {
                        collectpath += "\\";
                    }
                    if (Directory.Exists(collectpath))
                    {
                        Console.WriteLine("Checked the collect path {0} is acessable.", collectpath);
                        WriteLog(configpath, logfile, string.Format("The collect path {0} is acessable.", configpath));
                    }
                    else
                    {

                        PrintError("The collect path is unreachable, please make sure put in a accessible share,quitting the application.....");
                        WriteLog(configpath, logfile, "The collect path is unreachable,please make sure put in a accessible share,quitting the application.....");
                        return;

                    }

                }

                catch (Exception e)
                {
                    Console.WriteLine("The process failed: {0}", e.ToString());
                    WriteLog(configpath, logfile, e.Message);
                }               
              

                

                if (mode != null && jobname != null && location != null && collectpath != null)
                {
                    dropfile(fullogpath, out excemessage);
                    WriteLog(configpath,logfile,excemessage);
                    switch (mode)
                    {
                        case "find": DoMode.Find(location, collectpath, jobname, computername, configpath, outputfilename); break;
                        case "collect": DoMode.Collect(location, collectpath, jobname, computername, configpath, outputfilename); break;
                        case "remove": DoMode.Remove(location, collectpath, jobname, computername, configpath, outputfilename); break;
                    }
                }

                else
                {
                    PrintError("Looks like there is a paramter missing");
                    Environment.Exit(1);
                }

                CopyFile(logfile, configpath, collectpath, out excemessage);
                if(excemessage != null){
                WriteLog(configpath, logfile, excemessage); }
            }
        }

        public static class DoMode
        {
            public static void Find(string locationpath, string collectpath, string jobname, string computername, string configpath, string outputfilename)
            {
                string logfile = string.Format("{0}.{1}.log", computername, jobname);
                string message = "Start Find mode....";
                string dropexception, copyexception, searchexception; 
                WriteLog(configpath, logfile, message);
                
                dropfile(outputfilename, out dropexception);
                if (dropexception != null)
                { WriteLog(configpath, logfile, dropexception); };
               
                    List<string> files = SearchFile.Search(locationpath,out searchexception);
                    if(searchexception !=null)
                    { WriteLog(configpath, logfile, searchexception);}
                    using (StreamWriter text = new StreamWriter(configpath + outputfilename))
                    {
                        if ((files != null) && (files.Any()))
                        {
                            text.WriteLine("ComputerName, FilePath");

                            foreach (string f in files)
                            {
                                //Console.WriteLine("File Found:{0}", f);
                                text.WriteLine("{0},\"{1}\"", computername, f);

                            }
                        }

                        else
                        {
                            text.WriteLine("Search operaton completed but no .pst file being found");
                            WriteLog(configpath, logfile, string.Format("Search operaton completed but no .pst file being found"));
                        }
                    }               

              

                CopyFile(outputfilename, configpath, collectpath, out copyexception);
                if (copyexception != null)
                { WriteLog(configpath, logfile, copyexception); }                
                WriteLog(configpath, logfile, "Find operation completed, please check output in the csv file");
            }
            public static void Collect(string locationpath, string collectpath, string jobname, string computername, string configpath, string outputfilename)
            {
                string logfile = string.Format("{0}.{1}.log", computername, jobname);
                string message = "Start Collect mode....";
                string dropexception, copyexception, searchexception; 
                WriteLog(configpath, logfile, message);

                dropfile(outputfilename, out dropexception);
                if (dropexception != null)
                { WriteLog(configpath, logfile, dropexception); }

                //string logfile=string.Format("{0}.{1}.txt", computername, jobname);
                
                    List<string> files = SearchFile.Search(locationpath,out searchexception);                    
                    if(searchexception !=null)
                    { WriteLog(configpath, logfile, searchexception);}

                    using (StreamWriter text = new StreamWriter(configpath + outputfilename))
                    {
                        if ((files != null) && (files.Any()))
                        {
                            text.WriteLine("ComputerName, FilePath, Collected");

                            foreach (string f in files)
                            {
                                string flag = null;
                                FileInfo fileinfo = new FileInfo(f);
                                string filename = fileinfo.Name;
                                string dir = System.IO.Path.GetDirectoryName(f);
                                string partdir = dir.ToString().Replace(":", "");
                                string finalpath = partdir.TrimEnd('\\');
                                string destpath = string.Format("{0}\\{1}\\{2}\\{3}", collectpath, jobname, computername, finalpath);
                                string destFile = System.IO.Path.Combine(destpath, filename);

                                //Console.WriteLine("the partdir is :{0}, the final dir is {1},the destpath is {2}", partdir, finalpath, destpath);

                                CopyFile(filename, dir, destpath, out copyexception);
                                if (copyexception != null)
                                { WriteLog(configpath, logfile, copyexception); }

                                if (File.Exists(destFile))
                                {
                                    Console.WriteLine("The file: {0} successfully been copied to the collect path: {1}", f, collectpath);
                                    WriteLog(configpath, logfile, string.Format("The file: {0} successfully been copied to the collect path: {1}", f, collectpath));

                                    flag = "Success";
                                }

                                else
                                {
                                    flag = "Failed";
                                }

                                text.WriteLine("{0},\"{1}\",{2}", computername, f, flag);
                            }
                        }

                        else
                        {
                            text.WriteLine("Search operaton completed but no .pst file being found");
                            WriteLog(configpath, logfile, string.Format("Search operaton completed but no .pst file being found"));
                        }
                    }
                
                

                CopyFile(outputfilename, configpath, collectpath,out copyexception);              
                WriteLog(configpath, logfile, copyexception);
                WriteLog(configpath, logfile, "Collect operation completed, please check output in the csv file");
            }
            public static void Remove(string locationpath, string collectpath, string jobname, string computername, string configpath, string outputfilename)
            {
                string logfile = string.Format("{0}.{1}.log", computername, jobname);
                string message = "Start Remove mode....";
                string dropexception, copyexception, searchexception; 
                WriteLog(configpath, logfile, message);

                dropfile(outputfilename, out dropexception);
                if (dropexception != null)
                { WriteLog(configpath, logfile, dropexception); }               
                
                List<string> files = SearchFile.Search(locationpath,out searchexception);                    
                if(searchexception !=null)
                    { WriteLog(configpath, logfile, searchexception);}

                    using (StreamWriter text = new StreamWriter(configpath + outputfilename))
                    {
                        if ((files != null) && (files.Any()))
                        {
                            text.WriteLine("ComputerName, FilePath, Removed");

                            foreach (string f in files)

                            {
                                string flag = null;
                                FileInfo fileinfo = new FileInfo(f);
                                string filename = fileinfo.Name;
                                string dir = System.IO.Path.GetDirectoryName(f);
                                string partdir = dir.ToString().Replace(":", "");
                                string finalpath = partdir.TrimEnd('\\').TrimStart('\\');
                                string destpath = string.Format("{0}\\{1}\\{2}\\{3}", collectpath, jobname, computername, finalpath);
                                string destFile = System.IO.Path.Combine(destpath, filename);

                                if (File.Exists(destFile))
                                {
                                    File.Delete(f);

                                    if (!File.Exists(f))
                                    {
                                        Console.WriteLine("The file:{0} has successfully been removed", f); flag = "Success";
                                        WriteLog(configpath, logfile, string.Format("The file:{0} has successfully been removed", f));

                                    }
                                }

                                else
                                {
                                    Console.WriteLine("File:{0} has not been copyied to {1} yet, please run collect mode first", f, destpath); flag = "Failed";
                                    WriteLog(configpath, logfile, string.Format("File:{0} has not been copyied to {1} yet, please run collect mode first", f, destpath));
                                }

                                text.WriteLine("{0},\"{1}\",{2}", computername, f, flag);
                            }
                        }

                        else
                        {
                            text.WriteLine("Search operaton completed but no .pst file being found");
                        }
                    }               

               

                CopyFile(outputfilename, configpath, collectpath, out copyexception);
                if (copyexception != null)
                { WriteLog(configpath, logfile, copyexception); }  

                WriteLog(configpath, logfile, "Remove operation completed, please check output in the csv file");

            }
        }
        static void CopyFile(string fileName, string sourcePath, string targetPath, out string exmessage)
        {
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);
            exmessage = null;

            try
            {
                if (!System.IO.Directory.Exists(targetPath))
                {
                    System.IO.Directory.CreateDirectory(targetPath);
                }


                if (System.IO.Directory.Exists(sourcePath))
                {

                    File.Copy(sourceFile, destFile, true);

                }
            }
            catch (Exception iox)
            {
                PrintError(iox.Message);
                exmessage = iox.Message;
            }
            
        }
        public static class SearchFile
        {
            public static List<string> Search(string Loc, out string searchexception)
            {
                List<string> filegroup = new List<string>();
                searchexception = null;

                if (Loc == "alllocal")
                {
                    DriveInfo[] TotalDrives = DriveInfo.GetDrives();

                    foreach (DriveInfo drvinfo in TotalDrives)
                    {
                        if (drvinfo.IsReady == true)
                        {
                            string driveLetter = drvinfo.Name.ToString();

                            try
                            {
                                filegroup.AddRange(ApplyAllFiles(driveLetter));
                            }

                            catch (Exception e)
                            {
                                PrintError("Exception:" + e.Message);
                                searchexception = e.Message;
                                //WriteLog(configpath, logfile, "Exception in Find mode:" + e.Message);
                            }
                        }
                    }
                }

                else
                {
                    if (!Loc.EndsWith("\\"))
                    {
                        Loc += "\\";

                        try
                        {
                            filegroup = ApplyAllFiles(Loc);
                        }

                        catch (Exception e)
                        {
                            PrintError("Exception" + e.Message);
                            //WriteLog(configpath, logfile, "Exception in Find mode:" + e.Message);
                        }
                    }
                }

                return filegroup;
            }
            static bool isExcluded(List<string> exludedDirList, string target)
            {
                return exludedDirList.Any(d => new DirectoryInfo(target).Name.Contains(d));
            }
            public static List<string> ApplyAllFiles(string folder)
            {
                List<string> excludefolders = new List<string>() { "$Recycle.Bin", "System Volume Information", "OneDrive - Microsoft" };
                List<string> files = new List<string>();

                foreach (string file in Directory.GetFiles(folder, "*.pst"))
                {
                    Console.WriteLine(file);
                    files.Add(file);
                }

                foreach (string subDir in Directory.GetDirectories(folder).Where(d => !isExcluded(excludefolders, d)))
                {
                    try
                    {
                        //List<string> subfiles=ApplyAllFiles(subDir);
                        files.AddRange(ApplyAllFiles(subDir));
                    }

                    catch
                    {
                        //Console.WriteLine("{0}",e.Message);

                        // swallow, log, whatever
                    }
                }

                return files;
            }
        }
        static void PrintInfo(string txt)
        {
            ConsoleColor currentForegroud = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(txt);
            Console.ForegroundColor = currentForegroud;
        }
        static void PrintError(string txt)
        {
            ConsoleColor currentForegroud = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(txt);
            Console.ForegroundColor = currentForegroud;
        }
        static void WriteLog(string logPath, string fileName, string message)
        {
            try
            {
                FileStream objFilestream = new FileStream(string.Format("{0}\\{1}", logPath, fileName), FileMode.Append, FileAccess.Write);
                StreamWriter objStreamWriter = new StreamWriter((Stream)objFilestream);
                message = string.Format("{0}:{1}:{2}", DateTime.Now.ToString(), System.Environment.MachineName, message);

                objStreamWriter.WriteLine(message);

                objStreamWriter.Close();

                objFilestream.Close();
            }

            catch (Exception ex)
            {
                PrintError(ex.Message);
            }
        }
        static void dropfile(string file, out string exmessage)
        {
            exmessage = null;
            if (Directory.Exists(Path.GetDirectoryName(file)))
            {
                try
                {
                    File.Delete(file);
                }

                catch (System.IO.IOException e)
                {
                    PrintError(e.Message);
                    exmessage = e.Message;
                    return;
                }
            }
        }
        private static bool HelpRequired(string param)
        {
            return param == "-help" || param == "/help" || param == "/?";
        }
        private static void DisplayHelp()
        {
            string[] lines = {
            "",
            ".DESCRIPTION",
            "This application supports three feature: Find, Collect and Remove. ",
            "The Find feature will only get a result collection of the .PST files in your specified location.",
            "The Collect feature will get a result collection and copy the PST files to the path you specify, we recommand you NOT use a local path as collect path but use a share instead.",
            "The Remove feature will only delete the PST files which were scuessfully collected to your collect path, so make sure you run remove mode with jobname same as collect mode, also the collectpath should be same with collect mode.",
            "",
            "",
            ".EXAMPLE",
            "PSTMaster.exe -mode Find -jobname myjob -location C: -collectpath C:",
            ".EXAMPLE",
            "PSTMaster.exe -mode Find -jobname myjob -location alllocal -collectpath \\\\SharePath",
            ".EXAMPLE",
            "PSTMaster.exe -mode Collect -jobname myjob -location alllocal -collectpath \\\\SharePath",
            ".EXAMPLE",
            "PSTMaster.exe -mode Remove -jobname myjob -location alllocal -collectpath \\\\SharePath"
            };

            foreach (string line in lines)
            {
                Console.WriteLine(line);
            }
        }
    }
}
