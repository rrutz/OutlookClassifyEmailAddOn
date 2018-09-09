using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.IO;
using CsvHelper;
using System.Threading;


namespace Outlook_ClassifyEmailAddOn
{
    public partial class ThisAddIn
    {
        Outlook.Explorer thisExplorer;
        private string predictScriptPath = System.IO.Directory.GetCurrentDirectory() + @"\ML\predict.r";
        private string RscriptExe = @"C:\Program Files\Microsoft\R Client\R_SERVER\bin\Rscript.exe";
        public string dataFile = System.IO.Directory.GetCurrentDirectory() + @"\data\data2.csv";
        private string predictionFile = System.IO.Directory.GetCurrentDirectory() + @"\signalFiles\Prediction2.txt";
        private string signalFile = System.IO.Directory.GetCurrentDirectory() + @"\signalFiles\signal2.csv";
        private string killRscript = System.IO.Directory.GetCurrentDirectory() + @"\signalFiles\killR";
        private string modelsDirectory = System.IO.Directory.GetCurrentDirectory() + "\\ML";
        private string workingDirectory = System.IO.Directory.GetCurrentDirectory();


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(killRscript));
            Directory.CreateDirectory(Path.GetDirectoryName(dataFile));

            // File.WriteAllText(killRscript, " kill the rscript");


            Directory.CreateDirectory(Path.GetDirectoryName(predictScriptPath));
            var files = new List<string>() { signalFile, killRscript, predictionFile };
            foreach (var filePath in files)
            {
                if (File.Exists(filePath))
                {
                    while (IsFileLocked(filePath)) { System.Threading.Thread.Sleep(1); }
                    File.Delete(filePath);
                }
            }

            thisExplorer = this.Application.ActiveExplorer();
            thisExplorer.SelectionChange += new Microsoft.Office.Interop.Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Access_All_Form_Regions);
            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);
            Thread th = new Thread(StartPredictionScript);
            th.Start();
        }

        private void ThisAddIn_Quit()
        {
            File.Create(killRscript);
            var files = new List<string>() { signalFile, predictionFile };
            foreach (var filePath in files)
            {
                if (File.Exists(filePath))
                {
                    while (IsFileLocked(killRscript)) { System.Threading.Thread.Sleep(1); }
                    File.Delete(filePath);
                }
            }
        }

        private void StartPredictionScript()
        {
            var info = new ProcessStartInfo();
            info.Arguments = predictScriptPath + " " + signalFile + " " + predictionFile + " " + killRscript + " " + modelsDirectory + " " + workingDirectory;
            info.FileName = this.RscriptExe;
            info.WorkingDirectory = Path.GetDirectoryName(predictScriptPath);
            info.RedirectStandardInput = false;
            info.RedirectStandardOutput = true;
            info.UseShellExecute = false;
            info.CreateNoWindow = true;

            using (var proc = new Process())
            {
                proc.StartInfo = info;
                proc.Start();
            }
        }

        public void writeToFile(string text)
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    String from = mailItem.SenderName;
                    String to = mailItem.To;
                    String cc = mailItem.CC;
                    String subject = mailItem.Subject;
                    String body = mailItem.Body;

                    using (TextWriter writer = new StreamWriter(dataFile, append: true))
                    {
                        var csv = new CsvWriter(writer);
                        var list = new List<string[]>
                        {
                            new[] { text, from, to, cc, subject, body },
                        };
                        foreach (var item in list)
                        {
                            foreach (var field in item)
                            {
                                csv.WriteField(field);
                            }
                            csv.NextRecord();
                        }
                        writer.Flush();
                    }
                }
            }
        }

        private bool IsFileLocked(string filePath)
        {
            try
            {
                string prediction = File.ReadAllText(filePath);
            }
            catch (IOException)
            {
                return true;
            }
            return false;
        }


        // pulls email information and thne calls R script which returns the predicition
        public string classifyEmail()
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    String from = mailItem.SenderName;
                    String to = mailItem.To;
                    String cc = mailItem.CC;
                    String subject = mailItem.Subject;
                    String body = mailItem.Body;

                    using (TextWriter writer = new StreamWriter(signalFile, append: true))
                    {
                        var csv = new CsvWriter(writer);
                        var list = new List<string[]>
                        {
                            new[] { from, to, cc, subject, body },
                        };
                        foreach (var item in list)
                        {
                            foreach (var field in item)
                            {
                                csv.WriteField(field);
                            }
                            csv.NextRecord();
                        }
                        writer.Flush();
                    }
                }

                while (!File.Exists(predictionFile)) { System.Threading.Thread.Sleep(1); }
                while (IsFileLocked(predictionFile)) { System.Threading.Thread.Sleep(1); }

                string prediction = File.ReadAllText(predictionFile);
                prediction = prediction.Replace("\r\n", "");

                File.Delete(predictionFile);
                File.Delete(signalFile);
                return prediction;
            }
            return "error";
        }

        private void Access_All_Form_Regions()
        {
            foreach (Microsoft.Office.Tools.Outlook.IFormRegion formRegion in Globals.FormRegions)
            {
                if (formRegion is CategoryFormRegion)
                {
                    CategoryFormRegion formRegion1 = (CategoryFormRegion)formRegion;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        #endregion
    }


}

