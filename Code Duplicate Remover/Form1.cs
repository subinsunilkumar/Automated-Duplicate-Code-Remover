using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Execute_Selected_Tests
{
    public partial class Form1 : Form
    {
        private int counter01;
        private bool writeAll;
        private string lastLine;
        string tempText;
        private string functionDeclaration;
        private int firstLineActual;
        private int lastLineActual;
        private List<string> functionCall;
        public string directory02 = @"D:\SVN\ComTests\Develop\ComTests";
        public string backupDir = Path.Combine(@"D:\SVN\ComTests\Develop\ComTests", "CommonTests", "TEMP");
        public int totalRun;
        public int startIndex;
        public List<Tuple<string, string>> fileList = new List<Tuple<string, string>>
        {
            Tuple.Create("HardwareConfiguration.Tests", "HardwareConfiguration.DinDout"),
            //Tuple.Create("TD_Traffic.Tests", "TD_TrafficTests.Traffic"),
            Tuple.Create("SubConfiguration.Tests", "SubConfiguration.SubConfig"),
        };
        public List<Tuple<string, string>> backupList;

        public void BackupFiles()
        {
            var files = Directory.GetFiles(backupDir);
            var reportFile = Path.Combine(directory02, "bin", "x64", "CodeDuplicatesReport.xml");
            var xmlDocument = new XmlDocument();
            xmlDocument.Load(reportFile);
            var list = new List<Tuple<string, string>>();
            var duplicateNodeList = xmlDocument.SelectNodes("DuplicatesReport/Duplicates/Duplicate");
            Invoke(new Action(() =>
            {
                label2.Text = $"Total Duplicates Removed : 0/{duplicateNodeList.Count}";
                label1.Text = $"Remaining: {duplicateNodeList.Count}";
            }));
            var fragmentNodeList = duplicateNodeList[startIndex].SelectNodes("Fragment");
            foreach (XmlNode fragment in fragmentNodeList)
            {
                var fileName = fragment.SelectSingleNode("FileName");
                var actualPath = Path.Combine(directory02, fileName.InnerText);
                var backupPath = Path.Combine(backupDir, Path.GetFileName(actualPath));
                File.Copy(actualPath, backupPath, true);
                list.Add(Tuple.Create(actualPath, backupPath));
            }

            backupList = list;
        }

        public Form1()
        {
            InitializeComponent();
            totalRun = 0;
            startIndex = 0;
            counter01 = 95;
        }

        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private IntPtr handle;
        private Process process;

        private void Form1_Load(object sender, EventArgs e)
        {
            var reportFile = Path.Combine(directory02, "bin", "x64", "CodeDuplicatesReport.xml");
            var slnPath = Path.Combine(directory02, "CommonTests.sln");
            if (File.Exists(slnPath).Equals(false))
            {
                MessageBox.Show("Please Move the Exe File to ComTests Directory to start execution.");
                Application.Exit();
            }
            else
            {
                if (File.Exists(reportFile).Equals(false))
                {
                    MessageBox.Show(
                        "CodeDuplicatesReport.xml file was not found\nPlease Build the solution and restart the application.");
                    Application.Exit();
                }
                else
                {
                    var xmlDocument = new XmlDocument();
                    xmlDocument.Load(reportFile);
                    var duplicateNodeList = xmlDocument.SelectNodes("DuplicatesReport/Duplicates/Duplicate");
                    label2.Text = $"Total Duplicates Removed : 0/{duplicateNodeList.Count}";
                    label1.Text = $"Remaining: {duplicateNodeList.Count}";
                }
            }

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
        }

        private void RemoveDuplicateCode()
        {


            var reportFile = Path.Combine(directory02, "bin", "x64", "CodeDuplicatesReport.xml");
            var xmlDocument = new XmlDocument();
            xmlDocument.Load(reportFile);
            var duplicateNodeList = xmlDocument.SelectNodes("DuplicatesReport/Duplicates/Duplicate");
            //var comTestsDirectory = @"D:\SVN\ComTests\Develop\ComTests";
            var tempFile = Path.Combine(directory02, "CommonTests", "TEMP", "TempHelper.txt");
            var line02 = string.Empty;
            var finalOutput = string.Empty;
            var counter = 75;
            var firstLine = string.Empty;
            lastLine = string.Empty;
            var index = startIndex;
            var exitCondition = false;
            for (var duplicateNodeIndex = index; duplicateNodeIndex < duplicateNodeList.Count; duplicateNodeIndex++)
            {
                BackupFiles();
                var nameOfHelperFile = string.Empty;
                var fragmentNodeList = duplicateNodeList[duplicateNodeIndex].SelectNodes("Fragment");
                for (var positionIndex = 0; positionIndex < fragmentNodeList.Count; positionIndex++)
                {
                    var fileName = fragmentNodeList[positionIndex].SelectSingleNode("FileName");
                    var text = fileName.InnerText;
                    var projectName = text.Split(new[] { "\\" }, StringSplitOptions.None)[1];
                    var fragmentFile = Path.Combine(directory02, text);
                    var fragmentFileContent = File.ReadAllLines((fragmentFile));
                    var lineRange = fragmentNodeList[positionIndex].SelectSingleNode("LineRange");

                    foreach (var value in fileList)
                    {
                        if (projectName.Equals(value.Item1))
                        {
                            nameOfHelperFile = value.Item2;
                        }
                    }

                    if (nameOfHelperFile == string.Empty)
                    {
                        //MessageBox.Show($"This fragment cannot be removed for {projectName}");
                        goto Final;
                    }
                    var processes= Process.GetProcessesByName("devenv");
                    if (processes.Length>1)
                    {
                        MessageBox.Show("More Than One Visual Studio instances are running.\nPlease close them before starting.");
                    }
                    process = processes[0];
                    handle = process.MainWindowHandle;
                    SetForegroundWindow(handle);
                    Thread.Sleep(3000);
                    var attributes = lineRange.Attributes;

                    var start = attributes[0];

                    var end = attributes[1];
                    var originalFileContent = fragmentFileContent;
                    //tempText = string.Empty;

                    lastLineActual = Convert.ToInt16(end.Value);

                    firstLineActual = Convert.ToInt16(start.Value);
                    if (positionIndex == 0)
                    {
                        writeAll = true;
                        /*for (var index = Convert.ToInt16(start.Value); index <= Convert.ToInt16(end.Value); index++)
                        {
                            tempText = tempText + Environment.NewLine + fragmentFileContent[index + 1];
                            //fragmentFileContent[index + 1] = string.Empty;
                        }#1#

                        //Invoke(new Action(() => { richTextBox1.Text = tempText;}));
                        // File.WriteAllLines(fragmentFile,fragmentFileContent);
                        /*Close() AllowDrop documents Alt+W+L#1#*/
                        //SendKeys.SendWait("%(w)(l)");
                        Thread.Sleep(3000);
                        SendKeys.SendWait("^(,)");
                        Thread.Sleep(2000);
                        SendKeys.SendWait(Path.GetFileName(fragmentFile));
                        Thread.Sleep(2000);
                        SendKeys.SendWait("{ENTER}");
                        Thread.Sleep(4000);
                        SendKeys.SendWait("^(g)");
                        Thread.Sleep(2000);
                        SendKeys.SendWait(start.Value);
                        Thread.Sleep(1000);
                        SendKeys.SendWait("{ENTER}");
                        var max = Convert.ToInt16(end.Value) - Convert.ToInt16(start.Value);

                        for (int i = 1; i <= max; i++)
                        {
                            SendKeys.SendWait("+{DOWN}");
                            Thread.Sleep(200);
                        }

                        //Sends Shift + Down
                        SendKeys.SendWait("+{END}");
                        Thread.Sleep(1000);
                        SendKeys.SendWait("^(r)");
                        SendKeys.SendWait("^(m)");
                        Thread.Sleep(2000);
                        SendKeys.SendWait("^(a)");
                        var name = projectName.Replace(".", string.Empty);
                        var functionName = $"Helper{name}Im45{counter01}";
                        counter++;
                        SendKeys.SendWait(functionName);
                        Thread.Sleep(3000);
                        //Enter a text example=>SendKeys.SendWait("ENTER");
                        SendKeys.SendWait("{ENTER}");
                        Thread.Sleep(7000);
                        SendKeys.SendWait("^(s)");
                        Thread.Sleep(2000);
                        SendKeys.SendWait("%(w)(l)");
                        Thread.Sleep(3000);
                        var modifiedFileContent01 = File.ReadAllLines(fragmentFile);
                        var lineOfMethod = 0;
                        functionDeclaration = string.Empty;
                        firstLine = string.Empty;
                        var indexOfBrace = 0;
                        for (var i = Convert.ToInt16(start.Value) - 10; i < modifiedFileContent01.Length; i++)
                        {
                            if (modifiedFileContent01[i].Contains($".{functionName}"))
                            {
                                functionDeclaration = modifiedFileContent01[i];
                            }

                            if (modifiedFileContent01[i].Contains($"private static"))
                            {
                                lineOfMethod = i + 1;
                                modifiedFileContent01[i] = modifiedFileContent01[i].Replace("private", "public");
                                File.WriteAllLines(fragmentFile, modifiedFileContent01);
                                indexOfBrace = modifiedFileContent01[i].IndexOf("(");
                            }
                        }

                        Thread.Sleep(2000);
                        SendKeys.SendWait("^(,)");
                        Thread.Sleep(2000);
                        SendKeys.SendWait(Path.GetFileName(fragmentFile));
                        Thread.Sleep(1000);
                        SendKeys.SendWait("{ENTER}");
                        Thread.Sleep(3000);
                        SendKeys.SendWait("^(g)");
                        Thread.Sleep(1000);
                        SendKeys.SendWait($"{lineOfMethod}");
                        Thread.Sleep(1000);
                        SendKeys.SendWait("{ENTER}");
                        Thread.Sleep(2000);
                        for (int i = 1; i <= indexOfBrace - 10; i++)
                        {
                            SendKeys.SendWait("{RIGHT}");
                            Thread.Sleep(100);
                        }

                        SendKeys.SendWait("^(r)");
                        SendKeys.SendWait("^(o)");
                        Thread.Sleep(1000);
                        SendKeys.SendWait("^(a)");
                        Thread.Sleep(1000);
                        SendKeys.SendWait(nameOfHelperFile);
                        Thread.Sleep(1000);
                        SendKeys.SendWait("{ENTER}");
                        Thread.Sleep(10000);
                        SendKeys.SendWait("^+(s)");
                        counter01++;
                        Thread.Sleep(2000);
                        //SendKeys.SendWait("%(w)(l)");
                        lastLine = string.Empty;
                        Thread.Sleep(2000);
                        var newModifiedContent = File.ReadAllLines(fragmentFile);

                        var originalFile = originalFileContent;
                        var newModifiedFile = newModifiedContent;
                        var maxCount = 0;
                        if (newModifiedContent.Length > originalFileContent.Length)
                        {
                            maxCount = originalFileContent.Length;
                        }
                        else
                        {
                            maxCount = newModifiedContent.Length;
                        }

                        ;
                        List<string> NewLines = new List<string>();
                        for (int lineNo = 0; lineNo < maxCount - 1; lineNo++)
                        {
                            if (!String.IsNullOrEmpty(originalFile[lineNo]) &&
                                !String.IsNullOrEmpty(newModifiedFile[lineNo]))
                            {
                                if (String.Compare(originalFile[lineNo], newModifiedFile[lineNo]) != 0)
                                {
                                    NewLines.Add(newModifiedFile[lineNo]);
                                }
                            }
                            else if (!String.IsNullOrEmpty(originalFile[lineNo]))
                            {
                            }
                            else
                            {
                                NewLines.Add(newModifiedFile[lineNo]);
                            }
                        }

                        var callFirstLineValueCount = 0;
                        var callfirstLineValue = string.Empty;
                        functionCall = new List<string>();
                        var addFunctionCall = true;
                        var lineJustBeforeastLinetoReplace = string.Empty;
                        var firstLineToReplace = string.Empty;
                        var tempCounter03 = 0;
                        foreach (var value in NewLines)
                        {
                            var tempTxt = value.Replace(" ", string.Empty);

                            if (tempTxt.Length > 3 && callFirstLineValueCount == 0)
                            {
                                callfirstLineValue = value;
                                callFirstLineValueCount++;
                            }

                            if (addFunctionCall == false && tempTxt.Length > 0 && tempCounter03 == 0)
                            {
                                lineJustBeforeastLinetoReplace = value;
                                tempCounter03++;
                            }

                            if (callfirstLineValue != string.Empty && addFunctionCall == true)
                            {
                                functionCall.Add(value);
                            }

                            if (value.Contains(functionName) == true)
                            {
                                addFunctionCall = false;
                            }
                        }

                        var tempCounte04 = 0;
                        for (var lineNo = 0; lineNo < newModifiedFile.Length; lineNo++)
                        {
                            if (newModifiedFile[lineNo].Contains(callfirstLineValue) == true && tempCounte04 == 0)
                            {
                                firstLineToReplace = originalFile[lineNo];
                                tempCounte04++;
                            }
                        }

                        firstLine = firstLineToReplace;
                        lastLine = lineJustBeforeastLinetoReplace;
                        firstLine = firstLine.Replace(" ", string.Empty);
                        lastLine = lastLine.Replace(" ", string.Empty);

                        var finalizedContent01 = CheckOriginalAndNewContent(fragmentFileContent, firstLine, lastLine);
                        File.WriteAllLines(fragmentFile, finalizedContent01);
                        var temp01Txt = File.ReadAllText(fragmentFile);
                        if (Regex.Matches(temp01Txt, "{").Count != Regex.Matches(temp01Txt, "}").Count)
                        {
                            File.WriteAllLines(fragmentFile, fragmentFileContent);
                            writeAll = false;
                        }
                        if (Regex.Matches(temp01Txt, "{").Count < 3)
                        {
                            File.WriteAllLines(fragmentFile, fragmentFileContent);
                            writeAll = false;
                        }
                    } //Index = 0 
                    else
                    {
                        if (writeAll==true)
                        {
                            var content = File.ReadAllLines(fragmentFile);
                            var finalizedContent01 = CheckOriginalAndNewContentForOtherFragments(content, firstLine, lastLine);
                            File.WriteAllLines(fragmentFile, finalizedContent01);

                            var tempTxt = File.ReadAllText(fragmentFile);
                            if (Regex.Matches(tempTxt,"{").Count!= Regex.Matches(tempTxt,"}").Count)
                            {
                                File.WriteAllLines(fragmentFile, fragmentFileContent);
                            }
                            if (Regex.Matches(tempTxt, "{").Count <3)
                            {
                                File.WriteAllLines(fragmentFile, fragmentFileContent);
                            }
                        }

                    }

                    
                    exitCondition = true;
                }
                totalRun++;
                
                Invoke(new Action(() =>
                {
                    label2.Text = $"Total Duplicates Removed : {totalRun}/{duplicateNodeList.Count}";
                }));
                Invoke(new Action(() =>
                {
                    label1.Text = $"Remaining: {duplicateNodeList.Count - totalRun}";
                }));
            Final:
                startIndex++;
                if (exitCondition == true)
                {
                    process = Process.GetCurrentProcess();
                    handle = process.MainWindowHandle;
                    SetForegroundWindow(handle);
                    break;
                }
            }


        }

        private string[] CheckOriginalAndNewContent(string[] originalContent, string firstLineTxt, string lastLineTxt)
        {
            var tempStr = string.Empty;
            var startLineIndex = 0;
            var finalLineIndex = 0;
            var backupVarToUse = originalContent;
            for (var i = firstLineActual - 2; i <= lastLineActual + 2; i++)
            {
                tempStr = originalContent[i].Replace(" ", string.Empty);
                if (tempStr.Contains(firstLineTxt))
                {
                    startLineIndex = i;
                }

                if (tempStr.Contains(lastLineTxt))
                {
                    finalLineIndex = i - 1;
                }
            }

            var tem01 = true;
            for (var i = finalLineIndex; i >= 0; i--)
            {
                var tempTxt = backupVarToUse[i].Replace(" ", string.Empty);
                if (tempTxt.Length > 0 && tem01 == true)
                {
                    lastLine = originalContent[i];
                    tem01 = false;
                }
            }

            for (var i = startLineIndex; i <= finalLineIndex; i++)
            {
                backupVarToUse[i] = string.Empty;
            }

            var functionCallIndex = 0;
            for (var i = startLineIndex; i <= finalLineIndex; i++)
            {
                backupVarToUse[i] = functionCall[functionCallIndex];
                functionCallIndex++;
                if (functionCallIndex == functionCall.Count)
                {
                    i = finalLineIndex + 2;
                    break;
                }
            }



            if (startLineIndex != 0 && finalLineIndex != 0 && functionCall.Count > 0)
            {
                return backupVarToUse;
            }
            else
            {
                return originalContent;
            }
        }

        private string[] CheckOriginalAndNewContentForOtherFragments(string[] originalContent, string firstLineTxt, string lastLineTxt)
        {
            var tempStr = string.Empty;
            var startLineIndex = 0;
            var finalLineIndex = 0;
            var backupVarToUse = originalContent;
            for (var i = firstLineActual - 2; i <= lastLineActual + 2; i++)
            {
                tempStr = originalContent[i].Replace(" ", string.Empty);
                if (tempStr.Contains(firstLineTxt))
                {
                    startLineIndex = i;
                }

                if (tempStr.Contains(lastLineTxt.Replace(" ", string.Empty)))
                {
                    finalLineIndex = i;
                }
            }

            for (var i = startLineIndex; i <= finalLineIndex; i++)
            {
                backupVarToUse[i] = string.Empty;
            }

            var functionCallIndex = 0;
            for (var i = startLineIndex; i <= finalLineIndex; i++)
            {
                backupVarToUse[i] = functionCall[functionCallIndex];
                functionCallIndex++;
                if (functionCallIndex == functionCall.Count)
                {
                    i = finalLineIndex + 2;
                    break;
                }
            }

            if (startLineIndex != 0 && finalLineIndex != 0 && functionCall.Count > 0)
            {
                return backupVarToUse;
            }
            else
            {
                return originalContent;
            }
        }



        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button33_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button55_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
        private Point lastPoint;
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            lastPoint = new Point(e.X, e.Y);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Left += e.X - lastPoint.X;
                this.Top += e.Y - lastPoint.Y;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private BackgroundWorker worker01;
        private void button2_Click(object sender, EventArgs e)
        {
            worker01 = new BackgroundWorker();
            worker01.DoWork += (obj, ea) => RemoveDuplicateCode();
            worker01.RunWorkerAsync();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
        public void RestoreFiles()
        {
            foreach (var value in backupList)
            {
                File.Copy(value.Item2, value.Item1, true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RestoreFiles();
            MessageBox.Show("Revert Completed.");
        }
    }
}