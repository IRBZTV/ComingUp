
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ComingUpRndr
{
    public partial class Form1 : Form
    {
        string _AeProject = null;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(ConfigurationSettings.AppSettings["VideoFootage"].ToString().Trim());
                foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();
            }
            catch
            { }
        }
        protected void GenerateScript(string Date, string Time, int Id, int DisId, bool IsSad)
        {
                    StreamWriter Str = new StreamWriter(ConfigurationSettings.AppSettings["xml"].ToString().Trim());
            try
            {
                //if (IsSad)
                //{
                //    _AeProject = ConfigurationSettings.AppSettings["AeProjectPathSad"].ToString().Trim();
                //}
                //else
                //{
                _AeProject = ConfigurationSettings.AppSettings["AeProjectPath"].ToString().Trim();
                //}
                MyDBTableAdapters.DisplayProgTableAdapter Ta = new MyDBTableAdapters.DisplayProgTableAdapter();
                MyDB.DisplayProgDataTable Dt = Ta.SelectNextProgs(4, Date, Time);
                MyDB.DisplayProgDataTable Dt2 = new MyDB.DisplayProgDataTable();
                richTextBox1.Text += Date + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
                //Video:
                MyDBTableAdapters.MASTER_DATATableAdapter Arch_Ta = new MyDBTableAdapters.MASTER_DATATableAdapter();
                if (Dt.Rows.Count > 0)
                {
                    DateTime NewDateTime = DateConversion.JD2GD("13" + Date);
                    string NDate = DateConversion.GD2JD(NewDateTime.AddDays(1)).Remove(0, 2);
                    Dt2 = Ta.SelectNextProgs(4 - Dt.Rows.Count, NDate, "00:00:00");
                    richTextBox1.Text += "Date:" + NewDateTime + " NdateTxt:" + NDate + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();
                }
                if (Dt.Rows.Count + Dt2.Rows.Count == 4)
                {
                    richTextBox1.Text += "Generate XML" + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();
                    for (int i = 0; i < Dt.Rows.Count; i++)
                    {
                        richTextBox1.Text += "Part Today \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();
                        int Minute = int.Parse(Dt.Rows[i]["Time"].ToString().Substring(3, 2));
                        double Sec = Minute * 60 + int.Parse(Dt.Rows[i]["Time"].ToString().Substring(6, 2));
                        string TextTime = Dt.Rows[i]["Time"].ToString().Substring(3, 5);
                        string FinalMinute = "00";
                        #region RoundMinutes
                        if (Sec <= 150)
                        {
                            // <2.5
                            FinalMinute = "00";
                        }

                        if (Sec >= 151 && Sec <= 450)
                        {
                            //2.5 To 7.5
                            FinalMinute = "05";
                        }

                        if (Sec >= 451 && Sec <= 750)
                        {
                            //7.5 To 12.5
                            FinalMinute = "10";
                        }

                        if (Sec >= 751 && Sec <= 1050)
                        {
                            //12.5 To 17.5
                            FinalMinute = "15";
                        }

                        if (Sec >= 1051 && Sec <= 1350)
                        {
                            //17.5 To 22.5
                            FinalMinute = "20";
                        }

                        if (Sec >= 1351 && Sec <= 1650)
                        {
                            //22.5 To 27.5
                            FinalMinute = "25";
                        }

                        if (Sec >= 1651 && Sec <= 1950)
                        {
                            //27.5 To 32.5
                            FinalMinute = "30";
                        }

                        if (Sec >= 1951 && Sec <= 2250)
                        {
                            //32.5 To 37.5
                            FinalMinute = "35";
                        }

                        if (Sec >= 2251 && Sec <= 2550)
                        {
                            //37.5 To 42.5
                            FinalMinute = "40";
                        }

                        if (Sec >= 2551 && Sec <= 2850)
                        {
                            //42.5 To 47.5
                            FinalMinute = "45";
                        }


                        if (Sec >= 2851 && Sec <= 3150)
                        {
                            //47.5 To 52.5
                            FinalMinute = "50";
                        }

                        if (Sec >= 3151 && Sec <= 3599)
                        {
                            //52.5 To 59.59
                            FinalMinute = "55";
                        }


                        #endregion
                        richTextBox1.Text += "MM:SS " + TextTime + " >> " + FinalMinute + " \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();
                        Str.WriteLine("D" + (i + 1).ToString() + "=[\"" + Dt.Rows[i]["Caption"].ToString().Replace("\r\n", "\\r") + "\",\"" + Dt.Rows[i]["Time"].ToString().Substring(0, 2) + ":" + FinalMinute + "\"]");
                        MyDB.MASTER_DATADataTable Arch_Dt = Arch_Ta.GetData(Dt.Rows[i]["Caption"].ToString().Replace("\r\n", "\\r"));
                       
                        
                        //Arm1Q:
                        string dirArm1 = ConfigurationSettings.AppSettings["Arm1Q"].ToString().Trim()+"\\"+ Dt.Rows[i]["date"].ToString().Replace("\\", "-").Replace("/", "-")+"_"+Dt.Rows[i]["Time"].ToString().Substring(0, 2) + "-" + FinalMinute;
                        if (!Directory.Exists(dirArm1))
                            Directory.CreateDirectory(dirArm1);


                        if(Arch_Dt.Rows.Count>0)
                        {
                            Splitter(Arch_Dt[0]["Video_Path_Hi"].ToString(), Path.GetDirectoryName(Application.ExecutablePath) + "\\Splitted.mp4");
                            Repair(Path.GetDirectoryName(Application.ExecutablePath) + "\\Splitted.mp4", Path.GetDirectoryName(Application.ExecutablePath) + "\\Converted.mp4");
                            File.Copy(Path.GetDirectoryName(Application.ExecutablePath) + "\\Converted.mp4", ConfigurationSettings.AppSettings["VideoFootage"].ToString().Trim() + "\\" + (i + 1).ToString() + ".mp4", true);
                            File.Copy(Path.GetDirectoryName(Application.ExecutablePath) + "\\Converted.mp4", dirArm1 + "\\1.mp4", true);
                        }
                        else
                        {
                            File.Copy(getrandomfile2(ConfigurationSettings.AppSettings["VideoRepository"].ToString().Trim()), ConfigurationSettings.AppSettings["VideoFootage"].ToString().Trim() + "\\" + (i + 1).ToString() + ".mp4", true);
                            File.Copy(getrandomfile2(ConfigurationSettings.AppSettings["VideoRepository"].ToString().Trim()), dirArm1+"\\1.mp4", true);
                        }                        

                        StreamWriter StrArm1 = new StreamWriter(dirArm1+"\\data.xml");
                        StrArm1.WriteLine("D=[\"" + Dt.Rows[i]["Caption"].ToString().Replace("\r\n", "\\r") + "\",\"" + Dt.Rows[i]["duration"].ToString().Substring(0, 2) + "\"]");
                        StrArm1.Close();
                    }
                    for (int p = Dt.Rows.Count; p < 4; p++)
                    {
                        richTextBox1.Text += "Part Tomorrow \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();
                        int Minute = int.Parse(Dt2.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(3, 2));
                        double Sec = Minute * 60 + int.Parse(Dt2.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(6, 2));
                        string TextTime = Dt2.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(3, 5);
                        string FinalMinute = "00";
                        #region RoundMinutes
                        if (Sec <= 150)
                        {
                            // <2.5
                            FinalMinute = "00";
                        }

                        if (Sec >= 151 && Sec <= 450)
                        {
                            //2.5 To 7.5
                            FinalMinute = "05";
                        }

                        if (Sec >= 451 && Sec <= 750)
                        {
                            //7.5 To 12.5
                            FinalMinute = "10";
                        }

                        if (Sec >= 751 && Sec <= 1050)
                        {
                            //12.5 To 17.5
                            FinalMinute = "15";
                        }

                        if (Sec >= 1051 && Sec <= 1350)
                        {
                            //17.5 To 22.5
                            FinalMinute = "20";
                        }

                        if (Sec >= 1351 && Sec <= 1650)
                        {
                            //22.5 To 27.5
                            FinalMinute = "25";
                        }

                        if (Sec >= 1651 && Sec <= 1950)
                        {
                            //27.5 To 32.5
                            FinalMinute = "30";
                        }

                        if (Sec >= 1951 && Sec <= 2250)
                        {
                            //32.5 To 37.5
                            FinalMinute = "35";
                        }

                        if (Sec >= 2251 && Sec <= 2550)
                        {
                            //37.5 To 42.5
                            FinalMinute = "40";
                        }

                        if (Sec >= 2551 && Sec <= 2850)
                        {
                            //42.5 To 47.5
                            FinalMinute = "45";
                        }


                        if (Sec >= 2851 && Sec <= 3150)
                        {
                            //47.5 To 52.5
                            FinalMinute = "50";
                        }

                        if (Sec >= 3151 && Sec <= 3599)
                        {
                            //52.5 To 59.59
                            FinalMinute = "55";
                        }


                        #endregion
                        richTextBox1.Text += "MM:SS " + TextTime + " >> " + FinalMinute + " \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();
                        Str.WriteLine("D" + (p + 1).ToString() + "=[\"" + Dt.Rows[p - Dt.Rows.Count]["Caption"].ToString().Replace("\r\n", "\\r") + "\",\"" + Dt.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(0, 2) + ":" + FinalMinute  + "\"]");
                        MyDB.MASTER_DATADataTable Arch_Dt = Arch_Ta.GetData(Dt.Rows[p - Dt.Rows.Count]["Caption"].ToString().Replace("\r\n", "\\r"));

                        //Arm1Q:
                        string dirArm1 = ConfigurationSettings.AppSettings["Arm1Q"].ToString().Trim() + "\\" + Dt.Rows[p - Dt.Rows.Count]["date"].ToString().Replace("\\", "-").Replace("/", "-") + "_" + Dt.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(0, 2) + "-" + FinalMinute;
                        if (!Directory.Exists(dirArm1))
                            Directory.CreateDirectory(dirArm1);

                        if (Arch_Dt.Rows.Count > 0)
                        {
                            Splitter(Arch_Dt[0]["Video_Path_Hi"].ToString(), Path.GetDirectoryName(Application.ExecutablePath) + "\\Splitted.mp4");
                            Repair(Path.GetDirectoryName(Application.ExecutablePath) + "\\Splitted.mp4", Path.GetDirectoryName(Application.ExecutablePath) + "\\Converted.mp4");
                            File.Copy(Path.GetDirectoryName(Application.ExecutablePath) + "\\Converted.mp4", ConfigurationSettings.AppSettings["VideoFootage"].ToString().Trim() + "\\" + (p + 1).ToString() + ".mp4", true);
                            File.Copy(Path.GetDirectoryName(Application.ExecutablePath) + "\\Converted.mp4", dirArm1 + "\\1.mp4", true);
                        }
                        else
                        {
                            File.Copy(getrandomfile2(ConfigurationSettings.AppSettings["VideoRepository"].ToString().Trim()), ConfigurationSettings.AppSettings["VideoFootage"].ToString().Trim() + "\\" + (p + 1).ToString() + ".mp4", true);
                            File.Copy(getrandomfile2(ConfigurationSettings.AppSettings["VideoRepository"].ToString().Trim()), dirArm1 + "\\1.mp4", true);
                        }
                        StreamWriter StrArm1 = new StreamWriter(dirArm1 + "\\data.xml");
                        StrArm1.WriteLine("D=[\"" + Dt.Rows[p - Dt.Rows.Count]["Caption"].ToString().Replace("\r\n", "\\r") + "\",\"" + Dt.Rows[p - Dt.Rows.Count]["duration"].ToString().Substring(0, 2)+ "\"]");
                        StrArm1.Close();
                    }
                    Str.Close();
                    Render(Id, DisId, false);
                }
                else
                {
                    richTextBox1.Text += "There is no 4 Item after Coming Up in Conductor" + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();
                    MyDBTableAdapters.COMINGUPTableAdapter TaTS = new MyDBTableAdapters.COMINGUPTableAdapter();
                    TaTS.UpdateText("There is no 4 Item", Id);
                }
            }
            catch (Exception Exp)
            {
                Str.Close();
                richTextBox1.Text += Exp.Message + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
                MyDBTableAdapters.COMINGUPTableAdapter TaTS = new MyDBTableAdapters.COMINGUPTableAdapter();
                TaTS.UpdateText(Exp.Message, Id); throw;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;


            button1.ForeColor = Color.White;
            button1.Text = "Started";
            button1.BackColor = Color.Red;


            try
            {
                MyDBTableAdapters.COMINGUPTableAdapter TaTS = new MyDBTableAdapters.COMINGUPTableAdapter();
                MyDB.COMINGUPDataTable DtTS = TaTS.SelectTask();
                label2.Text = DtTS.Count.ToString();
                if (DtTS.Count > 0)
                {
                    MyDBTableAdapters.DisplayProgTableAdapter Ta = new MyDBTableAdapters.DisplayProgTableAdapter();
                    MyDB.DisplayProgDataTable Dt = Ta.SelectProgById(int.Parse(DtTS.Rows[0]["DISPLAYID"].ToString()));
                    if (Dt.Rows.Count > 0)
                    {
                        richTextBox1.Text = "";

                        richTextBox1.Text += "Start: [ " + Dt.Rows[0]["Caption"].ToString() + " ] " + Dt.Rows[0]["Date"].ToString() + " - " + Dt.Rows[0]["Time"].ToString() + " \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();

                        GenerateScript(Dt.Rows[0]["Date"].ToString(), Dt.Rows[0]["Time"].ToString(), int.Parse(DtTS.Rows[0]["ID"].ToString()),
                            int.Parse(Dt.Rows[0]["ID_Display"].ToString()), bool.Parse(DtTS.Rows[0]["SAD"].ToString()));
                    }
                }

            }
            catch (Exception Exp)
            {
                richTextBox1.Text += Exp.Message;
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }

            button1.ForeColor = Color.White;
            button1.Text = "Start";
            button1.BackColor = Color.Navy;
            timer1.Enabled = true;
        }
        protected void Render(int Id, int DisId, bool Sad)
        {
            richTextBox1.Text += "Start Render" + " \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();
            try
            {
                MyDBTableAdapters.DisplayProgTableAdapter Ta = new MyDBTableAdapters.DisplayProgTableAdapter();
                MyDB.DisplayProgDataTable Dt = Ta.SelectProgById(DisId);
                Process proc = new Process();
                proc.StartInfo.FileName = "\"" + ConfigurationSettings.AppSettings["AeRenderPath"].ToString().Trim() + "aerender.exe" + "\"";
                string Comp = ConfigurationSettings.AppSettings["AeComposition"].ToString().Trim();
                string DirPathDest = ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim() + "\\" + Dt.Rows[0]["date"].ToString().Replace("\\", "-").Replace("/", "-") + "\\" + ConfigurationSettings.AppSettings["OutputFolderName"].ToString().Trim();
                if (!Directory.Exists(DirPathDest))
                    Directory.CreateDirectory(DirPathDest);
                string OutFile = ConfigurationSettings.AppSettings["OutputFilePrefix"].ToString().Trim() + "_" + Dt.Rows[0]["CassetteNo"].ToString() + ".mp4";
                proc.StartInfo.Arguments = " -project " + "\"" + _AeProject + "\"" + "   -comp   \"" + Comp + "\" -output " + "\"" + DirPathDest + "\\" + OutFile + "\"";
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.CreateNoWindow = true;
                proc.EnableRaisingEvents = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.Start();
                proc.PriorityClass = ProcessPriorityClass.Normal;
                StreamReader reader = proc.StandardOutput;
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (richTextBox1.Lines.Length > 10)
                    {
                        richTextBox1.Text = "";
                    }
                    richTextBox1.Text += (line) + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();
                }
                proc.Close();
                MyDBTableAdapters.COMINGUPTableAdapter TaTS = new MyDBTableAdapters.COMINGUPTableAdapter();
                TaTS.UpdateTask(Id);
            }
            catch (Exception Exp)
            {
                richTextBox1.Text += Exp.Message;
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }
            richTextBox1.Text += DateTime.Now.ToString() + " \n";
            richTextBox1.Text += "=======Task Finished======" + " \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            string DirPathDest = ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim();
            if (Directory.Exists(DirPathDest))
            {
                richTextBox1.Text += "==== MAP DRIVE CHECKED ====" + " \n";
                richTextBox1.Text += DateTime.Now.ToString() + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                richTextBox1.BackColor = Color.DarkBlue;

                button1_Click(new object(), new EventArgs());
            }
            else
            {
                richTextBox1.Text += "==== MAP DRIVE DISCONNECTED ====" + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                richTextBox1.BackColor = Color.Red;
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult Dr = MessageBox.Show("Are you sure to delete render queue", "Delete Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (Dr == System.Windows.Forms.DialogResult.Yes)
            {
                ComingUpRndr.MyDBTableAdapters.COMINGUPTableAdapter Ta = new MyDBTableAdapters.COMINGUPTableAdapter();
                Ta.DeleteQueue();
                MessageBox.Show("All tasks deleted", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button1_Click(new object(), new EventArgs());
            }
        }
        private string getrandomfile2(string path)
        {
            string file = null;
            if (!string.IsNullOrEmpty(path))
            {
                var extensions = new string[] { ".mp4" };
                try
                {
                    var di = new DirectoryInfo(path);
                    var rgFiles = di.GetFiles("*.*").Where(f => extensions.Contains(f.Extension.ToLower()));
                    Random R = new Random();
                    file = rgFiles.ElementAt(R.Next(0, rgFiles.Count())).FullName;
                }
                catch { }
            }
            return file;
        }
        protected void Splitter(string inFile,string outFile)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = Path.GetDirectoryName(Application.ExecutablePath) + "//mencoder";
            proc.StartInfo.Arguments = " -ss 00:01:15  -endpos 00:00:30 -oac pcm -ovc x264 " + "  \"" + inFile + "\"   -o \"" + outFile+ "\"";
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.EnableRaisingEvents = true;
            proc.Start();
            StreamReader reader = proc.StandardError;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                LogWriter(line);
            }
            proc.Close();
        }
        protected void Repair(string Infile, string OutFile)
        {
            LogWriter("Star Fixing " + Infile);
            Process proc = new Process();
            proc.StartInfo.FileName = Path.GetDirectoryName(Application.ExecutablePath) + "//ffmpeg";
            proc.StartInfo.Arguments = " -i  \""+ Infile + "\"   -y \"" + OutFile + "\"";
            LogWriter(proc.StartInfo.Arguments);
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.EnableRaisingEvents = true;
            proc.Start();
            proc.PriorityClass = ProcessPriorityClass.Normal;
            StreamReader reader = proc.StandardError;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                LogWriter(line);
            }
            proc.Close();
            LogWriter("End Fixing " + Infile);
        }
        protected void LogWriter(string LogText)
        {
            if (richTextBox1.Lines.Length > 8)
            {
                richTextBox1.Text = "";
            }

            richTextBox1.Text += (LogText) + " [ " + DateTime.Now.ToString("hh:mm:ss") + " ] \n";
            richTextBox1.Text += "===================\n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();
        }
    }
}
