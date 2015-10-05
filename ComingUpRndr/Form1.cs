
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
        }
        protected void GenerateScript(string Date, string Time, int Id, int DisId, bool IsSad)
        {
            try
            {
                if (IsSad)
                {
                    _AeProject = ConfigurationSettings.AppSettings["AeProjectPathSad"].ToString().Trim();
                }
                else
                {
                    _AeProject = ConfigurationSettings.AppSettings["AeProjectPath"].ToString().Trim();
                }
                MyDBTableAdapters.DisplayProgTableAdapter Ta = new MyDBTableAdapters.DisplayProgTableAdapter();
                MyDB.DisplayProgDataTable Dt = Ta.SelectNextProgs(5, Date, Time);
                MyDB.DisplayProgDataTable Dt2 = new MyDB.DisplayProgDataTable();

                richTextBox1.Text += Date + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();


                if (Dt.Rows.Count > 0)
                {


                    DateTime NewDateTime = DateConversion.JD2GD("13" + Date);
                    string NDate = DateConversion.GD2JD(NewDateTime.AddDays(1)).Remove(0, 2);
                    Dt2 = Ta.SelectNextProgs(5 - Dt.Rows.Count, NDate, "00:00:00");

                    richTextBox1.Text += "Date:" + NewDateTime + " NdateTxt:" + NDate + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();
                }


                if (Dt.Rows.Count + Dt2.Rows.Count == 5)
                {
                    richTextBox1.Text += "Generate Script" + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();

                    StreamWriter Str = new StreamWriter(Path.GetDirectoryName(Application.ExecutablePath) + "//Scr.jsx");
                    Str.WriteLine("app.open(new File(\"" + _AeProject.Replace("\\", "\\\\") + "\"));  ");
                    Str.WriteLine("function LayerText(tname,text)  ");
                    Str.WriteLine("{  ");
                    Str.WriteLine("for(var i = 1; i <= app.project.numItems; i++) {  ");
                    Str.WriteLine("var B=app.project.item(i);  ");
                    Str.WriteLine("for(var j=1; j <= B.numLayers;j++) {  ");
                    Str.WriteLine("	var L=B.layer(j);  ");
                    Str.WriteLine("	if(L.name==tname) {  ");
                    Str.WriteLine("	L.sourceText.setValue(text);  ");
                    Str.WriteLine("	break;  ");
                    Str.WriteLine("}  ");
                    Str.WriteLine("}  ");
                    Str.WriteLine("}  ");
                    Str.WriteLine("}  ");

                    for (int i = 0; i < Dt.Rows.Count; i++)
                    {
                        richTextBox1.Text += "Part Today \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();


                        Str.WriteLine(" LayerText (\"T" + (i + 1).ToString() + "\",\"" + Dt.Rows[i]["Caption"].ToString().Replace("\r\n", "\\r") + "\");  ");

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


                        Str.WriteLine(" LayerText (\"T" + (i + 1).ToString() + "T1\",\"" + Dt.Rows[i]["Time"].ToString().Substring(0, 1) + "\");  ");
                        Str.WriteLine(" LayerText (\"T" + (i + 1).ToString() + "T2\",\"" + Dt.Rows[i]["Time"].ToString().Substring(1, 1) + "\");  ");
                        Str.WriteLine(" LayerText (\"T" + (i + 1).ToString() + "T3\",\"" + FinalMinute.Substring(0, 1) + "\");  ");
                        Str.WriteLine(" LayerText (\"T" + (i + 1).ToString() + "T4\",\"" + FinalMinute.Substring(1, 1) + "\");  ");
                    }


                    for (int p = Dt.Rows.Count; p < 5; p++)
                    {
                        richTextBox1.Text += "Part Tomorrow \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();


                        Str.WriteLine(" LayerText (\"T" + (p + 1).ToString() + "\",\"" + Dt2.Rows[p - Dt.Rows.Count]["Caption"].ToString().Replace("\r\n", "\\r") + "\");  ");

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


                        Str.WriteLine(" LayerText (\"T" + (p + 1).ToString() + "T1\",\"" + Dt2.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(0, 1) + "\");  ");
                        Str.WriteLine(" LayerText (\"T" + (p + 1).ToString() + "T2\",\"" + Dt2.Rows[p - Dt.Rows.Count]["Time"].ToString().Substring(1, 1) + "\");  ");
                        Str.WriteLine(" LayerText (\"T" + (p + 1).ToString() + "T3\",\"" + FinalMinute.Substring(0, 1) + "\");  ");
                        Str.WriteLine(" LayerText (\"T" + (p + 1).ToString() + "T4\",\"" + FinalMinute.Substring(1, 1) + "\");  ");
                    }

                    Str.WriteLine("app.project.save()");
                    Str.WriteLine("app.quit();");
                    Str.Close();

                    ApplyScript(Id, DisId, IsSad);
                }
                else
                {
                    richTextBox1.Text += "There is no 5 Item after Coming Up in Conductor" + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();

                    MyDBTableAdapters.COMINGUPTableAdapter TaTS = new MyDBTableAdapters.COMINGUPTableAdapter();
                    TaTS.UpdateText("There is no 5 Item", Id);
                }

            }
            catch (Exception Exp)
            {
                richTextBox1.Text += Exp.Message + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();

                MyDBTableAdapters.COMINGUPTableAdapter TaTS = new MyDBTableAdapters.COMINGUPTableAdapter();
                TaTS.UpdateText(Exp.Message, Id); throw;
            }



        }

        protected void ApplyScript(int Id, int DisId, bool Sad)
        {
            richTextBox1.Text += "Apply Script" + " \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();

            Process proc = new Process();

            proc.StartInfo.FileName = "\"" + ConfigurationSettings.AppSettings["AeRenderPath"].ToString().Trim() + "afterfx.com" + "\"";

            string ScriptFile = Path.GetDirectoryName(Application.ExecutablePath) + "\\Scr.jsx";
            proc.StartInfo.Arguments = "  -r  " + "\"" + ScriptFile + "\"";
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
                //if (richTextBox1.Lines.Length > 10)
                //{
                //    richTextBox1.Text = "";
                //}
                richTextBox1.Text += (line) + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();

            }
            proc.Close();
            Render(Id, DisId, Sad);
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
                    //if (richTextBox1.Lines.Length > 10)
                    //{
                    //    richTextBox1.Text = "";
                    //}
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

    }
}
