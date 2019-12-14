using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using KAutoHelper;
namespace WindowsFormsApp2 {
    public partial class Form1 : Form {
        string[] files;
        Process process = new Process();
        Thread newThread;
        public Form1() {
            InitializeComponent();
            button1.Enabled = false;
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
        }
        void runCMD(string cmd) {
            process.Start();
            process.StandardInput.WriteLine(cmd);
            process.StandardInput.Flush();
            process.StandardInput.Close();
            process.WaitForExit();
        }

        void getProfile(string url) {
            files = Directory.GetFiles(url);
            foreach (var item in files) {
                dataGridView1.Rows.Add(item,"waiting");
            }
        }
        bool isRunning = false;
        void runChrome() {
            if (isRunning) {
                for (int i = 1; i <= 9999; i++) {
                    if (isRunning) {
                        for (int rows = 0; rows < dataGridView1.Rows.Count; rows++) {
                            if (isRunning) {
                                IPHostEntry ip;
                                while (isRunning) {
                                    var check = false;
                                    dataGridView1.Rows[rows].Cells[1].Value = "Change IP";
                                    runCMD("rasdial viettel");
                                    string myip = "";
                                    ip = Dns.GetHostEntry(Dns.GetHostName());
                                    foreach (IPAddress ipa in ip.AddressList) {
                                        if (ipa.AddressFamily == AddressFamily.InterNetwork) {
                                            myip = ipa.ToString();
                                        }
                                    }

                                    if (File.Exists("ListIP.txt")) {
                                        dataGridView1.Rows[rows].Cells[1].Value = "Check duplicate IP";
                                        string[] listIP = File.ReadAllLines("ListIP.txt");
                                        foreach (var item in listIP) {
                                            var temp = item.Split(':');
                                            if (myip == temp[0].Trim()) {
                                                check = true;
                                            }
                                        }
                                    }
                                    if (!check) {
                                        File.AppendAllText("ListIP.txt",myip + " : " + dataGridView1.Rows[rows].Cells[0].Value + "\n");
                                        dataGridView1.Rows[rows].Cells[1].Value = "Open Chrome";
                                        Process.Start(dataGridView1.Rows[rows].Cells[0].Value.ToString());
                                        break;
                                    } else {
                                        runCMD("rasdial /disconnect");
                                    }
                                }
                               
                                if (isRunning) {
                                    Thread.Sleep(2500);
                                    var cap = CaptureHelper.CaptureScreen();
                                    cap.Save("res.png");
                                    var sub = ImageScanOpenCV.GetImage("icon.PNG");
                                    var count = 0;
                                    while (count <= 30) {
                                        Point? res = ImageScanOpenCV.FindOutPoint((Bitmap) cap,sub);
                                        if (res != null) {
                                            AutoControl.MouseClick(res.Value.X,res.Value.Y,EMouseKey.LEFT);
                                            Thread.Sleep(time * 1000);
                                            break;
                                        }
                                        count++;
                                    }
                                }
                                runCMD("taskkill /f /im chrome.exe");
                                dataGridView1.Rows[rows].Cells[1].Value = "Disconnect Dcom";
                                runCMD("rasdial /disconnect");
                                if (isRunning) {
                                    dataGridView1.Rows[rows].Cells[1].Value = "Done";
                                } else {
                                    dataGridView1.Rows[rows].Cells[1].Value = "Stop";
                                }
                            } else {
                                dataGridView1.Rows[rows].Cells[1].Value = "Stop";
                            }

                        }

                    }

                }
            }

        }
        int pID;
        int time;
        private void BtnStart_Click(object sender,EventArgs e) {
            isRunning = true;
            button1.Enabled = true;
            btnStart.Enabled = false;
            dataGridView1.Rows.Clear();
            time = Int32.Parse(textBox1.Text);
            using (var fbd = new FolderBrowserDialog()) {
                DialogResult result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath)) {
                    getProfile(fbd.SelectedPath);
                    newThread = new Thread(runChrome);
                    newThread.Start();
                    btnStart.Enabled = true;

                } else {
                    btnStart.Enabled = true;
                    button1.Enabled = false;
                }
            }

        }

        private void Button1_Click(object sender,EventArgs e) {
            isRunning = false;
            btnStart.Enabled = true;
            button1.Enabled = false;
        }
    }
}
