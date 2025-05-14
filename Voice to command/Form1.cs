// Voice Command Application - Enhanced Version
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Windows.Forms;

namespace Voice_to_command
{
    public partial class Form1 : Form
    {
        private SpeechRecognitionEngine recognizer;
        private bool isListening = false;
        private Timer timer;
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private List<string> commandHistory = new List<string>();
        [DllImport("user32.dll")] private static extern bool LockWorkStation();
        [DllImport("user32.dll", SetLastError = true)] private static extern bool ExitWindowsEx(uint uFlags, uint dwReason);
        [DllImport("user32.dll")] private static extern int SendMessage(IntPtr hWnd, int hMsg, int wParam, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)] private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const int SC_MONITORPOWER = 0xF170;
        private const int WM_SYSCOMMAND = 0x0112;
        private static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);
        private const byte VK_VOLUME_MUTE = 0xAD;
        private const byte VK_VOLUME_UP = 0xAF;
        private const byte VK_VOLUME_DOWN = 0xAE;
        private const int KEYEVENTF_KEYDOWN = 0x0000;
        private const int KEYEVENTF_KEYUP = 0x0002;

        public Form1()
        {
            InitializeComponent();
            InitializeSpeechRecognition();
            InitializeTimer();
            InitializeTrayIcon();
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.Black;
            this.TopMost = false;
            this.BackgroundImageLayout = ImageLayout.Stretch;
        }
        private void MinimizeToBackground()
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            trayIcon.Visible = true;
            trayIcon.BalloonTipTitle = "Voice to Command";
            trayIcon.BalloonTipText = "Hi,we here\nWhatever you want";
            trayIcon.ShowBalloonTip(1000);
        }

        private void InitializeTimer()
        {
            timer = new Timer { Interval = 1000 };
            timer.Tick += (s, e) =>
            {
                label2.Text = "Time: " + DateTime.Now.ToString("HH:mm:ss");
                label3.Text = "Date: " + DateTime.Now.ToString("yyyy-MM-dd");
            };
            timer.Start();
        }

        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();

            trayMenu.Items.Add("Restore", null, (s, e) => RestoreFromTray());
            trayMenu.Items.Add("Pause Listening", null, (s, e) => PauseRecognition());
            trayMenu.Items.Add("Resume Listening", null, (s, e) => ResumeRecognition());
            trayMenu.Items.Add("Exit", null, (s, e) => Application.Exit());

            trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                ContextMenuStrip = trayMenu,
                Visible = false,
                Text = "Voice to Command"
            };
            trayIcon.DoubleClick += (s, e) => RestoreFromTray();
        }
        private void PauseRecognition()
        {
            if (isListening)
            {
                recognizer.RecognizeAsyncStop();
                isListening = false;
                label1.Text = "Recognition paused.";
                trayIcon.BalloonTipText = "Voice recognition paused.";
                trayIcon.ShowBalloonTip(1000);
            }
        }

        private void ResumeRecognition()
        {
            if (!isListening)
            {
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                isListening = true;
                label1.Text = "Listening...";
                trayIcon.BalloonTipText = "Voice recognition resumed.";
                trayIcon.ShowBalloonTip(1000);
            }
        }


        private void RestoreFromTray()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
        }

        private void InitializeSpeechRecognition()
        {
            try
            {
                recognizer = new SpeechRecognitionEngine(new CultureInfo("en-US"));
                recognizer.SetInputToDefaultAudioDevice();

                var choices = new Choices(new string[]
                {
                    "open calc", "open mail", "open download", "open microsoft edge",
                    "open clock and alarm", "open paint", "open personalization settings",
                    "do a scan with windows defender", "open notepad", "close notepad",
                    "shutdown", "restart", "log off", "lock screen", "mute volume", "unmute volume",
                    "increase volume", "decrease volume", "open settings", "open browser",
                    "take screenshot", "open task manager", "turn off monitor",
                    "exit", "open cmd",
                    "open user temp folder", "open control panel", "open microsoft whiteboard", "open weather"
                });

                var grammar = new Grammar(new GrammarBuilder(choices));
                recognizer.LoadGrammar(grammar);
                recognizer.SpeechRecognized += Recognizer_SpeechRecognized;
                recognizer.SpeechRecognitionRejected += (s, e) => label1.Text = "Speech not recognized.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing speech recognition: " + ex.Message);
            }
        }

        private void Recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Confidence >= 0.75)
            {
                string command = e.Result.Text.ToLower();
                label1.Text = "Recognized: " + command;
                commandHistory.Add($"{DateTime.Now}: {command}");
                if (command == "shutdown" || command == "restart" || command == "log off")
                {
                    var result = MessageBox.Show($"Are you sure you want to {command}?", "Confirm", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                        ExecuteCommand(command);
                }
                else
                {
                    ExecuteCommand(command);
                }
            }
            else
            {
                label1.Text = "Low confidence, try again.";
            }
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                MinimizeToBackground();
            }
        }


        private void ExecuteCommand(string command)
        {
            try
            {
                switch (command)
                {
                    case "open calculator":
                    case "open calc":
                        Process.Start("calc");
                        break;
                    case "open mail": Process.Start("mailto:");
                        break;
                    case "open download":
                        string downloads = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
                        Process.Start("explorer.exe", downloads);
                        break;
                    case "open microsoft edge": Process.Start("msedge");
                        break;
                    case "open clock and alarm": Process.Start("ms-clock:");
                        break;
                    case "open paint": Process.Start("mspaint");
                        break;
                    case "open personalization settings":
                        Process.Start("ms-settings:personalization");
                        break;
                    case "do a scan with windows defender": Process.Start("powershell", "Start-MpScan -ScanType QuickScan");
                        break;
                    case "open notepad": Process.Start("notepad");
                        break;
                    case "close notepad": KillProcess("notepad");
                        break;
                    case "shutdown": Process.Start("shutdown", "/s /t 0");
                        break;
                    case "restart": Process.Start("shutdown", "/r /t 0");
                        break;
                    case "log off": ExitWindowsEx(0, 0);
                        break;
                    case "lock screen": LockWorkStation();
                        break;
                    case "mute volume":
                    case "unmute volume":
                        keybd_event(VK_VOLUME_MUTE, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        keybd_event(VK_VOLUME_MUTE, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                        break;
                    case "increase volume":
                        for (int i = 0; i < 5; i++)
                        {
                            keybd_event(VK_VOLUME_UP, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                            keybd_event(VK_VOLUME_UP, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                        }
                        break;
                    case "decrease volume":
                        for (int i = 0; i < 5; i++)
                        {
                            keybd_event(VK_VOLUME_DOWN, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                            keybd_event(VK_VOLUME_DOWN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                        }
                        break;
                    case "open settings": Process.Start("ms-settings:");
                        break;
                    case "open browser": Process.Start("chrome");
                        break;
                    case "take screenshot":
                        string screenshotPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), $"screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        using (var bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
                        {
                            using (var g = Graphics.FromImage(bmp))
                            {
                                g.CopyFromScreen(Point.Empty, Point.Empty, bmp.Size);
                            }
                            bmp.Save(screenshotPath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        MessageBox.Show("Screenshot saved to Pictures.");
                        break;
                    case "open task manager": Process.Start("taskmgr");
                        break;
                    case "turn off monitor": SendMessage(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, (IntPtr)2);
                        break;
                    case "exit":
                        Application.Exit();
                        break;
                    case "open cmd": Process.Start("cmd.exe");
                        break;
                    case "open user temp folder": Process.Start("explorer.exe", Environment.GetEnvironmentVariable("TEMP"));
                        break;
                    case "open control panel": Process.Start("control.exe");
                        break;
                    case "open microsoft whiteboard":
                        Process.Start("ms-whiteboard-cmd:");
                        break;
                    case "open weather": Process.Start("msnweather:");
                        break;
                    default:
                        MessageBox.Show("Command not supported or not yet implemented: " + command);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Execution failed: " + ex.Message);
            }
        }

        private void KillProcess(string name)
        {
            foreach (var proc in Process.GetProcessesByName(name))
            {
                proc.Kill();
            }
        }
        bool isclicked=false;
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (!isListening)
            {
                pictureBox1.Image = Image.FromFile("C:\Users\mohammad\Desktop\voice to command\project files\\8862366.png");
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                isListening = true;
                label1.Text = "Listening...";
            }
            else
            {
                pictureBox1.Image = Image.FromFile("E:\\University\\~Lectures pictures & notes\\project files\\1262311.png");
                recognizer.RecognizeAsyncStop();
                isListening = false;
                label1.Text = "Stopped.";
            }
            isclicked = !isclicked;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string[] categories = {
                "System: shutdown, restart, log off, lock screen, turn off monitor",
                "Apps: open calculator, open mail, open cmd, open browser",
                "Tools: open settings, open task manager, take screenshot, open notepad, close notepad, open paint",
                "Volume: mute volume, unmute volume, increase volume, decrease volume, Exit",
                "Security: do a scan with windows defender",
                "Settings: open personalization settings, open microsoft whiteboard",
                "Utilities: open clock and alarm, open user temp folder, open control panel, open weather"
            };
            MessageBox.Show("Supported Commands:\n\n" + string.Join("\n", categories));
        }

        private void button2_Click(object sender, EventArgs e) => MessageBox.Show("Contact: support@voicecommand.com\nPhone: +1234567890");

        private void button3_Click(object sender, EventArgs e) => MessageBox.Show("Help:\n- Click mic icon to toggle listening.\n- Speak a supported command clearly.\n- Confirm shutdown/restart commands.");

        private void button4_Click(object sender, EventArgs e)
        {
            this.TopMost = !this.TopMost;
            button4.Text = this.TopMost ? "Un top" : "Keep on top";
        }
        private void button5_Click(object sender, EventArgs e)
        {
            MinimizeToBackground();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (commandHistory.Count == 0)
            {
                MessageBox.Show("No commands recognized yet.");
            }
            else
            {
                string history = string.Join("\n", commandHistory);
                MessageBox.Show("Command History:\n\n" + history);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
