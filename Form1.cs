using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Timer = System.Timers.Timer;
using OpenQA.Selenium.Firefox; // Implement browser mode switch
using System.Linq;
using System.Security.Policy;

namespace VISR
{
    public partial class wndMaster : Form
    {
        System.Windows.Forms.Timer timerC = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer timerD = new System.Windows.Forms.Timer();

        private readonly Dictionary<int, studyRoom> dict_RoomsTUP = new Dictionary<int, studyRoom>(33);
        private IWebDriver driver;
        private IWebDriver driver2;

        private int evictRoom;
        private int freeRooms, refreshSeconds, refreshSecondsExpress = 300, refreshRate = 20, catType = 1;
        private int lastCheckOut;
        private int tally90 = 0, tally60 = 0, tally30 = 0, tallyOverdue = 0;

        bool supressPopups = false;

        Color clrBackground = Color.FromArgb(32, 32, 32);
        Color clrFont = Color.White;
        Color clrTbBack = Color.FromArgb(23, 23, 23);

        private DateTime? oldestOverDueEVICT = DateTime.Now;
        private DateTime? oldestOverDueWARN = DateTime.Now;

        List<int> lstBtnPressed = new List<int>
        {

        };

        private string
            rm2112status,
            rm2113status,
            rm2114status,
            rm2115status,
            rm2116status,
            rm2124status,
            rm2125status,
            rm2126status,
            rm2127status,
            rm2128status,
            rm3101status,
            rm3103status,
            rm3104status,
            rm3105status,
            rm3106status,
            rm3107status,
            rm3112status,
            rm3113status,
            rm3114status,
            rm3115status,
            rm3116status,
            rm3118status,
            rm3119status,
            rm3121status,
            rm3122status,
            rm3123status,
            rm3124status,
            rm3501status,
            rm3502status,
            rm3508status,
            rm3509status,
            rm3510status,
            rm3511status;

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnAccessoryUpdate(object sender, EventArgs e)
        {
            refreshSecondsExpress = 500; // Force update of accessories
            // Don't need this, button1_click does thi already
            //refreshSeconds = 0; // Prevent succesive refresh after schedueled one
            button1_Click_1(null, null);
        }

        private void linkOpener(string recURL)
        {
            Process.Start(new ProcessStartInfo() { FileName = recURL, UseShellExecute = true });
        }


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            linkOpener("https://novacat.nova.edu/search~S13?/atampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&25%2C%2C27/indexsort=-");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&16%2C16%2C/indexsort=r");

        
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1,28,28,E/holdings&FF=atampa+bay+regional+campus+library&13,13,/indexsort=r");
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/atampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&17%2C%2C27/indexsort=-");
        
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/aTampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&19%2C%2C27/indexsort=-");

        
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/aTampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&13%2C%2C27/indexsort=-");

        
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/XTampa+Laptop&searchscope=13&SORT=DZ/XTampa+Laptop&searchscope=13&SORT=DZ&extended=1&SUBKEY=Tampa+Laptop/1%2C3%2C3%2CE/marc&FF=XTampa+Laptop&searchscope=13&SORT=DZ&1%2C1%2C");
        
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&4%2C4%2C/indexsort=r");
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&22%2C22%2C/indexsort=r");
        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkOpener("https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&3%2C3%2C/indexsort=r");
        }

        private void tsmiSettings_Click(object sender, EventArgs e)
        {

        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {



        }

        private void txbAppStatus_TextChanged(object sender, EventArgs e)
        {

        }

        private bool steadyPhase1;

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private int upcomingRoom;

        private string visrSTATUS = "Main Init";
        //string exePath = ".\\geckodriver.exe";


        public wndMaster()
        {
            InitializeComponent();
        }

        private void updateVISRStatus(string Status)
        {
            txbAppStatus.Invoke(new MethodInvoker(delegate { txbAppStatus.Text = visrSTATUS = Status; }));
        }

        private void tsmiSettingsRRDisabled_Click(object sender, EventArgs e)
        {
            refreshRate = 0;
            updateVISRStatus("Waiting");
            tsmiSettingsRRDisabled.Checked = true;
            tsmiSettingsRR30seconds.Checked = false;
            tsmiSettingsRR60seconds.Checked = false;
            tsmiSettingsRR5min.Checked = false;
            tsmiSettingsRRCustom.Checked = false;
        }

        private void tsmiSettingsRR30seconds_Click(object sender, EventArgs e)
        {
            refreshRate = 30;
            updateVISRStatus("Waiting " + Convert.ToString(refreshRate) + "s");
            tsmiSettingsRRDisabled.Checked = false;
            tsmiSettingsRR30seconds.Checked = true;
            tsmiSettingsRR60seconds.Checked = false;
            tsmiSettingsRR5min.Checked = false;
            tsmiSettingsRRCustom.Checked = false;
        }

        private void tsmiSettingsRR60seconds_Click(object sender, EventArgs e)
        {
            refreshRate = 60;
            updateVISRStatus("Waiting " + Convert.ToString(refreshRate) + "s");
            tsmiSettingsRRDisabled.Checked = false;
            tsmiSettingsRR30seconds.Checked = false;
            tsmiSettingsRR60seconds.Checked = true;
            tsmiSettingsRR5min.Checked = false;
            tsmiSettingsRRCustom.Checked = false;
        }

        private void tsmiSettingsRR5min_Click(object sender, EventArgs e)
        {
            refreshRate = 300;
            updateVISRStatus("Waiting " + Convert.ToString(refreshRate) + "s");
            tsmiSettingsRRDisabled.Checked = false;
            tsmiSettingsRR30seconds.Checked = false;
            tsmiSettingsRR60seconds.Checked = false;
            tsmiSettingsRR5min.Checked = true;
            tsmiSettingsRRCustom.Checked = false;
        }

        private void tsmiSettingsRRCustom_Click(object sender, EventArgs e)
        {
            refreshRate = 0;
            updateVISRStatus("Waiting " + Convert.ToString(refreshRate) + "s");
            tsmiSettingsRRDisabled.Checked = false;
            tsmiSettingsRR30seconds.Checked = false;
            tsmiSettingsRR60seconds.Checked = false;
            tsmiSettingsRR5min.Checked = false;
            tsmiSettingsRRCustom.Checked = true;
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            
            Invoke((MethodInvoker)delegate { btnRefreshRooms.BackColor = Color.LimeGreen; });
            Invoke((MethodInvoker)delegate { btnRefreshRooms.Enabled = false; });
            timerD.Interval = 5000;
            timerD.Tick += timer_Tick2;
            timerD.Start();
            Console.WriteLine("FERROW");


            
            updateVISRStatus("Resfresh Start");
            Invoke((MethodInvoker)delegate { _ = SeleniumNOVACAT(sender, e); });
            refreshSeconds = 0;
            //Invoke((MethodInvoker)delegate { btnRefreshRooms.BackColor = SystemColors.Control; });
            //Invoke((MethodInvoker)delegate { btnRefreshRooms.Enabled = true; });
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            updateVISRStatus("Closing");
            driver.Quit();
            //File.Delete(exePath); // Old non NuGet chromedriver deletion
            Application.Exit();
        }

        private void btnRestart_Click(object sender, EventArgs e)
        {
            Application.Restart();
            Environment.Exit(0);
        }

        private void tsmiSettingsBFOff_Click(object sender, EventArgs e)
        {
            steadyPhase1 = true;
            tsmiSettingsBFOn.Checked = false;
            tsmiSettingsBFOff.Checked = true;
        }

        private void tsmiSettingsBFOn_Click(object sender, EventArgs e)
        {
            steadyPhase1 = false;
            tsmiSettingsBFOn.Checked = true;
            tsmiSettingsBFOff.Checked = false;
        }

        private void tsmiSettingsBMChrome_Click(object sender, EventArgs e)
        {
        }

        //Reading and creating the chrome driver file
        private void ExtractResource(string path)
        {
            //byte[] bytes = Properties.Resources.geckodriver;
            //File.WriteAllBytes(path, bytes);
        }

        //Deleting the created file
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            updateVISRStatus("Closing");
            driver.Quit();
            //File.Delete(exePath); // Old chromedriver deletion
            Application.Exit();
        }

        // Dark mode test
        private void SwitchDesign()
        {
            this.ForeColor = clrFont;
            this.BackColor = clrBackground;
            //Now for every special-control that does need an extra color / property to be set use something like this
            foreach (TextBox tb in this.Controls.OfType<TextBox>())
            {
                tb.BackColor = clrTbBack;
                //Maybe do more here...
            }
            //You could now add more controls in a similar fashion.
            this.Invalidate(); //Forces a re-draw of your controls / form
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            updateVISRStatus("Form Load");
            ToolTipSetup();
            // Create a thread and call a background method   
            var backgroundThread = new Thread(driverStartNOVACAT);
            backgroundThread.Name = "Selenium-Worker";
            backgroundThread.Priority = ThreadPriority.Highest;
            // Start thread  
            backgroundThread.Start();
            tsmiSettingsRRCustom.Checked = true;
            //SwitchDesign(); // Dark mode
        }

        private void button_Click(object sender, EventArgs e)
        {
            // Change the background color of the button that was clicked
            Button current = (Button)sender;

            if (current.BackColor != Color.DarkViolet)
            {
                if (!lstBtnPressed.Contains(Convert.ToInt32(current.Text)))
                { 
                    lstBtnPressed.Add(Convert.ToInt32(current.Text));
                }
                current.BackColor = Color.DarkViolet;
                //dict_RoomsTUP[Int32.Parse(current.Text)].roomAvailable = false; // Get room from dict and mark unavailable

            }
            else if (current.BackColor == Color.DarkViolet)
            {
                lstBtnPressed.Remove(Convert.ToInt32(current.Text));
                current.BackColor = Color.LawnGreen;
            }

            foreach (var btn in lstBtnPressed)
            {
                Console.WriteLine(btn);
            }

            // Revert the background color of the previously-colored button, if any
            //if (lastButton != null)
                //lastButton.BackColor = SystemColors.Control;

            // Update the previously-colored button
            //lastButton = current;
        }


        private void ToolTipSetup()
        {
            //// TOOL TIPS
            // Menu buttons
            var toolTipRefresh = new ToolTip();

            // 2nd Floor Courtyard
            var toolTip2112 = new ToolTip();
            var toolTip2113 = new ToolTip();
            var toolTip2114 = new ToolTip();
            var toolTip2115 = new ToolTip();
            var toolTip2116 = new ToolTip();
            var toolTip2124 = new ToolTip();
            var toolTip2125 = new ToolTip();
            var toolTip2126 = new ToolTip();
            var toolTip2127 = new ToolTip();
            var toolTip2128 = new ToolTip();

            // 3rd Floor West Side (Clearwater)
            var toolTip3104 = new ToolTip();
            var toolTip3105 = new ToolTip();
            var toolTip3106 = new ToolTip();
            var toolTip3107 = new ToolTip();

            // 3rd Floor East Side (Tampa)
            var toolTip3508 = new ToolTip();
            var toolTip3509 = new ToolTip();
            var toolTip3510 = new ToolTip();
            var toolTip3511 = new ToolTip();

            // 3rd Floor Courtyard
            var toolTip3112 = new ToolTip();
            var toolTip3113 = new ToolTip();
            var toolTip3114 = new ToolTip();
            var toolTip3115 = new ToolTip();
            var toolTip3116 = new ToolTip();
            var toolTip3118 = new ToolTip();
            var toolTip3119 = new ToolTip();
            var toolTip3121 = new ToolTip();
            var toolTip3122 = new ToolTip();
            var toolTip3123 = new ToolTip();
            var toolTip3124 = new ToolTip();

            // 3rd Floor Waterfront View Tampa Bay
            var toolTip3501 = new ToolTip();
            var toolTip3502 = new ToolTip();

            // Not Study rooms
            var toolTip3101 = new ToolTip();
            var toolTip3103 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTipRefresh.AutoPopDelay =
            toolTip2112.AutoPopDelay = toolTip2113.AutoPopDelay = 
            toolTip2114.AutoPopDelay = toolTip2115.AutoPopDelay =
            toolTip2116.AutoPopDelay = toolTip2124.AutoPopDelay =
            toolTip2125.AutoPopDelay = toolTip2126.AutoPopDelay =
            toolTip2127.AutoPopDelay = toolTip2128.AutoPopDelay =
            toolTip3101.AutoPopDelay = toolTip3103.AutoPopDelay =
            toolTip3104.AutoPopDelay = toolTip3105.AutoPopDelay =
            toolTip3106.AutoPopDelay = toolTip3107.AutoPopDelay =
            toolTip3112.AutoPopDelay = toolTip3113.AutoPopDelay =
            toolTip3114.AutoPopDelay = toolTip3115.AutoPopDelay =
            toolTip3116.AutoPopDelay = toolTip3118.AutoPopDelay =
            toolTip3119.AutoPopDelay = toolTip3121.AutoPopDelay =
            toolTip3122.AutoPopDelay = toolTip3123.AutoPopDelay =
            toolTip3124.AutoPopDelay = toolTip3501.AutoPopDelay =
            toolTip3502.AutoPopDelay = toolTip3508.AutoPopDelay =
            toolTip3509.AutoPopDelay = toolTip3510.AutoPopDelay =
            toolTip3511.AutoPopDelay = 10000;
            
            toolTipRefresh.InitialDelay =
            toolTip2112.InitialDelay = toolTip2113.InitialDelay =
            toolTip2114.InitialDelay = toolTip2115.InitialDelay =
            toolTip2116.InitialDelay = toolTip2124.InitialDelay =
            toolTip2125.InitialDelay = toolTip2126.InitialDelay =
            toolTip2127.InitialDelay = toolTip2128.InitialDelay =
            toolTip3101.InitialDelay = toolTip3103.InitialDelay =
            toolTip3104.InitialDelay = toolTip3105.InitialDelay =
            toolTip3106.InitialDelay = toolTip3107.InitialDelay =
            toolTip3112.InitialDelay = toolTip3113.InitialDelay =
            toolTip3114.InitialDelay = toolTip3115.InitialDelay =
            toolTip3116.InitialDelay = toolTip3118.InitialDelay =
            toolTip3119.InitialDelay = toolTip3121.InitialDelay =
            toolTip3122.InitialDelay = toolTip3123.InitialDelay =
            toolTip3124.InitialDelay = toolTip3501.InitialDelay =
            toolTip3502.InitialDelay = toolTip3508.InitialDelay =
            toolTip3509.InitialDelay = toolTip3510.InitialDelay =
            toolTip3511.InitialDelay = 1000;
            
            toolTipRefresh.ReshowDelay =
            toolTip2112.ReshowDelay = toolTip2113.ReshowDelay =
            toolTip2114.ReshowDelay = toolTip2115.ReshowDelay =
            toolTip2116.ReshowDelay = toolTip2124.ReshowDelay =
            toolTip2125.ReshowDelay = toolTip2126.ReshowDelay =
            toolTip2127.ReshowDelay = toolTip2128.ReshowDelay =
            toolTip3101.ReshowDelay = toolTip3103.ReshowDelay =
            toolTip3104.ReshowDelay = toolTip3105.ReshowDelay =
            toolTip3106.ReshowDelay = toolTip3107.ReshowDelay =
            toolTip3112.ReshowDelay = toolTip3113.ReshowDelay =
            toolTip3114.ReshowDelay = toolTip3115.ReshowDelay =
            toolTip3116.ReshowDelay = toolTip3118.ReshowDelay =
            toolTip3119.ReshowDelay = toolTip3121.ReshowDelay =
            toolTip3122.ReshowDelay = toolTip3123.ReshowDelay =
            toolTip3124.ReshowDelay = toolTip3501.ReshowDelay =
            toolTip3502.ReshowDelay = toolTip3508.ReshowDelay =
            toolTip3509.ReshowDelay = toolTip3510.ReshowDelay =
            toolTip3511.ReshowDelay = 500;

            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTipRefresh.ShowAlways = true;

            // Set up the ToolTip text for the Buttons
            toolTipRefresh.SetToolTip(btnRefreshRooms, 
            "Refreshes all the rooms as they appear on Novacat.\n" +
                   "Make sure that your check-out window on Sierra has\n" +
                   "been refreshed/cleared in order for Novacat to update.");

            toolTip2112.SetToolTip(btn2112, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip2113.SetToolTip(btn2113, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip2114.SetToolTip(btn2114, "Capacity: 6\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip2115.SetToolTip(btn2115, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip2116.SetToolTip(btn2116, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip2124.SetToolTip(btn2124, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip2125.SetToolTip(btn2125, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip2126.SetToolTip(btn2126, "Capacity: 6\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip2127.SetToolTip(btn2127, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip2128.SetToolTip(btn2128, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip3104.SetToolTip(btn3104, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Cold\n" +
                                            "Waterfront View: No");

            toolTip3101.SetToolTip(btn3101, "Capacity: 3\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: Yes");

            toolTip3103.SetToolTip(btn3103, "Capacity: 4\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Cold\n" +
                                            "Waterfront View: No");

            toolTip3105.SetToolTip(btn3105, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Cold\n" +
                                            "Waterfront View: No");

            toolTip3106.SetToolTip(btn3106, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Cold\n" +
                                            "Waterfront View: No");

            toolTip3107.SetToolTip(btn3107, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Cold\n" +
                                            "Waterfront View: No");

            toolTip3112.SetToolTip(btn3112, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3113.SetToolTip(btn3113, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3114.SetToolTip(btn3114, "Capacity: 6\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3115.SetToolTip(btn3115, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3116.SetToolTip(btn3116, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3118.SetToolTip(btn3118, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Cold\n" +
                                            "Waterfront View: No");

            toolTip3119.SetToolTip(btn3119, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3121.SetToolTip(btn3121, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3122.SetToolTip(btn3122, "Capacity: 6\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3123.SetToolTip(btn3123, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3124.SetToolTip(btn3124, "Capacity: 4\n" +
                                            "Outdoor Windows: Yes\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: No");

            toolTip3501.SetToolTip(btn3501, "Capacity: 4\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: Yes");

            toolTip3502.SetToolTip(btn3502, "Capacity: 4\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Warm\n" +
                                            "Waterfront View: Yes");

            toolTip3508.SetToolTip(btn3508, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip3509.SetToolTip(btn3509, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip3510.SetToolTip(btn3510, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");

            toolTip3511.SetToolTip(btn3511, "Capacity: 6\n" +
                                            "Outdoor Windows: No\n" +
                                            "Temperature: Average\n" +
                                            "Waterfront View: No");
        }

        private void driverStartNOVACAT()
        {
            //StartTimer();
            //timer1.Start();
            //ExtractResource(exePath);

            // Chromedrive version (disabled because kept looking in program files for Chrome)
            //string path = Directory.GetCurrentDirectory();
            // Initiate Chrome
            //txtbAppStatus.Invoke(new MethodInvoker(delegate { txtbAppStatus.Text = visrSTATUS =  = "Browser Init";
            updateVISRStatus("Broswer Init");

            tsmiSettingsBMChrome.Checked = true;
            tsmiSettingsBMFirefox.Checked = false;
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("--headless", "--no-sandbox", "--disable-web-security", "--disable-gpu",
                "--incognito", "--proxy-bypass-list=*", "--proxy-server='direct://'", "--log-level=3",
                "--hide-scrollbars"); // Comment this out to see the chrome browser itself
            //chromeOptions.AddArguments();
            

            //var chromeDriverService = ChromeDriverService.CreateDefaultService();
            //chromeDriverService.HideCommandPromptWindow = true;

            // = new ChromeDriver(chromeOptions);
            //driver = new FirefoxDriver();
            var cService = ChromeDriverService.CreateDefaultService();
            cService.HideCommandPromptWindow = true; // Set to false to see chrome browser console
            driver = new ChromeDriver(cService, chromeOptions);
            //driver2 = new ChromeDriver(cService, chromeOptions); // second driver for pagersg


            updateVISRStatus("Get Novacat");
            // Go to NOVACAT

            
            /*driver.Navigate() // Before PagerFetch was used, uncomment to go back
                .GoToUrl("https://novacat.nova.edu/search~S13?/.b3944335/.b3944335/1,1,1,B/holdings~3944335&FF=&1,0");       */
            button1_Click_1(null, null);
            

            //Timer for button flashes
            var aTimer = new Timer();
            // Create a timer and set a two and a half second interval.
            aTimer = new Timer();
            aTimer.Interval = 2500;

            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += OnTimedEvent;

            // Have the timer fire repeated events (true is the default)
            aTimer.AutoReset = true;

            // Start the timer
            aTimer.Enabled = true;

            timer1.Enabled = true;

            var bTimer = new Timer();
            aTimer.Elapsed += OnTimedEvent2;
            // Set the Interval to 1 second.
            aTimer.Interval = 1000;
            bTimer.Enabled = true;
        }


        private async Task SeleniumNOVACAT(object sender, EventArgs e)
        {
            updateVISRStatus("⟳ Novacat");
            await Task.Run(() => driver.Navigate().Refresh()); // Update NOVACAT source
            await Task.Run(() => driver.Navigate()
                .GoToUrl("https://novacat.nova.edu/search~S13?/.b3944335/.b3944335/1,1,1,B/holdings~3944335&FF=&1,0"));

            txbTotalFreeRooms.Text = "31";
            updateVISRStatus("Parse Novacat");
            await Task.Run(() => roomFetch(sender, e)); // Get new information on rooms
            updateVISRStatus("Clear Dict");
            await Task.Run(() => dict_RoomsTUP.Clear()); // Remove all rooms from dictionary
            updateVISRStatus("Populate Dict");
            await Task.Run(() => populateDicts(sender, e)); // Add all rooms to dictionary again
            updateVISRStatus("Reset Buttons");
            resetButtons(sender, e); // Change all button displays to load colors
            updateVISRStatus("Update Buttons");
            roomButtonUpdate(sender, e); // Change all button displays to current room status

            if (refreshSecondsExpress >= 300)
            {
                // Check day, month, year too
                txbCurrentDate.Text = DateTime.Now.ToString("dddd, MMMM d, yyyy");

                textBox6.Text = txbCurrentTime.Text;
                updateVISRStatus("Slow-Count Pagers");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&25%2C%2C27/indexsort=-", textBox10));
                updateVISRStatus("Slow-Count iPads");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&16%2C16%2C/indexsort=r", textBox5));
                updateVISRStatus("Slow-Count Laptops");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/XTampa+Laptop&searchscope=13&SORT=DZ/XTampa+Laptop&searchscope=13&SORT=DZ&extended=1&SUBKEY=Tampa+Laptop/1%2C3%2C3%2CE/marc&FF=XTampa+Laptop&searchscope=13&SORT=DZ&1%2C1%2C", textBox7));
                updateVISRStatus("Slow-Count Markers");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1,28,28,E/holdings&FF=atampa+bay+regional+campus+library&13,13,/indexsort=r", textBox9));
                updateVISRStatus("Slow-Count Lightning");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&17%2C%2C27/indexsort=-", textBox3));
                updateVISRStatus("Slow-Count Headphones");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/aTampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&13%2C%2C27/indexsort=-", textBox8));
                updateVISRStatus("🐢-Count Macbook/USB C");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/aTampa/atampa/1%2C16%2C49%2CE/marc&FF=atampa+bay+regional+campus+library&19%2C%2C27/indexsort=-", textBox4));
                updateVISRStatus("Slow-Count Calculators");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&4%2C4%2C/indexsort=r", textBox300));
                updateVISRStatus("Slow-Count Android /USB C");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&3%2C3%2C/indexsort=r", textBox100));
                updateVISRStatus("Slow-Count Extension Cords");
                await Task.Run(() => fetchItems(sender, e, "https://novacat.nova.edu/search~S13?/atampa+bay/atampa+bay/1%2C28%2C28%2CE/marc&FF=atampa+bay+regional+campus+library&22%2C22%2C/indexsort=r", textBox200));



                //pagerFetch(sender, e);
                //markerFetch(sender, e);

                refreshSecondsExpress = 0;
            }
            // Active the button again and make it control color
            Invoke((MethodInvoker)delegate { btnRefreshRooms.BackColor = SystemColors.Control; });
            Invoke((MethodInvoker)delegate { btnRefreshRooms.Enabled = true; });



            if (refreshRate > 0)
                { 
                    updateVISRStatus("Waiting " + Convert.ToString(refreshRate) + "s");
                }
            else
                { 
                    updateVISRStatus("Waiting");
                }

        }


        private static void OnTimedEvent2(object source, ElapsedEventArgs e)
        {
            //Check reservations
        }

        void timer_Tick(object sender, System.EventArgs e)
        {
            Console.WriteLine("MEEP");
            btnRefreshRooms.Enabled = true;
            btnRefreshRooms.BackColor = SystemColors.Control;
            refreshRate = 29;
            timerC.Stop();
            supressPopups = false;
        }

        void timer_Tick2(object sender, System.EventArgs e)
        {
            Console.WriteLine("MEEP2");
            btnRefreshRooms.Enabled = true;
            btnRefreshRooms.BackColor = SystemColors.Control;
            timerD.Stop();
            supressPopups = false;
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            //button1_Click_1(sender, e);
        }


        private void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            refreshSeconds++;

            refreshSecondsExpress++;
            

            var timeSignalNow = e.SignalTime;
            var timeNow = timeSignalNow.ToString("hh:mm:ss tt");

            // Check current time before attempting to refresh, so that VISR does not update itself outisde of library hours and ping novacat unnecessarily 
            var timeTBRCLOpen = new DateTime(timeSignalNow.Year, timeSignalNow.Month, timeSignalNow.Day, 7, 30, 0);
            var timeTBRCLClose = new DateTime(timeSignalNow.Year, timeSignalNow.Month, timeSignalNow.Day, 23, 59, 0);


            // If signal time is inside library hours, refresh should happen
            if (timeSignalNow >= timeTBRCLOpen && timeSignalNow <= timeTBRCLClose)
            {
                if (refreshRate > 0)
                { 
                    if (refreshSeconds >= refreshRate)
                    {
                        updateVISRStatus("Auto Refresh (" + Convert.ToString(refreshRate) + ")");
                        refreshSeconds = 0;
                        button1_Click_1(sender, e);
                    }
                }
                else // Negative refresh rate
                {
                    // Do not refresh
                }

            }
            // Outside of library hours as of July 2022, no refresh will happen
            else
            {
                updateVISRStatus("Lib Closed");
                // Reset refreshSeconds? Or nah?
            }

            // Update the interface with current time now as a 
            MethodInvoker inv = delegate { txbCurrentTime.Text = timeNow; };
            Invoke(inv);

            var phase1 = false;
            var phase2 = false;
            var idleCount = 0;
            foreach (var entry in dict_RoomsTUP)
            {
                if (entry.Value.roomEvict)
                {
                    if (steadyPhase1)
                    {
                        entry.Value.RoomButton.BackColor = Color.Yellow;
                    }
                    else if (entry.Value.RoomButton.BackColor == Color.Red)
                    {
                        entry.Value.RoomButton.BackColor = Color.Yellow;
                        phase1 = true;
                        phase2 = false;
                    }

                    else
                    {
                        entry.Value.RoomButton.BackColor = Color.Red;
                        phase1 = false;
                        phase2 = true;
                    }
                }

                // Idle phase
                //idleCount++;
                //Console.WriteLine("STATUS OF LATEST: ");
                //Console.WriteLine(entry.Value.roomLatestCheckOut);

                /*if (entry.Value.roomLatestCheckOut == true)
                {
                    if (steadyPhase1 == true)
                    {
                        entry.Value.RoomButton.BackColor = Color.DarkOliveGreen;
                    }
                    else if (entry.Value.RoomButton.BackColor == Color.OrangeRed)
                    {
                        entry.Value.RoomButton.BackColor = Color.DarkOliveGreen;
                        phase1 = true;
                        phase2 = false;
                    }
                    else
                    {
                        entry.Value.RoomButton.BackColor = Color.OrangeRed;
                        phase1 = false;
                        phase2 = true;
                    }
                }*/

                if (entry.Value.roomUpcomingDue)
                {
                    if (steadyPhase1)
                    {
                        entry.Value.RoomButton.BackColor = Color.LightPink;
                    }
                    else if (entry.Value.RoomButton.BackColor == Color.OrangeRed)
                    {
                        entry.Value.RoomButton.BackColor = Color.LightPink;
                        phase1 = true;
                        phase2 = false;
                    }
                    else
                    {
                        entry.Value.RoomButton.BackColor = Color.OrangeRed;
                        phase1 = false;
                        phase2 = true;
                    }
                }

                // Idle phase
                //idleCount++;
                // If neither color states are present on any of the room buttons, meaning two idle counts are collected, then fall back on phase 1F2T
                if (idleCount >= 31)
                {
                    //phase1 = true;
                    //phase2 = false;
                    // Report idleCount reaching 2? To get feedback on how often this is used, I have a feeling it never even gets to 31
                }
                // Reset idle count for next study room
            }
            //idleCount = 0;

            // Match the color phase of room buttons with color key
            if (steadyPhase1)
            {
                button6.BackColor = Color.Yellow;
                button7.BackColor = Color.LightPink;
            }
            else if (phase1 || phase2)
            {
                // Yellow and Pink ON
                if (phase1)
                {
                    button6.BackColor = Color.Yellow;
                    button7.BackColor = Color.LightPink;
                }
                // Yellow and Pink OFF
                else if (phase2)
                {
                    button6.BackColor = Color.Red;
                    button7.BackColor = Color.OrangeRed;
                }
            }
        }


        //STERILIZE TIME
        // 24 hour fix for:
        //roomDueTime = timeParser(sender, e, rm2113status),
        // Word DUE was being formatted into DateTime object, there was no stack trace, had to dig thru and manually comment things out until I found this
        // Still no idea how the program was running fine with this occurring the past 2 weeks

        // Could have been an issue with regional time format of time that I switched over to UK time that caused this problem
        private DateTime? timeParser(object sender, EventArgs e, string inputString)
        {
            Console.WriteLine("INPUT STRING");
            Console.WriteLine(inputString);

            DateTime? timeObj = null;
            // Get rid of words in stats string so that they can be formatted
            if (inputString.Contains("DUE"))
            {
                var stringToParse = inputString.Replace("DUE", "");
                //stringToParse = stringToParse.Replace("PM", "");
                //stringToParse = stringToParse.Replace("AM", "");
                Console.WriteLine("STRING TO PARSE STRING");
                stringToParse = stringToParse.Trim();
                Console.WriteLine(stringToParse);

                //timeObj = DateTime.Parse(stringToParse); DUE 06-23-22 12:58PM
                timeObj = DateTime.ParseExact(stringToParse, "MM-dd-yy hh:mmtt", CultureInfo.InvariantCulture);
            }

            Console.WriteLine("RETURN TIME");
            Console.WriteLine(timeObj);
            return timeObj;
        }

        private void populateDicts(object sender, EventArgs e)
        {
            dict_RoomsTUP.Add(2112,
                new studyRoom
                {
                    roomNumber = 2112,
                    RoomButton = btn2112,
                    RoomLabel = lbl2112,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2112status),
                    roomStatus = rm2112status
                });


            dict_RoomsTUP.Add(2113,
                new studyRoom
                {
                    roomNumber = 2113,
                    RoomButton = btn2113,
                    RoomLabel = lbl2113,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2113status),
                    roomStatus = rm2113status
                });


            dict_RoomsTUP.Add(2114,
                new studyRoom
                {
                    roomNumber = 2114,
                    RoomButton = btn2114,
                    RoomLabel = lbl2114,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2114status),
                    roomStatus = rm2114status
                });


            dict_RoomsTUP.Add(2115,
                new studyRoom
                {
                    roomNumber = 2115,
                    RoomButton = btn2115,
                    RoomLabel = lbl2115,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2115status),
                    roomStatus = rm2115status
                });

            dict_RoomsTUP.Add(2116,
                new studyRoom
                {
                    roomNumber = 2116,
                    RoomButton = btn2116,
                    RoomLabel = lbl2116,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2116status),
                    roomStatus = rm2116status
                });

            dict_RoomsTUP.Add(2124,
                new studyRoom
                {
                    roomNumber = 2124,
                    RoomButton = btn2124,
                    RoomLabel = lbl2124,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2124status),
                    roomStatus = rm2124status
                });

            dict_RoomsTUP.Add(2125,
                new studyRoom
                {
                    roomNumber = 2125,
                    RoomButton = btn2125,
                    RoomLabel = lbl2125,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2125status),
                    roomStatus = rm2125status
                });

            dict_RoomsTUP.Add(2126,
                new studyRoom
                {
                    roomNumber = 2126,
                    RoomButton = btn2126,
                    RoomLabel = lbl2126,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2126status),
                    roomStatus = rm2126status
                });

            dict_RoomsTUP.Add(2127,
                new studyRoom
                {
                    roomNumber = 2127,
                    RoomButton = btn2127,
                    RoomLabel = lbl2127,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2127status),
                    roomStatus = rm2127status
                });

            dict_RoomsTUP.Add(2128,
                new studyRoom
                {
                    roomNumber = 2128,
                    RoomButton = btn2128,
                    RoomLabel = lbl2128,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm2128status),
                    roomStatus = rm2128status
                });

            dict_RoomsTUP.Add(3104,
                new studyRoom
                {
                    roomNumber = 3104,
                    RoomButton = btn3104,
                    RoomLabel = lbl3104,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3104status),
                    roomStatus = rm3104status
                });

            dict_RoomsTUP.Add(3105,
                new studyRoom
                {
                    roomNumber = 3105,
                    RoomButton = btn3105,
                    RoomLabel = lbl3105,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3105status),
                    roomStatus = rm3105status
                });

            dict_RoomsTUP.Add(3106,
                new studyRoom
                {
                    roomNumber = 3106,
                    RoomButton = btn3106,
                    RoomLabel = lbl3106,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3106status),
                    roomStatus = rm3106status
                });

            dict_RoomsTUP.Add(3107,
                new studyRoom
                {
                    roomNumber = 3107,
                    RoomButton = btn3107,
                    RoomLabel = lbl3107,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3107status),
                    roomStatus = rm3107status
                });

            dict_RoomsTUP.Add(3112,
                new studyRoom
                {
                    roomNumber = 3112,
                    RoomButton = btn3112,
                    RoomLabel = lbl3112,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3112status),
                    roomStatus = rm3112status
                });

            dict_RoomsTUP.Add(3113,
                new studyRoom
                {
                    roomNumber = 3113,
                    RoomButton = btn3113,
                    RoomLabel = lbl3113,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3113status),
                    roomStatus = rm3113status
                });

            dict_RoomsTUP.Add(3114,
                new studyRoom
                {
                    roomNumber = 3114,
                    RoomButton = btn3114,
                    RoomLabel = lbl3114,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3114status),
                    roomStatus = rm3114status
                });

            dict_RoomsTUP.Add(3115,
                new studyRoom
                {
                    roomNumber = 3115,
                    RoomButton = btn3115,
                    RoomLabel = lbl3115,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3115status),
                    roomStatus = rm3115status
                });

            dict_RoomsTUP.Add(3116,
                new studyRoom
                {
                    roomNumber = 3116,
                    RoomButton = btn3116,
                    RoomLabel = lbl3116,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3116status),
                    roomStatus = rm3116status
                });

            dict_RoomsTUP.Add(3118,
                new studyRoom
                {
                    roomNumber = 3118,
                    RoomButton = btn3118,
                    RoomLabel = lbl3118,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3118status),
                    roomStatus = rm3118status
                });

            dict_RoomsTUP.Add(3119,
                new studyRoom
                {
                    roomNumber = 3119,
                    RoomButton = btn3119,
                    RoomLabel = lbl3119,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3119status),
                    roomStatus = rm3119status
                });

            dict_RoomsTUP.Add(3121,
                new studyRoom
                {
                    roomNumber = 3121,
                    RoomButton = btn3121,
                    RoomLabel = lbl3121,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3121status),
                    roomStatus = rm3121status
                });

            dict_RoomsTUP.Add(3122,
                new studyRoom
                {
                    roomNumber = 3122,
                    RoomButton = btn3122,
                    RoomLabel = lbl3122,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3122status),
                    roomStatus = rm3122status
                });

            dict_RoomsTUP.Add(3123,
                new studyRoom
                {
                    roomNumber = 3123,
                    RoomButton = btn3123,
                    RoomLabel = lbl3123,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3123status),
                    roomStatus = rm3123status
                });

            dict_RoomsTUP.Add(3124,
                new studyRoom
                {
                    roomNumber = 3124,
                    RoomButton = btn3124,
                    RoomLabel = lbl3124,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3124status),
                    roomStatus = rm3124status
                });

            dict_RoomsTUP.Add(3508,
                new studyRoom
                {
                    roomNumber = 3508,
                    RoomButton = btn3508,
                    RoomLabel = lbl3508,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3508status),
                    roomStatus = rm3508status
                });

            dict_RoomsTUP.Add(3509,
                new studyRoom
                {
                    roomNumber = 3509,
                    RoomButton = btn3509,
                    RoomLabel = lbl3509,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3509status),
                    roomStatus = rm3509status
                });

            dict_RoomsTUP.Add(3510,
                new studyRoom
                {
                    roomNumber = 3510,
                    RoomButton = btn3510,
                    RoomLabel = lbl3510,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3510status),
                    roomStatus = rm3510status
                });

            dict_RoomsTUP.Add(3511,
                new studyRoom
                {
                    roomNumber = 3511,
                    RoomButton = btn3511,
                    RoomLabel = lbl3511,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3511status),
                    roomStatus = rm3511status
                });

            dict_RoomsTUP.Add(3501,
                new studyRoom
                {
                    roomNumber = 3501,
                    RoomButton = btn3501,
                    RoomLabel = lbl3501,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3501status),
                    roomStatus = rm3501status
                });

            dict_RoomsTUP.Add(3502,
                new studyRoom
                {
                    roomNumber = 3502,
                    RoomButton = btn3502,
                    RoomLabel = lbl3502,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3502status),
                    roomStatus = rm3502status
                });

            // IDEA LABS
            dict_RoomsTUP.Add(3101,
                new studyRoom
                {
                    roomNumber = 3101,
                    RoomButton = btn3101,
                    RoomLabel = lbl3101,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3101status),
                    roomStatus = rm3101status
                });
            dict_RoomsTUP.Add(3103,
                new studyRoom
                {
                    roomNumber = 3103,
                    RoomButton = btn3103,
                    RoomLabel = lbl3103,
                    roomAvailable = true,
                    roomEvict = false,
                    roomDueTime = timeParser(sender, e, rm3103status),
                    roomStatus = rm3103status
                });
        }

        private void resetButtons(object sender, EventArgs e)
        {
            foreach (var entry in dict_RoomsTUP)
            {
                if (!lstBtnPressed.Contains(entry.Key))
                {
                    entry.Value.RoomButton.BackColor = Color.White;
                    entry.Value.RoomLabel.BackColor = Color.Magenta;
                    entry.Value.RoomLabel.Text = "...";
                }
            }
        }

        // Consider using Dynamic return type if String and Int are used
        private string xPATHMatrix(int format, int column = 2, int row = 4)
        {
            var format1 = "//*[@id=\"bib_items\"]/tbody/tr[" + column + "]/td[" + row + "]";
            var format2 = "//*[@id=\"bib_items\"]/tbody/tr[" + column + "]/td[" + row + "]";
            if (format == 1)
                return driver.FindElement(By.XPath(format1)).Text;
            return driver.FindElement(By.XPath(format2)).Text;
        }

        // Count pagers due
        public String fetchItems(object sender, EventArgs e, String searchCategory, TextBox TxtbTarget)
        {
            refreshSeconds = 0; // Reset this so we don't have accidental room refresh while doing an accessory refresh
            driver.Navigate()
                .GoToUrl(searchCategory);           
            //var body = driver.FindElement(By.XPath("//div[. = 'TextToFind']"));
            IWebElement l = driver.FindElement(By.TagName("body"));
            string mar = l.Text;     
            Console.WriteLine("lantern");
            Console.WriteLine(mar);
            int magic = Regex.Matches(mar, "(DUE)").Count;
            var magicString = Convert.ToString(magic);  // this never appears
            //txtbStudyRoomsFree.Text = Convert.ToString(magic);
            Console.WriteLine("magic");
            Console.WriteLine(magic);
            Invoke((MethodInvoker)delegate { TxtbTarget.Text = magicString; });
            Thread.Sleep(2000); // Sleep 2 seconds
            return magicString;
        }


        private void roomFetch(object sender, EventArgs e)
        {
            try
            {
                rm2112status = xPATHMatrix(2);
                rm2113status = xPATHMatrix(2, 3);
                rm2114status = xPATHMatrix(2, 4);
                rm2115status = xPATHMatrix(2, 5);
                rm2116status = xPATHMatrix(2, 6);
                rm2124status = xPATHMatrix(2, 7);
                rm2125status = xPATHMatrix(2, 8);
                rm2126status = xPATHMatrix(2, 9);
                rm2127status = xPATHMatrix(2, 10);
                rm2128status = xPATHMatrix(2, 11);
                rm3101status = xPATHMatrix(2, 12);
                rm3103status = xPATHMatrix(2, 13);
                rm3104status = xPATHMatrix(2, 14);
                rm3105status = xPATHMatrix(2, 15);
                rm3106status = xPATHMatrix(2, 16);
                rm3107status = xPATHMatrix(2, 17);
                rm3112status = xPATHMatrix(2, 18);
                rm3113status = xPATHMatrix(2, 19);
                rm3114status = xPATHMatrix(2, 20);
                rm3115status = xPATHMatrix(2, 21);
                rm3116status = xPATHMatrix(2, 22);
                rm3118status = xPATHMatrix(2, 23);
                rm3119status = xPATHMatrix(2, 24);
                rm3121status = xPATHMatrix(2, 25);
                rm3122status = xPATHMatrix(2, 26);
                rm3123status = xPATHMatrix(2, 27);
                rm3124status = xPATHMatrix(2, 28);
                rm3501status = xPATHMatrix(2, 29);
                rm3502status = xPATHMatrix(2, 30);
                rm3508status = xPATHMatrix(2, 31);
                rm3509status = xPATHMatrix(2, 32);
                rm3510status = xPATHMatrix(2, 33);
                rm3511status = xPATHMatrix(2, 34);
            }

            catch (Exception ex)
            {
                var error429 = driver.FindElement(By.XPath("/ html / body / h1")).Text;
                var errorMessage = "Novacat cannot be reached, either the website is down, you don't have Internet, or too many refresh requests were made within a short amount of time.\n\n Click OK to close this message.\n\n" + error429 + ex.Message + "\n\n Time of error: " + txbCurrentTime.Text;

                // 429 Too many requests temporary fix

                if (supressPopups == false)
                {
                    MethodInvoker inv = delegate
                    {
                        // Safety first, avoid spamming any more requests
                        refreshSeconds = 0;
                        refreshSecondsExpress = 0;
                        updateVISRStatus("Error!");
                        // Stop refreshing to avoid spam of pop-up boxes

                        btnRefreshRooms.BackColor = Color.Red;
                        btnRefreshRooms.Enabled = false;
                        // New method that runs on new thread that should avoid pausing program
                        new Thread(() => System.Windows.Forms.MessageBox.Show(errorMessage, "There is a problem",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)).Start();
                        // Old method that runs on same thread, causing pop up box to stop execution
                        /*
                        MessageBox.Show(errorMessage, "There is a problem",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);*/
                        supressPopups = true;

                        //refreshRate = 60; // Try a refresh again after OK was hit
                        btnRefreshRooms.BackColor = Color.Orange;
                        btnRefreshRooms.Enabled = false; // Redundant enable off


                        //wait(5000); //wait five second
                        //// Create a timer and set a one second interval.

                        // NEW wait
                        btnRefreshRooms.Enabled = false;

                        timerC.Interval = 10000;
                        timerC.Tick += timer_Tick;
                        timerC.Start();
                        Console.WriteLine("YARGON");

                    };
                    Invoke(inv);
                }
            }
        }

        private void roomButtonUpdate(object sender, EventArgs e)
        {
            Console.WriteLine("BUTTON UPDATE");
            var x30MinsLater = DateTime.Now.AddMinutes(30);
            var x60MinsLater = DateTime.Now.AddMinutes(60);
            var x90MinsLater = DateTime.Now.AddMinutes(90);
            freeRooms = -2;
            evictRoom = 0;
            oldestOverDueEVICT = DateTime.Now;
            TimeSpan? shortestDue = new TimeSpan(52, 0, 0, 0);
            // Reset color counters
            tally90 = tally60 = tally30 = tallyOverdue = 0;

            foreach (var entry in dict_RoomsTUP)
            {

                //if (!lstBtnPressed.Contains(entry.Key))
                //{ //BtnPressed Nest

                string formatStatus;
                entry.Value.latestCheckOut = false;


                Debug.WriteLine(entry.Value.roomNumber);
                entry.Value.RoomButton.BackColor = Color.Magenta;
                entry.Value.RoomLabel.BackColor = Color.Magenta;

                //if (entry.Value.roomStatus.Contains("AVAILABLE") || (entry.Value.roomStatus.Contains("Recently")))


                if (entry.Value.roomStatus.Contains("AVAILABLE")
                    || entry.Value.roomStatus.Contains("Recently")
                    || entry.Value.roomStatus.Contains("DUE"))
                {
                    if (entry.Value.roomStatus.Contains("AVAILABLE"))
                    {
                            if (lstBtnPressed.Contains(entry.Key))
                            { 
                                entry.Value.RoomButton.BackColor = Color.DarkViolet; 
                            }
                            else
                            {
                                entry.Value.RoomLabel.Text = "Ready";
                                entry.Value.roomAvailable = true;

                                entry.Value.RoomLabel.BackColor = Color.PaleGreen;
                                entry.Value.RoomButton.BackColor = Color.LawnGreen;
                            }
                    }
                    else
                    {
                        var bodySPC = @"\s";
                        var headSPC = @"^\s";
                        var tailSPC = @"\s$";

                        formatStatus = Regex.Replace(entry.Value.roomStatus, bodySPC, "\n");
                        formatStatus = Regex.Replace(formatStatus, headSPC, "");
                        formatStatus = Regex.Replace(formatStatus, tailSPC, "");

                        entry.Value.RoomLabel.Text = formatStatus;
                        //entry.Value.RoomLabel.BackColor = Color.GreenYellow;
                    }


                        if (entry.Value.roomStatus.Contains("Recently"))
                        {
                            if (lstBtnPressed.Contains(entry.Key))
                            { 
                                entry.Value.RoomButton.BackColor = Color.DarkViolet;
                            }
                            else 
                            {
                                entry.Value.roomAvailable = true;
                                entry.Value.RoomLabel.BackColor = Color.LightGreen;
                                entry.Value.RoomButton.BackColor = Color.LawnGreen; 
                            }
                    }


                    else if (entry.Value.roomStatus.Contains("DUE"))
                    {
                            // Automatic removal of unavailable rooms, comment this section out
                            /*if (lstBtnPressed.Contains(entry.Key))
                            { 
                                entry.Value.RoomButton.BackColor = Color.DarkViolet;
                            }*/
                            //else

                            // Manually make rooms unavailable, but when they are due, remove unavailable marker automatically
                            {
                                entry.Value.roomAvailable = false;

                                entry.Value.RoomLabel.BackColor = Color.SkyBlue;
                                if (entry.Value.roomEvict ==
                                    false) // Only OrangeRed if its not a soon to be evicted room, no sense in changing colors twice in one pass
                                    entry.Value.RoomButton.BackColor = Color.OrangeRed;
                                else
                                    entry.Value.RoomButton.BackColor = Color.Pink;


                                Console.WriteLine("ROOM DUE TIME >>>");
                                Console.WriteLine(entry.Value.roomDueTime);
                                Console.WriteLine("TIME NOW >>>");
                                Console.WriteLine(DateTime.Now);

                                // Nearing due time
                                // Nearing due time
                                if (entry.Value.roomDueTime <= x90MinsLater)
                                    entry.Value.RoomLabel.BackColor = Color.AntiqueWhite;
                                if (entry.Value.roomDueTime <= x60MinsLater) entry.Value.RoomLabel.BackColor = Color.RosyBrown;
                                if (entry.Value.roomDueTime <= x30MinsLater) entry.Value.RoomLabel.BackColor = Color.Orange;

                                // OVERDUE
                                if (DateTime.Now > entry.Value.roomDueTime)
                                {
                                    Console.Write(Convert.ToString(DateTime.Now), " > ",
                                        Convert.ToString(entry.Value.roomDueTime));
                                    // Change to overdue and mark yellow
                                    entry.Value.RoomLabel.BackColor = Color.Yellow;
                                    if (entry.Value.roomStatus.Contains("OVERDUE"))
                                    {
                                        // Nothing
                                    }
                                    else
                                    {
                                        //var lblStringTime = entry.Value.RoomLabel.Text.Substring(4);
                                        //DateTime? lblTime = timeParser(sender, e, lblStringTime);
                                        //Console.WriteLine("EVICT TIME OPEN");
                                        //Console.WriteLine(lblStringTime);
                                        //Console.WriteLine("EVICT TIME CLOSE");


                                        // Find oldest overdue room
                                        if (entry.Value.roomDueTime < oldestOverDueEVICT)
                                        {
                                            oldestOverDueEVICT = entry.Value.roomDueTime;
                                            evictRoom = entry.Value.roomNumber;

                                            entry.Value.roomEvict = true;
                                            Console.WriteLine("OLDEST");
                                            Console.WriteLine(evictRoom);
                                        }
                                        else
                                        {
                                            entry.Value.roomEvict = false;
                                        }

                                        Console.WriteLine(entry.Value.roomDueTime);
                                        Console.WriteLine("TIME NOW IS ", DateTime.Now);
                                        Console.WriteLine("ROOM DUE TIME IS ", entry.Value.roomDueTime);

                                        entry.Value.RoomLabel.Text = entry.Value.RoomLabel.Text.Replace("DUE", "OVERDUE");
                                        entry.Value.RoomLabel.BackColor = Color.Yellow;
                                        entry.Value.RoomButton.BackColor = Color.Red; // Moved over from tick rate
                                    }
                                }

                                // Oldest Due

                                else if (DateTime.Now <= entry.Value.roomDueTime)
                                {
                                    Console.WriteLine("OLDEST DUE");
                                    if (entry.Value.roomStatus.Contains("DUE"))
                                    {
                                        Console.WriteLine(entry.Value.roomStatus);
                                        Console.WriteLine(entry.Value.roomDueTime);
                                        Console.WriteLine("WARN TIME CLOSE");
                                        var lblShortDiff = entry.Value.roomDueTime - DateTime.Now;
                                        Console.WriteLine("Difference " + lblShortDiff);

                                        if (lblShortDiff < shortestDue)
                                        {
                                            shortestDue = lblShortDiff;
                                            upcomingRoom = entry.Value.roomNumber;
                                            //entry.Value.roomUpcomingDue = true;
                                        }

                                        Console.WriteLine("SHORTEST DUE" + shortestDue);
                                        // Find oldest overdue room
                                        //if (lblTime > oldestOverDueWARN)
                                        if (entry.Value.roomDueTime > oldestOverDueWARN)
                                        {
                                            oldestOverDueWARN = entry.Value.roomDueTime;
                                            lastCheckOut = entry.Value.roomNumber;

                                            //entry.Value.latestCheckOut = true;
                                            //entry.Value.roomLatestCheckOut = true;
                                            Console.WriteLine("OLDEST WARNING");
                                            Console.WriteLine(lastCheckOut);
                                        }
                                    }
                                }
                            }
                    }
                //} //BtnPressed Nest

                }

                if (entry.Value.roomAvailable) freeRooms++;             

                /*if (entry.Value.roomLatestCheckOut == true)
                {
                    entry.Value.RoomLabel.BackColor = Color.LawnGreen;
                }*/

                

                // Count colors
                if (entry.Value.RoomLabel.BackColor == Color.AntiqueWhite)
                    tally90++;
                else if (entry.Value.RoomLabel.BackColor == Color.RosyBrown)
                    tally60++;
                else if (entry.Value.RoomLabel.BackColor == Color.Orange)
                    tally30++;
                else if (entry.Value.RoomLabel.BackColor == Color.Yellow)
                    tallyOverdue++;
            }

            count90.Text = tally90.ToString();
            count60.Text = tally60.ToString();
            count30.Text = tally30.ToString();
            countOverdue.Text = tallyOverdue.ToString();

            if (evictRoom > 0)
            {
                dict_RoomsTUP[evictRoom].roomEvict = true;
                Console.WriteLine("EVICTROOM IS");
                Console.WriteLine(evictRoom);
            }

            if (lastCheckOut > 0)
            {
                dict_RoomsTUP[lastCheckOut].roomLatestCheckOut = true;
                dict_RoomsTUP[lastCheckOut].RoomLabel.BackColor = Color.MediumAquamarine;
                Console.WriteLine("lastCheckOut IS");
                Console.WriteLine(lastCheckOut);
            }

            if (upcomingRoom > 0)
            {
                dict_RoomsTUP[upcomingRoom].roomUpcomingDue = true;
                //dict_RoomsTUP[upcomingRoom].RoomLabel.BackColor = Color.Peru;
                Console.WriteLine("UPCOMING IS");
                Console.WriteLine(upcomingRoom);
            }

            txbCurrentFreeRooms.Text = Convert.ToString(freeRooms);
        }
    }


    internal class studyRoom
    {
        public int roomNumber { get; set; }
        public Button RoomButton { get; set; }
        public Label RoomLabel { get; set; }
        public bool roomAvailable { get; set; }
        public bool roomEvict { get; set; }
        public bool latestCheckOut { get; set; }
        public bool roomUpcomingDue { get; set; }
        public bool roomLatestCheckOut { get; set; }
        public DateTime? roomDueTime { get; set; }
        public string roomStatus { get; set; } = "N/A";
    }
}