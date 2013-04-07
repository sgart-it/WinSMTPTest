using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Resources;
using System.Reflection;

namespace WinSMTPTest
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>


    public class FormWinSMTP : System.Windows.Forms.Form
    {
        private const string FILE_TO_LOG = "winsmtptest.log";
        private ClassTCP cTcp;
        private int sendCount = 1;
        private string helpAboutMessage = "";
        private bool writeToLogFile = false;

        private static string serverSMTP = "";
        private static string serverSMTPPort = "";
        private static string commandTemplate;
        private static bool flagSilent;

        private System.Windows.Forms.Label lblServerPort;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtServerSMTP;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.ContextMenu cmLstResponse;
        private System.Windows.Forms.MenuItem miCopy;
        private System.Windows.Forms.MenuItem miClear;
        private System.Windows.Forms.TextBox txtServerPort;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem menuItem4;
        private System.Windows.Forms.MenuItem menuItem9;
        private System.Windows.Forms.MenuItem mmiFile;
        private System.Windows.Forms.MenuItem mmiOpenModel;
        private System.Windows.Forms.MenuItem mmiSaveModel;
        private System.Windows.Forms.MenuItem mmiExit;
        private System.Windows.Forms.MenuItem mmiLog;
        private System.Windows.Forms.MenuItem mmiCopy;
        private System.Windows.Forms.MenuItem mmiClear;
        private System.Windows.Forms.MenuItem mmiSaveLog;
        private System.Windows.Forms.MenuItem mmiHelp;
        private System.Windows.Forms.MenuItem mmiAbout;
        private System.Windows.Forms.MenuItem menuItem3;
        private System.Windows.Forms.MenuItem mmiConnection;
        private System.Windows.Forms.MenuItem mmiOpenClose;
        private System.Windows.Forms.MenuItem mmiSend;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem mmiWriteLog;
        private System.Windows.Forms.MenuItem menuItem5;
        private System.Windows.Forms.MenuItem menuItem7;
        private System.Windows.Forms.MenuItem miSaveLog;
        private System.Windows.Forms.MenuItem miWriteLog;
        private System.Windows.Forms.MenuItem mmiDefaultModel;
        private SplitContainer splitContainer1;
        private ListBox lstResponse;
        private LinkLabel lblSgart;
        private Button btnHelp;
        private Button btnSend;
        private TextBox txtCommand;
        private Label lblPwd;
        private TextBox txtPassword;
        private TextBox txtUser;
        private Label lblUser;
        private IContainer components;

        public FormWinSMTP()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormWinSMTP));
            this.panelTop = new System.Windows.Forms.Panel();
            this.lblPwd = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.lblUser = new System.Windows.Forms.Label();
            this.btnOpen = new System.Windows.Forms.Button();
            this.txtServerPort = new System.Windows.Forms.TextBox();
            this.lblServerPort = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtServerSMTP = new System.Windows.Forms.TextBox();
            this.cmLstResponse = new System.Windows.Forms.ContextMenu();
            this.miCopy = new System.Windows.Forms.MenuItem();
            this.miClear = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.miSaveLog = new System.Windows.Forms.MenuItem();
            this.menuItem7 = new System.Windows.Forms.MenuItem();
            this.miWriteLog = new System.Windows.Forms.MenuItem();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.mmiFile = new System.Windows.Forms.MenuItem();
            this.mmiDefaultModel = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.mmiOpenModel = new System.Windows.Forms.MenuItem();
            this.mmiSaveModel = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.mmiExit = new System.Windows.Forms.MenuItem();
            this.mmiConnection = new System.Windows.Forms.MenuItem();
            this.mmiOpenClose = new System.Windows.Forms.MenuItem();
            this.mmiSend = new System.Windows.Forms.MenuItem();
            this.mmiLog = new System.Windows.Forms.MenuItem();
            this.mmiCopy = new System.Windows.Forms.MenuItem();
            this.mmiClear = new System.Windows.Forms.MenuItem();
            this.menuItem9 = new System.Windows.Forms.MenuItem();
            this.mmiSaveLog = new System.Windows.Forms.MenuItem();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.mmiWriteLog = new System.Windows.Forms.MenuItem();
            this.mmiHelp = new System.Windows.Forms.MenuItem();
            this.mmiAbout = new System.Windows.Forms.MenuItem();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lstResponse = new System.Windows.Forms.ListBox();
            this.lblSgart = new System.Windows.Forms.LinkLabel();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnSend = new System.Windows.Forms.Button();
            this.txtCommand = new System.Windows.Forms.TextBox();
            this.panelTop.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.lblPwd);
            this.panelTop.Controls.Add(this.txtPassword);
            this.panelTop.Controls.Add(this.txtUser);
            this.panelTop.Controls.Add(this.lblUser);
            this.panelTop.Controls.Add(this.btnOpen);
            this.panelTop.Controls.Add(this.txtServerPort);
            this.panelTop.Controls.Add(this.lblServerPort);
            this.panelTop.Controls.Add(this.label1);
            this.panelTop.Controls.Add(this.txtServerSMTP);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(384, 58);
            this.panelTop.TabIndex = 10;
            // 
            // lblPwd
            // 
            this.lblPwd.AutoSize = true;
            this.lblPwd.Location = new System.Drawing.Point(163, 35);
            this.lblPwd.Name = "lblPwd";
            this.lblPwd.Size = new System.Drawing.Size(53, 13);
            this.lblPwd.TabIndex = 10;
            this.lblPwd.Text = "Password";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(222, 32);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(90, 20);
            this.txtPassword.TabIndex = 3;
            this.txtPassword.UseSystemPasswordChar = true;
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(48, 32);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(108, 20);
            this.txtUser.TabIndex = 2;
            // 
            // lblUser
            // 
            this.lblUser.AutoSize = true;
            this.lblUser.Location = new System.Drawing.Point(3, 35);
            this.lblUser.Name = "lblUser";
            this.lblUser.Size = new System.Drawing.Size(29, 13);
            this.lblUser.TabIndex = 7;
            this.lblUser.Text = "User";
            // 
            // btnOpen
            // 
            this.btnOpen.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnOpen.Location = new System.Drawing.Point(320, 0);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(64, 58);
            this.btnOpen.TabIndex = 4;
            this.btnOpen.Text = "Open";
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // txtServerPort
            // 
            this.txtServerPort.Location = new System.Drawing.Point(264, 8);
            this.txtServerPort.Name = "txtServerPort";
            this.txtServerPort.Size = new System.Drawing.Size(48, 20);
            this.txtServerPort.TabIndex = 1;
            // 
            // lblServerPort
            // 
            this.lblServerPort.Location = new System.Drawing.Point(230, 11);
            this.lblServerPort.Name = "lblServerPort";
            this.lblServerPort.Size = new System.Drawing.Size(32, 16);
            this.lblServerPort.TabIndex = 6;
            this.lblServerPort.Text = "Port";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(3, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "Server";
            // 
            // txtServerSMTP
            // 
            this.txtServerSMTP.Location = new System.Drawing.Point(48, 7);
            this.txtServerSMTP.Name = "txtServerSMTP";
            this.txtServerSMTP.Size = new System.Drawing.Size(176, 20);
            this.txtServerSMTP.TabIndex = 0;
            // 
            // cmLstResponse
            // 
            this.cmLstResponse.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.miCopy,
            this.miClear,
            this.menuItem5,
            this.miSaveLog,
            this.menuItem7,
            this.miWriteLog});
            // 
            // miCopy
            // 
            this.miCopy.Index = 0;
            this.miCopy.Text = "&Copy";
            this.miCopy.Click += new System.EventHandler(this.miCopy_Click);
            // 
            // miClear
            // 
            this.miClear.Index = 1;
            this.miClear.Text = "Clear";
            this.miClear.Click += new System.EventHandler(this.miClear_Click);
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 2;
            this.menuItem5.Text = "-";
            // 
            // miSaveLog
            // 
            this.miSaveLog.Index = 3;
            this.miSaveLog.Text = "&Save";
            this.miSaveLog.Click += new System.EventHandler(this.miSaveLog_Click);
            // 
            // menuItem7
            // 
            this.menuItem7.Index = 4;
            this.menuItem7.Text = "-";
            // 
            // miWriteLog
            // 
            this.miWriteLog.Index = 5;
            this.miWriteLog.Text = "&Write to file";
            this.miWriteLog.Click += new System.EventHandler(this.miWriteLog_Click);
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mmiFile,
            this.mmiConnection,
            this.mmiLog,
            this.mmiHelp});
            // 
            // mmiFile
            // 
            this.mmiFile.Index = 0;
            this.mmiFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mmiDefaultModel,
            this.menuItem3,
            this.mmiOpenModel,
            this.mmiSaveModel,
            this.menuItem4,
            this.mmiExit});
            this.mmiFile.Text = "&File";
            // 
            // mmiDefaultModel
            // 
            this.mmiDefaultModel.Index = 0;
            this.mmiDefaultModel.Text = "&Default";
            this.mmiDefaultModel.Click += new System.EventHandler(this.mmiDefaultModel_Click);
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 1;
            this.menuItem3.Text = "-";
            // 
            // mmiOpenModel
            // 
            this.mmiOpenModel.Index = 2;
            this.mmiOpenModel.Text = "&Open";
            this.mmiOpenModel.Click += new System.EventHandler(this.mmiOpenModel_Click);
            // 
            // mmiSaveModel
            // 
            this.mmiSaveModel.Index = 3;
            this.mmiSaveModel.Text = "&Save";
            this.mmiSaveModel.Click += new System.EventHandler(this.mmiSaveModel_Click);
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 4;
            this.menuItem4.Text = "-";
            // 
            // mmiExit
            // 
            this.mmiExit.Index = 5;
            this.mmiExit.Text = "E&xit";
            this.mmiExit.Click += new System.EventHandler(this.mmiExit_Click);
            // 
            // mmiConnection
            // 
            this.mmiConnection.Index = 1;
            this.mmiConnection.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mmiOpenClose,
            this.mmiSend});
            this.mmiConnection.Text = "&Connection";
            this.mmiConnection.Popup += new System.EventHandler(this.mmiConnection_Popup);
            // 
            // mmiOpenClose
            // 
            this.mmiOpenClose.Index = 0;
            this.mmiOpenClose.Text = "&Open";
            this.mmiOpenClose.Click += new System.EventHandler(this.mmiOpenClose_Click);
            // 
            // mmiSend
            // 
            this.mmiSend.Index = 1;
            this.mmiSend.Text = "&Send";
            this.mmiSend.Click += new System.EventHandler(this.mmiSend_Click);
            // 
            // mmiLog
            // 
            this.mmiLog.Index = 2;
            this.mmiLog.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mmiCopy,
            this.mmiClear,
            this.menuItem9,
            this.mmiSaveLog,
            this.menuItem1,
            this.mmiWriteLog});
            this.mmiLog.Text = "&Log";
            // 
            // mmiCopy
            // 
            this.mmiCopy.Index = 0;
            this.mmiCopy.Text = "&Copy";
            this.mmiCopy.Click += new System.EventHandler(this.mmiCopy_Click);
            // 
            // mmiClear
            // 
            this.mmiClear.Index = 1;
            this.mmiClear.Text = "Clear";
            this.mmiClear.Click += new System.EventHandler(this.mmiClear_Click);
            // 
            // menuItem9
            // 
            this.menuItem9.Index = 2;
            this.menuItem9.Text = "-";
            // 
            // mmiSaveLog
            // 
            this.mmiSaveLog.Index = 3;
            this.mmiSaveLog.Text = "&Save";
            this.mmiSaveLog.Click += new System.EventHandler(this.mmiSaveLog_Click);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 4;
            this.menuItem1.Text = "-";
            // 
            // mmiWriteLog
            // 
            this.mmiWriteLog.Index = 5;
            this.mmiWriteLog.Text = "&Write to file";
            this.mmiWriteLog.Click += new System.EventHandler(this.mmiWriteLog_Click);
            // 
            // mmiHelp
            // 
            this.mmiHelp.Index = 3;
            this.mmiHelp.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mmiAbout});
            this.mmiHelp.Text = "&Help";
            // 
            // mmiAbout
            // 
            this.mmiAbout.Index = 0;
            this.mmiAbout.Text = "&About";
            this.mmiAbout.Click += new System.EventHandler(this.mmiAbout_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 58);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.lstResponse);
            this.splitContainer1.Panel1MinSize = 80;
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.lblSgart);
            this.splitContainer1.Panel2.Controls.Add(this.btnHelp);
            this.splitContainer1.Panel2.Controls.Add(this.btnSend);
            this.splitContainer1.Panel2.Controls.Add(this.txtCommand);
            this.splitContainer1.Panel2MinSize = 80;
            this.splitContainer1.Size = new System.Drawing.Size(384, 235);
            this.splitContainer1.SplitterDistance = 117;
            this.splitContainer1.SplitterWidth = 6;
            this.splitContainer1.TabIndex = 12;
            // 
            // lstResponse
            // 
            this.lstResponse.ContextMenu = this.cmLstResponse;
            this.lstResponse.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstResponse.HorizontalScrollbar = true;
            this.lstResponse.Location = new System.Drawing.Point(0, 0);
            this.lstResponse.Name = "lstResponse";
            this.lstResponse.Size = new System.Drawing.Size(384, 117);
            this.lstResponse.TabIndex = 5;
            // 
            // lblSgart
            // 
            this.lblSgart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSgart.Location = new System.Drawing.Point(320, 93);
            this.lblSgart.Name = "lblSgart";
            this.lblSgart.Size = new System.Drawing.Size(64, 16);
            this.lblSgart.TabIndex = 9;
            this.lblSgart.TabStop = true;
            this.lblSgart.Text = "by Sgart.it";
            this.lblSgart.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblSgart.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblSgart_LinkClicked);
            // 
            // btnHelp
            // 
            this.btnHelp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnHelp.Location = new System.Drawing.Point(320, 42);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(64, 24);
            this.btnHelp.TabIndex = 8;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnSend
            // 
            this.btnSend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSend.Location = new System.Drawing.Point(320, 2);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(64, 32);
            this.btnSend.TabIndex = 7;
            this.btnSend.Text = "Send";
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // txtCommand
            // 
            this.txtCommand.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCommand.Location = new System.Drawing.Point(0, 3);
            this.txtCommand.MaxLength = 0;
            this.txtCommand.Multiline = true;
            this.txtCommand.Name = "txtCommand";
            this.txtCommand.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtCommand.Size = new System.Drawing.Size(312, 106);
            this.txtCommand.TabIndex = 6;
            // 
            // FormWinSMTP
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(384, 293);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.panelTop);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.mainMenu1;
            this.MinimumSize = new System.Drawing.Size(400, 300);
            this.Name = "FormWinSMTP";
            this.Text = "WinSMTPTest";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            flagSilent = false;
            commandTemplate = "";
            serverSMTP = "127.0.0.1";
            serverSMTPPort = "25";
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i].ToLower();
                if (arg.CompareTo("/h") == 0 || arg.CompareTo("/?") == 0)
                {
                    MessageBox.Show(
                        String.Format(ReadResourceString("helpCommand"), Application.ProductName, Application.ProductVersion),
                        Application.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Question);

                    return;
                }
                else if (arg.CompareTo("/s") == 0)
                {
                    flagSilent = true;
                }
                else if (arg.IndexOf(':') > 0)
                {
                    string[] s = arg.Split(':');
                    if (s[0].Length > 0)
                        serverSMTP = s[0];
                    if (s[1].Length > 0)
                        serverSMTPPort = s[1];
                }
                else
                {
                    commandTemplate = ReadCommandTemplateFromFile(arg);
                }
            }
            Application.Run(new FormWinSMTP());
        }

        private static string ReadCommandTemplateFromFile(string fn)
        {
            return ReadFile(fn, true);
        }

        private void btnOpen_Click(object sender, System.EventArgs e)
        {
            if (cTcp.IsConnect == false)
            {
                //ClearMessage();
                string msg = cTcp.Open(txtServerSMTP.Text, txtServerPort.Text);
                if (msg.Length > 0)
                {
                    AddMessage(msg);
                    //SendCommand("HELP");
                }
                else
                {
                    AddMessage("__OPEN_FAILED__");
                }
            }
            else
            {
                SendCommand("QUIT");
                if (cTcp.IsConnect == true)
                {
                    // se la connessione è rimasta aperta, forzo la chiusura
                    cTcp.Close();
                }
            }
            RedrawOpenCloseButton();
        }

        private void RedrawOpenCloseButton()
        {
            if (cTcp.IsConnect == false)
            {
                //panelBottom.Enabled = false;
                txtCommand.Enabled = false;
                btnSend.Enabled = false;
                btnHelp.Enabled = false;
                txtServerSMTP.Enabled = true;
                txtServerPort.Enabled = true;
                btnOpen.Text = "Open";
            }
            else
            {
                //panelBottom.Enabled = true;
                txtCommand.Enabled = true;
                btnSend.Enabled = true;
                btnHelp.Enabled = true;
                txtServerSMTP.Enabled = false;
                txtServerPort.Enabled = false;
                btnOpen.Text = "Close";
            }
            mmiOpenClose.Text = btnOpen.Text;
        }

        private void btnSend_Click(object sender, System.EventArgs e)
        {
            string[] ln = (txtCommand.Text + "\n").Replace("\r", "").Split('\n');
            string mailFrom = "";
            string rcptTo = "";
            string n = "";
            foreach (string s in ln)
            {
                n = s.Trim();
                if (n.Length >= 0)
                {
                    if (n.ToLower().StartsWith("mail from:") == true)
                    {
                        mailFrom += n.Substring(10) + ";";
                    }
                    if (n.ToLower().StartsWith("rcpt to:") == true)
                    {
                        rcptTo += n.Substring(8) + ";";
                    }
                    DateTime now = DateTime.Now;
                    n = n.Replace("{%now%}", now.ToLongDateString() + " " + now.ToLongTimeString() + " " + now.Millisecond + " ms");
                    n = n.Replace("{%count%}", sendCount.ToString());
                    n = n.Replace("{%from%}", mailFrom);
                    n = n.Replace("{%to%}", rcptTo);
                    n = n.Replace("{%server%}", txtServerSMTP.Text + ":" + txtServerPort.Text);
                    n = n.Replace("{%manifest%}", cTcp.ServerManifest);
                    n = n.Replace("{%user%}", EncodeToBase64(txtUser.Text));
                    n = n.Replace("{%pwd%}", EncodeToBase64(txtPassword.Text));
                    n = n.Replace("{%password%}", EncodeToBase64(txtPassword.Text));
                    SendCommand(n);
                    if (cTcp.IsConnect == false)
                        break;
                }
            }
            sendCount++;
            RedrawOpenCloseButton();
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {
            /*
            ResourceManager rm = new System.Resources.ResourceManager(
                "WinSMTP.Resource", 
                Assembly.GetExecutingAssembly()
                );
            */
            //panelBottom.Enabled = false;
            txtCommand.Enabled = false;
            btnSend.Enabled = false;
            btnHelp.Enabled = false;
            txtServerSMTP.Enabled = true;
            txtServerPort.Enabled = true;
            if (commandTemplate.Length == 0)
            {
                //commandTemplate = rm.GetString("defaultModel");
                commandTemplate = ReadResourceString("defaultModel");
            }
            txtCommand.Text = commandTemplate;
            //helpAboutMessage = rm.GetString("helpAbout");
            string s = ReadResourceString("helpAbout");
            //string s = "{0} {1} {0}";
            helpAboutMessage = String.Format(s, Application.ProductName, Application.ProductVersion);

            txtServerSMTP.Text = serverSMTP;
            txtServerPort.Text = serverSMTPPort;

            cTcp = new ClassTCP(String.Format("\n\r\n\r{0} by http://www.sgart.it - v. {1} (2005)", Application.ProductName, Application.ProductVersion));
        }


        private void lblSgart_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.sgart.it/go.asp?prg=WinSMTP");
        }


        private void ClearMessage()
        {
            lstResponse.Items.Clear();
        }

        private void AddMessage(string msg)
        {
            StreamWriter wr = null;
            //controlla se nel log ho più di 1000 item ne elimino 50 alla volta
            if (lstResponse.Items.Count > 1000)
            {
                for (int i = 0; i < 50; i++)
                    lstResponse.Items.RemoveAt(0);
            }

            if (msg.EndsWith("\r\n") == false)
            {
                msg = String.Concat(msg, "\n");
            }
            string[] ln = msg.Replace("\r", "").Split('\n');
            ln[ln.Length - 1] = null;
            DateTime now = DateTime.Now;
            string nowString = "[" + now.Hour + ":" + now.Minute + ":" + now.Second + "." + now.Millisecond + "] ";
            if (writeToLogFile == true)
            {
                try
                {
                    wr = File.AppendText(FILE_TO_LOG);
                }
                catch
                {
                    writeToLogFile = false;
                    MessageBox.Show("Can't write to log file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            foreach (string s in ln)
            {
                if (s != null)
                {
                    if (writeToLogFile == true)
                        try
                        {
                            wr.WriteLine(nowString + s);
                        }
                        catch
                        {
                        }
                    lstResponse.Items.Add(nowString + s);
                }
            }
            lstResponse.SelectedIndex = lstResponse.Items.Count - 1;
            if (writeToLogFile == true)
                try
                {
                    wr.Close();
                }
                catch
                {
                }
        }


        private void SendCommand(string cmd)
        {
            string r;
            AddMessage(cmd);
            r = cTcp.SendCommand(cmd);
            if (r.Length > 0)
            {
                AddMessage(r);
            }
        }

        private void miCopy_Click(object sender, System.EventArgs e)
        {
            string s = "";
            foreach (string i in lstResponse.Items)
                s += i + "\r\n";
            Clipboard.SetDataObject(s);
        }

        private void miClear_Click(object sender, System.EventArgs e)
        {
            lstResponse.Items.Clear();
        }
        private void mmiCopy_Click(object sender, System.EventArgs e)
        {
            miCopy_Click(sender, e);
        }
        private void miSaveLog_Click(object sender, System.EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Log files (*.log)|*.log|All files (*.*)|*.*";

            if (saveDialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            string s = "";
            foreach (string i in lstResponse.Items)
                s += i + "\r\n";
            WriteFile(saveDialog.FileName, s);

        }
        private void miWriteLog_Click(object sender, System.EventArgs e)
        {
            writeToLogFile = !writeToLogFile;
            mmiWriteLog.Checked = writeToLogFile;
        }

        private void mmiClear_Click(object sender, System.EventArgs e)
        {
            miClear_Click(sender, e);
        }

        private void mmiAbout_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show(helpAboutMessage, "About",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void mmiOpenClose_Click(object sender, System.EventArgs e)
        {
            btnOpen_Click(sender, e);
        }


        private void mmiSend_Click(object sender, System.EventArgs e)
        {
            btnSend_Click(sender, e);
        }

        private void mmiConnection_Popup(object sender, System.EventArgs e)
        {
            mmiSend.Enabled = btnSend.Enabled;
        }

        private void mmiWriteLog_Click(object sender, System.EventArgs e)
        {
            miWriteLog_Click(sender, e);
        }

        private void mmiExit_Click(object sender, System.EventArgs e)
        {
            Application.Exit();
        }

        private void mmiSaveLog_Click(object sender, System.EventArgs e)
        {
            miSaveLog_Click(sender, e);
        }

        private void mmiDefaultModel_Click(object sender, System.EventArgs e)
        {
            txtCommand.Text = ReadResourceString("defaultModel");
            /*
            ResourceManager rm = new System.Resources.ResourceManager(
                "WinSMTP.Resource", 
                Assembly.GetExecutingAssembly()
                );
            txtCommand.Text = rm.GetString("defaultModel");
            */
        }

        private void mmiOpenModel_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Command files (*.txt)|*.txt|All files (*.*)|*.*";

            if (openDialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            txtCommand.Text = ReadFile(openDialog.FileName, false);
        }


        private static string ReadFile(string fn, bool noDialog)
        {
            string s;
            try
            {
                StreamReader re = File.OpenText(fn);
                s = re.ReadToEnd();
                re.Close();
            }
            catch
            {
                if (noDialog != true)
                    MessageBox.Show("Can't read file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                s = "";
            }
            return s;
        }

        private static void WriteFile(string fn, string s)
        {
            try
            {
                StreamWriter wr = File.CreateText(fn);
                wr.Write(s);
                wr.Close();
            }
            catch
            {
                MessageBox.Show("Can't write file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }

        private void mmiSaveModel_Click(object sender, System.EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Command files (*.txt)|*.txt|All files (*.*)|*.*";

            if (saveDialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            WriteFile(saveDialog.FileName, txtCommand.Text);
        }

        private static string ReadResourceString(string rn)
        {
            ResourceManager rm = new System.Resources.ResourceManager(
                "WinSMTPTest.Resource",
                Assembly.GetExecutingAssembly()
                );

            return rm.GetString(rn);
        }

        private void btnHelp_Click(object sender, System.EventArgs e)
        {
            SendCommand("HELP");
        }

        static public string EncodeToBase64(string text)
        {
            byte[] textBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(text);
            return System.Convert.ToBase64String(textBytes);
        }
    }
}
