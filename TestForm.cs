using System;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using System.Data;
using System.Collections.Generic;
using OpenPop.Mime.Header;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OpenPop.Mime;
using Message = OpenPop.Mime.Message;
using System.Data.SqlClient;
using System.Drawing;
using OpenPop.TestApplication;

namespace cs499
{
    public class TestForm : Form
    {


        private Label lblUsername;
        private readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
        private readonly Pop3Client pop3Client;
        private Button connectButton;
        private Panel attachmentPanel;
        private ContextMenu contextMenuMessages;
        private DataGrid gridHeaders;
        private Panel panelMiddle;
        private Label labelPassword;
        private TextBox popServerTextBox;
        private TextBox passwordTextBox;
        private TextBox portTextBox;
        private MenuItem Tasking;
        private TextBox totalMessagesTextBox;
        private Panel panelMessageBody;
        private Panel panelMessagesView;
        private TextBox messageTextBox;
        private Label lblAttachments;
        private Label lblMessageBody;
        private Label lblMessageNumber;
        private Label labelTotalMessages;
        private SaveFileDialog saveFile;
        private TextBox loginTextBox;
        private Label labelServerAddress;
        private Label labelServerPort;
        private TreeView listAttachments;
        private TreeView listMessages;
        private MenuItem menuViewSource;
        private Panel panelTop;
        private Panel panelProperties;
        private ProgressBar progressBar;
        private TextBox databaseNameTxt;
        private TextBox serverNameTxt;
        private Label label3;
        private Label label2;
        private Label label1;
        private ComboBox loginTypeCb;
        private TextBox usernameTxt;
        private Label label4;
        private Label label5;
        private TextBox passwordTxt;
        private Button userAccountsBtn;
        private MenuItem menuItem1;
        private MenuItem menuItem2;
        private CheckBox useSslCheckBox;

        private TestForm()
        {

            InitializeComponent();

            pop3Client = new Pop3Client();

            string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string file = Path.Combine(myDocs, "OpenPopLogin.txt");
            if (File.Exists(file))
            {
                using (StreamReader reader = new StreamReader(File.OpenRead(file)))
                {
                    popServerTextBox.Text = reader.ReadLine(); // Hostname
                    portTextBox.Text = reader.ReadLine(); // Port
                    useSslCheckBox.Checked = bool.Parse(reader.ReadLine() ?? "true"); // Whether to use SSL or not
                    loginTextBox.Text = reader.ReadLine(); // Username
                    passwordTextBox.Text = reader.ReadLine(); // Password
                }
            }
            attachmentPanel.Visible = true;
        }

        #region Windows Form Designer generated code
        /// <summary>
        ///   Required method for Designer support - do not modify
        ///   the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TestForm));
            this.panelTop = new System.Windows.Forms.Panel();
            this.totalMessagesTextBox = new System.Windows.Forms.TextBox();
            this.labelTotalMessages = new System.Windows.Forms.Label();
            this.useSslCheckBox = new System.Windows.Forms.CheckBox();
            this.labelPassword = new System.Windows.Forms.Label();
            this.passwordTextBox = new System.Windows.Forms.TextBox();
            this.lblUsername = new System.Windows.Forms.Label();
            this.loginTextBox = new System.Windows.Forms.TextBox();
            this.connectButton = new System.Windows.Forms.Button();
            this.labelServerPort = new System.Windows.Forms.Label();
            this.portTextBox = new System.Windows.Forms.TextBox();
            this.labelServerAddress = new System.Windows.Forms.Label();
            this.popServerTextBox = new System.Windows.Forms.TextBox();
            this.panelProperties = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.userAccountsBtn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.gridHeaders = new System.Windows.Forms.DataGrid();
            this.passwordTxt = new System.Windows.Forms.TextBox();
            this.loginTypeCb = new System.Windows.Forms.ComboBox();
            this.serverNameTxt = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.usernameTxt = new System.Windows.Forms.TextBox();
            this.databaseNameTxt = new System.Windows.Forms.TextBox();
            this.panelMiddle = new System.Windows.Forms.Panel();
            this.panelMessageBody = new System.Windows.Forms.Panel();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.messageTextBox = new System.Windows.Forms.TextBox();
            this.lblMessageBody = new System.Windows.Forms.Label();
            this.panelMessagesView = new System.Windows.Forms.Panel();
            this.listMessages = new System.Windows.Forms.TreeView();
            this.contextMenuMessages = new System.Windows.Forms.ContextMenu();
            this.menuViewSource = new System.Windows.Forms.MenuItem();
            this.Tasking = new System.Windows.Forms.MenuItem();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.lblMessageNumber = new System.Windows.Forms.Label();
            this.attachmentPanel = new System.Windows.Forms.Panel();
            this.listAttachments = new System.Windows.Forms.TreeView();
            this.lblAttachments = new System.Windows.Forms.Label();
            this.saveFile = new System.Windows.Forms.SaveFileDialog();
            this.panelTop.SuspendLayout();
            this.panelProperties.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridHeaders)).BeginInit();
            this.panelMiddle.SuspendLayout();
            this.panelMessageBody.SuspendLayout();
            this.panelMessagesView.SuspendLayout();
            this.attachmentPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTop
            // 
            this.panelTop.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panelTop.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelTop.Controls.Add(this.totalMessagesTextBox);
            this.panelTop.Controls.Add(this.labelTotalMessages);
            this.panelTop.Controls.Add(this.useSslCheckBox);
            this.panelTop.Controls.Add(this.labelPassword);
            this.panelTop.Controls.Add(this.passwordTextBox);
            this.panelTop.Controls.Add(this.lblUsername);
            this.panelTop.Controls.Add(this.loginTextBox);
            this.panelTop.Controls.Add(this.connectButton);
            this.panelTop.Controls.Add(this.labelServerPort);
            this.panelTop.Controls.Add(this.portTextBox);
            this.panelTop.Controls.Add(this.labelServerAddress);
            this.panelTop.Controls.Add(this.popServerTextBox);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(903, 64);
            this.panelTop.TabIndex = 0;
            // 
            // totalMessagesTextBox
            // 
            this.totalMessagesTextBox.Location = new System.Drawing.Point(13, 30);
            this.totalMessagesTextBox.Name = "totalMessagesTextBox";
            this.totalMessagesTextBox.Size = new System.Drawing.Size(44, 20);
            this.totalMessagesTextBox.TabIndex = 10;
            // 
            // labelTotalMessages
            // 
            this.labelTotalMessages.Location = new System.Drawing.Point(10, 7);
            this.labelTotalMessages.Name = "labelTotalMessages";
            this.labelTotalMessages.Size = new System.Drawing.Size(100, 23);
            this.labelTotalMessages.TabIndex = 11;
            this.labelTotalMessages.Text = "Total Messages";
            // 
            // useSslCheckBox
            // 
            this.useSslCheckBox.AutoSize = true;
            this.useSslCheckBox.Checked = true;
            this.useSslCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.useSslCheckBox.Location = new System.Drawing.Point(161, 28);
            this.useSslCheckBox.Name = "useSslCheckBox";
            this.useSslCheckBox.Size = new System.Drawing.Size(68, 17);
            this.useSslCheckBox.TabIndex = 4;
            this.useSslCheckBox.Text = "Use SSL";
            this.useSslCheckBox.UseVisualStyleBackColor = true;
            // 
            // labelPassword
            // 
            this.labelPassword.Location = new System.Drawing.Point(407, 32);
            this.labelPassword.Name = "labelPassword";
            this.labelPassword.Size = new System.Drawing.Size(64, 23);
            this.labelPassword.TabIndex = 8;
            this.labelPassword.Text = "Password";
            // 
            // passwordTextBox
            // 
            this.passwordTextBox.Location = new System.Drawing.Point(477, 30);
            this.passwordTextBox.Name = "passwordTextBox";
            this.passwordTextBox.PasswordChar = '*';
            this.passwordTextBox.Size = new System.Drawing.Size(128, 20);
            this.passwordTextBox.TabIndex = 2;
            // 
            // lblUsername
            // 
            this.lblUsername.Location = new System.Drawing.Point(407, 6);
            this.lblUsername.Name = "lblUsername";
            this.lblUsername.Size = new System.Drawing.Size(64, 23);
            this.lblUsername.TabIndex = 6;
            this.lblUsername.Text = "Email";
            // 
            // loginTextBox
            // 
            this.loginTextBox.Location = new System.Drawing.Point(477, 3);
            this.loginTextBox.Name = "loginTextBox";
            this.loginTextBox.Size = new System.Drawing.Size(128, 20);
            this.loginTextBox.TabIndex = 1;
            this.loginTextBox.Text = "nedimp_97@hotmail.com";
            this.loginTextBox.TextChanged += new System.EventHandler(this.loginTextBox_TextChanged);
            // 
            // connectButton
            // 
            this.connectButton.Location = new System.Drawing.Point(631, 1);
            this.connectButton.Name = "connectButton";
            this.connectButton.Size = new System.Drawing.Size(111, 54);
            this.connectButton.TabIndex = 5;
            this.connectButton.Text = "Connect";
            this.connectButton.Click += new System.EventHandler(this.ConnectAndRetrieveButtonClick);
            // 
            // labelServerPort
            // 
            this.labelServerPort.Location = new System.Drawing.Point(239, 32);
            this.labelServerPort.Name = "labelServerPort";
            this.labelServerPort.Size = new System.Drawing.Size(31, 23);
            this.labelServerPort.TabIndex = 3;
            this.labelServerPort.Text = "Port";
            // 
            // portTextBox
            // 
            this.portTextBox.Location = new System.Drawing.Point(273, 29);
            this.portTextBox.Name = "portTextBox";
            this.portTextBox.Size = new System.Drawing.Size(128, 20);
            this.portTextBox.TabIndex = 3;
            this.portTextBox.Text = "995";
            // 
            // labelServerAddress
            // 
            this.labelServerAddress.Location = new System.Drawing.Point(158, 2);
            this.labelServerAddress.Name = "labelServerAddress";
            this.labelServerAddress.Size = new System.Drawing.Size(112, 23);
            this.labelServerAddress.TabIndex = 1;
            this.labelServerAddress.Text = "POP Server Address";
            // 
            // popServerTextBox
            // 
            this.popServerTextBox.Location = new System.Drawing.Point(273, 3);
            this.popServerTextBox.Name = "popServerTextBox";
            this.popServerTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.popServerTextBox.Size = new System.Drawing.Size(128, 20);
            this.popServerTextBox.TabIndex = 0;
            this.popServerTextBox.Text = "pop-mail.outlook.com";
            // 
            // panelProperties
            // 
            this.panelProperties.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panelProperties.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelProperties.Controls.Add(this.label3);
            this.panelProperties.Controls.Add(this.label1);
            this.panelProperties.Controls.Add(this.label5);
            this.panelProperties.Controls.Add(this.userAccountsBtn);
            this.panelProperties.Controls.Add(this.label2);
            this.panelProperties.Controls.Add(this.gridHeaders);
            this.panelProperties.Controls.Add(this.passwordTxt);
            this.panelProperties.Controls.Add(this.loginTypeCb);
            this.panelProperties.Controls.Add(this.serverNameTxt);
            this.panelProperties.Controls.Add(this.label4);
            this.panelProperties.Controls.Add(this.usernameTxt);
            this.panelProperties.Controls.Add(this.databaseNameTxt);
            this.panelProperties.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelProperties.Location = new System.Drawing.Point(0, 421);
            this.panelProperties.Name = "panelProperties";
            this.panelProperties.Size = new System.Drawing.Size(903, 299);
            this.panelProperties.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 11);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(162, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "SQL Login or Win Authentication";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Server Name";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(343, 60);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "Password";
            this.label5.Visible = false;
            // 
            // userAccountsBtn
            // 
            this.userAccountsBtn.Location = new System.Drawing.Point(763, 69);
            this.userAccountsBtn.Name = "userAccountsBtn";
            this.userAccountsBtn.Size = new System.Drawing.Size(128, 37);
            this.userAccountsBtn.TabIndex = 12;
            this.userAccountsBtn.Text = "Manage User Accounts";
            this.userAccountsBtn.UseVisualStyleBackColor = true;
            this.userAccountsBtn.Click += new System.EventHandler(this.userAccountsBtn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(122, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Database Name";
            // 
            // gridHeaders
            // 
            this.gridHeaders.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridHeaders.BackgroundColor = System.Drawing.Color.LightSlateGray;
            this.gridHeaders.DataMember = "";
            this.gridHeaders.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.gridHeaders.Location = new System.Drawing.Point(0, 117);
            this.gridHeaders.Name = "gridHeaders";
            this.gridHeaders.PreferredColumnWidth = 400;
            this.gridHeaders.ReadOnly = true;
            this.gridHeaders.Size = new System.Drawing.Size(899, 182);
            this.gridHeaders.TabIndex = 3;
            // 
            // passwordTxt
            // 
            this.passwordTxt.Location = new System.Drawing.Point(346, 78);
            this.passwordTxt.Name = "passwordTxt";
            this.passwordTxt.Size = new System.Drawing.Size(100, 20);
            this.passwordTxt.TabIndex = 15;
            this.passwordTxt.Visible = false;
            // 
            // loginTypeCb
            // 
            this.loginTypeCb.FormattingEnabled = true;
            this.loginTypeCb.Items.AddRange(new object[] {
            "Windows Authentication",
            "Sql Login"});
            this.loginTypeCb.Location = new System.Drawing.Point(13, 27);
            this.loginTypeCb.Name = "loginTypeCb";
            this.loginTypeCb.Size = new System.Drawing.Size(166, 21);
            this.loginTypeCb.TabIndex = 12;
            this.loginTypeCb.Text = "Windows Authentication";
            this.loginTypeCb.SelectedIndexChanged += new System.EventHandler(this.loginTypeCb_SelectedIndexChanged);
            // 
            // serverNameTxt
            // 
            this.serverNameTxt.Location = new System.Drawing.Point(10, 78);
            this.serverNameTxt.Name = "serverNameTxt";
            this.serverNameTxt.Size = new System.Drawing.Size(100, 20);
            this.serverNameTxt.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(228, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Username";
            this.label4.Visible = false;
            // 
            // usernameTxt
            // 
            this.usernameTxt.Location = new System.Drawing.Point(231, 78);
            this.usernameTxt.Name = "usernameTxt";
            this.usernameTxt.Size = new System.Drawing.Size(100, 20);
            this.usernameTxt.TabIndex = 14;
            this.usernameTxt.Visible = false;
            // 
            // databaseNameTxt
            // 
            this.databaseNameTxt.Location = new System.Drawing.Point(124, 78);
            this.databaseNameTxt.Name = "databaseNameTxt";
            this.databaseNameTxt.Size = new System.Drawing.Size(100, 20);
            this.databaseNameTxt.TabIndex = 11;
            // 
            // panelMiddle
            // 
            this.panelMiddle.Controls.Add(this.panelMessageBody);
            this.panelMiddle.Controls.Add(this.panelMessagesView);
            this.panelMiddle.Controls.Add(this.attachmentPanel);
            this.panelMiddle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMiddle.Location = new System.Drawing.Point(0, 64);
            this.panelMiddle.Name = "panelMiddle";
            this.panelMiddle.Size = new System.Drawing.Size(903, 357);
            this.panelMiddle.TabIndex = 2;
            // 
            // panelMessageBody
            // 
            this.panelMessageBody.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.panelMessageBody.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelMessageBody.Controls.Add(this.progressBar);
            this.panelMessageBody.Controls.Add(this.messageTextBox);
            this.panelMessageBody.Controls.Add(this.lblMessageBody);
            this.panelMessageBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMessageBody.Location = new System.Drawing.Point(281, 0);
            this.panelMessageBody.Name = "panelMessageBody";
            this.panelMessageBody.Size = new System.Drawing.Size(414, 357);
            this.panelMessageBody.TabIndex = 6;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(7, 329);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(394, 18);
            this.progressBar.TabIndex = 11;
            // 
            // messageTextBox
            // 
            this.messageTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.messageTextBox.Location = new System.Drawing.Point(7, 22);
            this.messageTextBox.MaxLength = 999999999;
            this.messageTextBox.Multiline = true;
            this.messageTextBox.Name = "messageTextBox";
            this.messageTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.messageTextBox.Size = new System.Drawing.Size(394, 300);
            this.messageTextBox.TabIndex = 9;
            // 
            // lblMessageBody
            // 
            this.lblMessageBody.Location = new System.Drawing.Point(8, 8);
            this.lblMessageBody.Name = "lblMessageBody";
            this.lblMessageBody.Size = new System.Drawing.Size(136, 16);
            this.lblMessageBody.TabIndex = 5;
            this.lblMessageBody.Text = "Message Body";
            // 
            // panelMessagesView
            // 
            this.panelMessagesView.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.panelMessagesView.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.panelMessagesView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelMessagesView.Controls.Add(this.listMessages);
            this.panelMessagesView.Controls.Add(this.lblMessageNumber);
            this.panelMessagesView.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelMessagesView.Location = new System.Drawing.Point(0, 0);
            this.panelMessagesView.Name = "panelMessagesView";
            this.panelMessagesView.Size = new System.Drawing.Size(281, 357);
            this.panelMessagesView.TabIndex = 5;
            // 
            // listMessages
            // 
            this.listMessages.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listMessages.ContextMenu = this.contextMenuMessages;
            this.listMessages.Location = new System.Drawing.Point(8, 24);
            this.listMessages.Name = "listMessages";
            this.listMessages.Size = new System.Drawing.Size(262, 317);
            this.listMessages.TabIndex = 8;
            this.listMessages.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.ListMessagesMessageSelected);
            // 
            // contextMenuMessages
            // 
            this.contextMenuMessages.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuViewSource,
            this.Tasking,
            this.menuItem1,
            this.menuItem2});
            // 
            // menuViewSource
            // 
            this.menuViewSource.Index = 0;
            this.menuViewSource.Text = "View source";
            this.menuViewSource.Click += new System.EventHandler(this.MenuViewSourceClick);
            // 
            // Tasking
            // 
            this.Tasking.Index = 1;
            this.Tasking.Text = "Tasking";
            this.Tasking.Click += new System.EventHandler(this.Tasking_Click);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 2;
            this.menuItem1.Text = "Delete";
            this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
            // 
            // menuItem2
            // 
            this.menuItem2.Index = 3;
            this.menuItem2.Text = "Save in Doc";
            this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
            // 
            // lblMessageNumber
            // 
            this.lblMessageNumber.Location = new System.Drawing.Point(8, 8);
            this.lblMessageNumber.Name = "lblMessageNumber";
            this.lblMessageNumber.Size = new System.Drawing.Size(136, 16);
            this.lblMessageNumber.TabIndex = 1;
            this.lblMessageNumber.Text = "Messages";
            // 
            // attachmentPanel
            // 
            this.attachmentPanel.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.attachmentPanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.attachmentPanel.Controls.Add(this.listAttachments);
            this.attachmentPanel.Controls.Add(this.lblAttachments);
            this.attachmentPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.attachmentPanel.Location = new System.Drawing.Point(695, 0);
            this.attachmentPanel.Name = "attachmentPanel";
            this.attachmentPanel.Size = new System.Drawing.Size(208, 357);
            this.attachmentPanel.TabIndex = 4;
            this.attachmentPanel.Visible = false;
            // 
            // listAttachments
            // 
            this.listAttachments.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listAttachments.Location = new System.Drawing.Point(8, 24);
            this.listAttachments.Name = "listAttachments";
            this.listAttachments.ShowLines = false;
            this.listAttachments.ShowRootLines = false;
            this.listAttachments.Size = new System.Drawing.Size(188, 317);
            this.listAttachments.TabIndex = 10;
            this.listAttachments.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.listAttachments_AfterSelect);
            // 
            // lblAttachments
            // 
            this.lblAttachments.Location = new System.Drawing.Point(12, 8);
            this.lblAttachments.Name = "lblAttachments";
            this.lblAttachments.Size = new System.Drawing.Size(136, 16);
            this.lblAttachments.TabIndex = 3;
            this.lblAttachments.Text = "Attachments";
            // 
            // saveFile
            // 
            this.saveFile.Title = "Save Attachment";
            // 
            // TestForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(903, 720);
            this.Controls.Add(this.panelMiddle);
            this.Controls.Add(this.panelProperties);
            this.Controls.Add(this.panelTop);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TestForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EmailClient";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.panelProperties.ResumeLayout(false);
            this.panelProperties.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridHeaders)).EndInit();
            this.panelMiddle.ResumeLayout(false);
            this.panelMessageBody.ResumeLayout(false);
            this.panelMessageBody.PerformLayout();
            this.panelMessagesView.ResumeLayout(false);
            this.attachmentPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        [STAThread]
        private static void Main()
        {
            Application.Run(new TestForm());
        }

        private void ReceiveMails()
        {

            connectButton.Enabled = false;
            progressBar.Value = 0;


            try
            {
                if (pop3Client.Connected)
                    pop3Client.Disconnect();
                pop3Client.Connect(popServerTextBox.Text, int.Parse(portTextBox.Text), useSslCheckBox.Checked);
                pop3Client.Authenticate(loginTextBox.Text, passwordTextBox.Text, AuthenticationMethod.UsernameAndPassword);
                int count = pop3Client.GetMessageCount();
                totalMessagesTextBox.Text = count.ToString();
                messageTextBox.Text = "";
                messages.Clear();
                listMessages.Nodes.Clear();
                listAttachments.Nodes.Clear();

                int s = 0;
                int f = 0;

                for (int i = count; i >= 1; i -= 1)
                {

                    Application.DoEvents();

                    try
                    {
                        Message message = pop3Client.GetMessage(i);

                        messages.Add(i, message);

                        TreeNode node = new TreeNodeBuilder().VisitMessage(message);

                        node.Tag = i;

                        listMessages.Nodes.Add(node);
                        s++;

                    }
                    catch (Exception e)
                    {
                        DefaultLogger.Log.LogError(
                            "Retrieving Emails: Message fetching failed: " + e.Message);
                        f++;
                    }
                    progressBar.Value = (int)(((double)(count - 1) / count) * 100);
                }

                MessageBox.Show(this, "Mail received! \n Succesfully: " + s + "\nFailed: " + f, "Messages Retrieved.");

                if (f > 0)
                {
                    MessageBox.Show("Some messages were not parsed correctly");
                }
            }
            catch (InvalidLoginException)
            {
                MessageBox.Show(this, "The server did not accept the username and password.");
            }
            catch (PopServerNotFoundException)
            {
                MessageBox.Show(this, "The server couldnt be found");
            }
            catch (PopServerLockedException)
            {
                MessageBox.Show(this, "The mailbox is locked and seems to be in use.");
            }
            catch (LoginDelayException)
            {
                MessageBox.Show(this, "You tried to log in too quickly between two logins.");
            }
            catch (Exception e)
            {
                MessageBox.Show(this, "Error ocurred retrieveing mail " + e.Message);
            }
            finally
            {
                connectButton.Enabled = true;
                progressBar.Value = 100;
            }

        }
        private void ConnectAndRetrieveButtonClick(object sender, EventArgs e)
        {
            ReceiveMails();
        }

        private void ListMessagesMessageSelected(object sender, TreeViewEventArgs e)
        {

            Message message = messages[GetMessageNumberFromSelectedNode(listMessages.SelectedNode)];

            if (listMessages.SelectedNode.Tag is MessagePart)
            {
                MessagePart selectedMessagePart = (MessagePart)listMessages.SelectedNode.Tag;
                if (selectedMessagePart.IsText)
                {

                    messageTextBox.Text = selectedMessagePart.GetBodyAsText();
                }
                else
                {
                    messageTextBox.Text = "This part of the email is not text";
                }
            }
            else
            {
                MessagePart plainTextPart = message.FindFirstPlainTextVersion();
                if (plainTextPart != null)
                {
                    messageTextBox.Text = plainTextPart.GetBodyAsText();
                }

                listAttachments.Nodes.Clear();

                List<MessagePart> attachments = message.FindAllAttachments();

                foreach (MessagePart attachment in attachments)
                {
                    TreeNode addedNode = listAttachments.Nodes.Add((attachment.FileName));
                    addedNode.Tag = attachment;
                }

                DataSet dataSet = new DataSet();
                DataTable table = dataSet.Tables.Add("Headers");
                table.Columns.Add("Header");
                table.Columns.Add("Value");

                DataRowCollection rows = table.Rows;

                foreach (RfcMailAddress cc in message.Headers.Cc) rows.Add(new object[] { "Cc", cc });
                foreach (RfcMailAddress bcc in message.Headers.Bcc) rows.Add(new object[] { "Bcc", bcc });
                foreach (RfcMailAddress to in message.Headers.To) rows.Add(new object[] { "To", to });
                rows.Add(new object[] { "From", message.Headers.From });
                rows.Add(new object[] { "Reply-To", message.Headers.ReplyTo });
                rows.Add(new object[] { "Date", message.Headers.Date });
                rows.Add(new object[] { "Subject", message.Headers.Subject });


                gridHeaders.DataMember = table.TableName;
                gridHeaders.DataSource = dataSet;
            }
        }

        private void Tasking_Click(object sender, EventArgs e)
        {
            SqlConnection connection;
             
            int messageNumber = GetMessageNumberFromSelectedNode(listMessages.SelectedNode);
            Message message = messages[messageNumber];
            MessagePart plainTextPart = message.FindFirstPlainTextVersion();          

            String connetionString = GetConnectionString();           
            connection = new SqlConnection(connetionString);

            int maxId = 0;

            //reader
            try
            {
                //open connection
                connection.Open();


                    //command1
                    String sqlGetMaxId = "select MAX(Id) from Emails";
                    SqlCommand command1 = new SqlCommand(sqlGetMaxId, connection);
                    using (SqlDataReader rdr = command1.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (!rdr.IsDBNull(0))
                                maxId = rdr.GetInt32(0);
                        }
                    }
                    command1.Dispose();

                    //command2
                    String sqlInsertEmail = "Insert into Emails(EmailSubject, Body, EmailFrom, EmailReplyTo, EmailDate) Values('" + message.Headers.Subject.ToString() + "','" + plainTextPart.GetBodyAsText() + "','" + message.Headers.From.ToString() + "','" + message.Headers.ReplyTo.ToString() + "','" + message.Headers.Date.ToString() + "')";
                    SqlCommand command2 = new SqlCommand(sqlInsertEmail, connection);
                    command2.ExecuteNonQuery();
                    command2.Dispose();
                


                //command3
                List<MessagePart> attachments = message.FindAllAttachments();
                foreach (MessagePart attachment in attachments)
                {                 
                    String sqlInsertAttachment = "Insert into Attachments(EmailId, AttachmentName) Values('" + (maxId + 1) + "','" + attachment.FileName + "')";
                    SqlCommand command3 = new SqlCommand(sqlInsertAttachment, connection);
                    command3.ExecuteNonQuery();                  
                    command3.Dispose();               
                }

                //close connection
                connection.Close();

                if (listMessages.SelectedNode != null)
                    listMessages.SelectedNode.BackColor = Color.Yellow;              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }        

        }

        private static int GetMessageNumberFromSelectedNode(TreeNode node)
        {
            if (node == null)
                throw new ArgumentNullException("node");

            if (node.Tag is int)
            {
                return (int)node.Tag;
            }

            return GetMessageNumberFromSelectedNode(node.Parent);
        }
        private void MenuViewSourceClick(object sender, EventArgs e)
        {

            if (listMessages.SelectedNode != null)
            {
                int messageNumber = GetMessageNumberFromSelectedNode(listMessages.SelectedNode);
                Message m = messages[messageNumber];

                ShowSourceForm sourceForm = new ShowSourceForm(Encoding.ASCII.GetString(m.RawMessage));
                sourceForm.ShowDialog();
            }
        }

        private void loginTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void menuDeleteMessage_Click(object sender, EventArgs e)
        {

        }

        private void listAttachments_AfterSelect(object sender, TreeViewEventArgs e)
        {
            MessagePart attachment = (MessagePart)listAttachments.SelectedNode.Tag;

            if (attachment != null)
            {
                saveFile.FileName = attachment.FileName;
                DialogResult result = saveFile.ShowDialog();
                if (result != DialogResult.OK)
                    return;
                FileInfo file = new FileInfo(saveFile.FileName);

                if (file.Exists)
                {
                    file.Delete();
                }
                try
                {
                    attachment.Save(file);

                    MessageBox.Show(this, "Attachment saved successfully");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Attachment saving failed. Exception message: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show(this, "Attachment object was null!");
            }
        }

        private String GetConnectionString()
        {
            string connetionString = null;
 
            if (loginTypeCb.Text == "Windows Authentication")
            {
                //Server = localhost
                //Database = EmailClient
                return connetionString = "Server= " + serverNameTxt.Text + "; Database= " + databaseNameTxt.Text + "; Integrated Security=SSPI;";
            }
            else
            {
                //Data Source = 
                //Initial Catalog = EmailClient
                //User ID = 
                //Password = 
               return "Data Source=" + serverNameTxt.Text + ";Initial Catalog=" + databaseNameTxt.Text + ";User ID=" + usernameTxt.Text + "; Password=" + passwordTxt.Text + "";
            }         
           
        }       

        private void loginTypeCb_SelectedIndexChanged(object sender,
        System.EventArgs e)
        {

            ComboBox comboBox = (ComboBox)sender;
            string selectedValue = (string)loginTypeCb.SelectedItem;
            if (selectedValue == "Windows Authentication")
            {
                usernameTxt.Visible = false;
                passwordTxt.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
            }
            else
            {
                usernameTxt.Visible = true;
                passwordTxt.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
            }       
            
        }

        private void userAccountsBtn_Click(object sender, EventArgs e)
        {
            UserAccountsForm userAccountForms = new UserAccountsForm( GetConnectionString());
            userAccountForms.ShowDialog();
        }

        private void menuItem1_Click(object sender, EventArgs e)
        {
            {
                if (listMessages.SelectedNode != null)
                {
                    DialogResult drRet = MessageBox.Show(this, "Are you sure to delete the email?", "Delete email", MessageBoxButtons.YesNo);
                    if (drRet == DialogResult.Yes)
                    {
                        int messageNumber = GetMessageNumberFromSelectedNode(listMessages.SelectedNode);
                        pop3Client.DeleteMessage(messageNumber);

                        listMessages.Nodes[messageNumber].Remove();

                        drRet = MessageBox.Show(this, "Do you want to receive email again (this will commit your changes)?", "Receive email", MessageBoxButtons.YesNo);
                        if (drRet == DialogResult.Yes)
                            ReceiveMails();
                    }
                }
            }
        }

        private void menuItem2_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = new StreamWriter("C://Users/Laptop/Desktop/out.txt"))
            {

                Message message = messages[GetMessageNumberFromSelectedNode(listMessages.SelectedNode)];

                foreach (RfcMailAddress cc in message.Headers.Cc) sw.WriteLine("Cc: " + cc.ToString());
                foreach (RfcMailAddress bcc in message.Headers.Bcc) sw.WriteLine("Bcc: " + bcc.ToString());
                foreach (RfcMailAddress to in message.Headers.To) sw.WriteLine("To: " + to.ToString());
                sw.WriteLine("From: " + message.Headers.From.ToString());
                sw.WriteLine("Reply-To: " + message.Headers.ReplyTo.ToString());
                sw.WriteLine("Date: " + message.Headers.Date.ToString());
                sw.WriteLine("Subject: " + message.Headers.Subject.ToString());
                //sw.WriteLine("Sender" + message.Headers.Sender.ToString());
                //sw.WriteLine( "Content-Type" + message.Headers.ContentType.ToString());
                //sw.WriteLine( "Content-Disposition"+ message.Headers.ContentDisposition.ToString());
                sw.WriteLine("Message-Id" + message.Headers.MessageId.ToString());

                sw.WriteLine("--------------------------------------------------------------------------------------------------------------------------------------------------------------");
                sw.WriteLine("-------------------------------------------------------------------BODY---------------------------------------------------------------------------------------");
                sw.WriteLine("--------------------------------------------------------------------------------------------------------------------------------------------------------------");


                int messageNumber = GetMessageNumberFromSelectedNode(listMessages.SelectedNode);
                MessagePart plainTextPart = message.FindFirstPlainTextVersion();
                sw.WriteLine(plainTextPart.GetBodyAsText());


            }
        }
    }
}