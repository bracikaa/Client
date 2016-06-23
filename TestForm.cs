using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;

namespace cs499
{
	public class TestForm : Form
	{

		private Label labelAttachments;
		private Label labelMessageBody;
		private Label labelMessageNumber;
		private Label labelPassword;
		private Label labelUsername;
        private readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
        private readonly Pop3Client pop3Client;
        private Button connectButton;
        private Panel attachmentPanel;
        private ContextMenu contextMenuMessages;
        private DataGrid gridHeaders;
        private Panel panelMiddle;
        private Panel panelMessageBody;
        private Panel panelMessagesView;
        private TextBox messageTextBox;
        private TextBox popServerTextBox;
        private TextBox passwordTextBox;
        private TextBox portTextBox;
        private MenuItem Tasking;
        private TextBox totalMessagesTextBox;
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
            this.useSslCheckBox = new System.Windows.Forms.CheckBox();
            this.labelPassword = new System.Windows.Forms.Label();
            this.passwordTextBox = new System.Windows.Forms.TextBox();
            this.labelUsername = new System.Windows.Forms.Label();
            this.loginTextBox = new System.Windows.Forms.TextBox();
            this.connectButton = new System.Windows.Forms.Button();
            this.labelServerPort = new System.Windows.Forms.Label();
            this.portTextBox = new System.Windows.Forms.TextBox();
            this.labelServerAddress = new System.Windows.Forms.Label();
            this.popServerTextBox = new System.Windows.Forms.TextBox();
            this.panelProperties = new System.Windows.Forms.Panel();
            this.gridHeaders = new System.Windows.Forms.DataGrid();
            this.panelMiddle = new System.Windows.Forms.Panel();
            this.panelMessageBody = new System.Windows.Forms.Panel();
            this.messageTextBox = new System.Windows.Forms.TextBox();
            this.labelMessageBody = new System.Windows.Forms.Label();
            this.panelMessagesView = new System.Windows.Forms.Panel();
            this.listMessages = new System.Windows.Forms.TreeView();
            this.contextMenuMessages = new System.Windows.Forms.ContextMenu();
            this.menuViewSource = new System.Windows.Forms.MenuItem();
            this.Tasking = new System.Windows.Forms.MenuItem();
            this.labelMessageNumber = new System.Windows.Forms.Label();
            this.attachmentPanel = new System.Windows.Forms.Panel();
            this.listAttachments = new System.Windows.Forms.TreeView();
            this.labelAttachments = new System.Windows.Forms.Label();
            this.saveFile = new System.Windows.Forms.SaveFileDialog();
            this.totalMessagesTextBox = new System.Windows.Forms.TextBox();
            this.labelTotalMessages = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
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
            this.panelTop.Controls.Add(this.totalMessagesTextBox);
            this.panelTop.Controls.Add(this.labelTotalMessages);
            this.panelTop.Controls.Add(this.useSslCheckBox);
            this.panelTop.Controls.Add(this.labelPassword);
            this.panelTop.Controls.Add(this.passwordTextBox);
            this.panelTop.Controls.Add(this.labelUsername);
            this.panelTop.Controls.Add(this.loginTextBox);
            this.panelTop.Controls.Add(this.connectButton);
            this.panelTop.Controls.Add(this.labelServerPort);
            this.panelTop.Controls.Add(this.portTextBox);
            this.panelTop.Controls.Add(this.labelServerAddress);
            this.panelTop.Controls.Add(this.popServerTextBox);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(865, 64);
            this.panelTop.TabIndex = 0;
            // 
            // useSslCheckBox
            // 
            this.useSslCheckBox.AutoSize = true;
            this.useSslCheckBox.Checked = true;
            this.useSslCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.useSslCheckBox.Location = new System.Drawing.Point(316, 30);
            this.useSslCheckBox.Name = "useSslCheckBox";
            this.useSslCheckBox.Size = new System.Drawing.Size(68, 17);
            this.useSslCheckBox.TabIndex = 4;
            this.useSslCheckBox.Text = "Use SSL";
            this.useSslCheckBox.UseVisualStyleBackColor = true;
            // 
            // labelPassword
            // 
            this.labelPassword.Location = new System.Drawing.Point(562, 32);
            this.labelPassword.Name = "labelPassword";
            this.labelPassword.Size = new System.Drawing.Size(64, 23);
            this.labelPassword.TabIndex = 8;
            this.labelPassword.Text = "Password";
            // 
            // passwordTextBox
            // 
            this.passwordTextBox.Location = new System.Drawing.Point(632, 32);
            this.passwordTextBox.Name = "passwordTextBox";
            this.passwordTextBox.PasswordChar = '*';
            this.passwordTextBox.Size = new System.Drawing.Size(128, 20);
            this.passwordTextBox.TabIndex = 2;
            // 
            // labelUsername
            // 
            this.labelUsername.Location = new System.Drawing.Point(562, 4);
            this.labelUsername.Name = "labelUsername";
            this.labelUsername.Size = new System.Drawing.Size(64, 23);
            this.labelUsername.TabIndex = 6;
            this.labelUsername.Text = "Email";
            // 
            // loginTextBox
            // 
            this.loginTextBox.Location = new System.Drawing.Point(632, 5);
            this.loginTextBox.Name = "loginTextBox";
            this.loginTextBox.Size = new System.Drawing.Size(128, 20);
            this.loginTextBox.TabIndex = 1;
            this.loginTextBox.Text = "nedimp_97@hotmail.com";
            this.loginTextBox.TextChanged += new System.EventHandler(this.loginTextBox_TextChanged);
            // 
            // connectAndRetrieveButton
            // 
            this.connectButton.Location = new System.Drawing.Point(771, 5);
            this.connectButton.Name = "connectAndRetrieveButton";
            this.connectButton.Size = new System.Drawing.Size(86, 54);
            this.connectButton.TabIndex = 5;
            this.connectButton.Text = "Connect";
            this.connectButton.Click += new System.EventHandler(this.ConnectAndRetrieveButtonClick);
            // 
            // labelServerPort
            // 
            this.labelServerPort.Location = new System.Drawing.Point(394, 34);
            this.labelServerPort.Name = "labelServerPort";
            this.labelServerPort.Size = new System.Drawing.Size(31, 23);
            this.labelServerPort.TabIndex = 3;
            this.labelServerPort.Text = "Port";
            // 
            // portTextBox
            // 
            this.portTextBox.Location = new System.Drawing.Point(428, 35);
            this.portTextBox.Name = "portTextBox";
            this.portTextBox.Size = new System.Drawing.Size(128, 20);
            this.portTextBox.TabIndex = 3;
            this.portTextBox.Text = "995";
            // 
            // labelServerAddress
            // 
            this.labelServerAddress.Location = new System.Drawing.Point(313, 4);
            this.labelServerAddress.Name = "labelServerAddress";
            this.labelServerAddress.Size = new System.Drawing.Size(112, 23);
            this.labelServerAddress.TabIndex = 1;
            this.labelServerAddress.Text = "POP Server Address";
            // 
            // popServerTextBox
            // 
            this.popServerTextBox.Location = new System.Drawing.Point(428, 5);
            this.popServerTextBox.Name = "popServerTextBox";
            this.popServerTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.popServerTextBox.Size = new System.Drawing.Size(128, 20);
            this.popServerTextBox.TabIndex = 0;
            this.popServerTextBox.Text = "pop-mail.outlook.com";
            // 
            // panelProperties
            // 
            this.panelProperties.Controls.Add(this.gridHeaders);
            this.panelProperties.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelProperties.Location = new System.Drawing.Point(0, 260);
            this.panelProperties.Name = "panelProperties";
            this.panelProperties.Size = new System.Drawing.Size(865, 184);
            this.panelProperties.TabIndex = 1;
            // 
            // gridHeaders
            // 
            this.gridHeaders.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridHeaders.DataMember = "";
            this.gridHeaders.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.gridHeaders.Location = new System.Drawing.Point(0, 0);
            this.gridHeaders.Name = "gridHeaders";
            this.gridHeaders.PreferredColumnWidth = 400;
            this.gridHeaders.ReadOnly = true;
            this.gridHeaders.Size = new System.Drawing.Size(865, 188);
            this.gridHeaders.TabIndex = 3;
            // 
            // panelMiddle
            // 
            this.panelMiddle.Controls.Add(this.panelMessageBody);
            this.panelMiddle.Controls.Add(this.panelMessagesView);
            this.panelMiddle.Controls.Add(this.attachmentPanel);
            this.panelMiddle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMiddle.Location = new System.Drawing.Point(0, 64);
            this.panelMiddle.Name = "panelMiddle";
            this.panelMiddle.Size = new System.Drawing.Size(865, 196);
            this.panelMiddle.TabIndex = 2;
            // 
            // panelMessageBody
            // 
            this.panelMessageBody.Controls.Add(this.progressBar);
            this.panelMessageBody.Controls.Add(this.messageTextBox);
            this.panelMessageBody.Controls.Add(this.labelMessageBody);
            this.panelMessageBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMessageBody.Location = new System.Drawing.Point(281, 0);
            this.panelMessageBody.Name = "panelMessageBody";
            this.panelMessageBody.Size = new System.Drawing.Size(376, 196);
            this.panelMessageBody.TabIndex = 6;
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
            this.messageTextBox.Size = new System.Drawing.Size(360, 143);
            this.messageTextBox.TabIndex = 9;
            // 
            // labelMessageBody
            // 
            this.labelMessageBody.Location = new System.Drawing.Point(8, 8);
            this.labelMessageBody.Name = "labelMessageBody";
            this.labelMessageBody.Size = new System.Drawing.Size(136, 16);
            this.labelMessageBody.TabIndex = 5;
            this.labelMessageBody.Text = "Message Body";
            // 
            // panelMessagesView
            // 
            this.panelMessagesView.Controls.Add(this.listMessages);
            this.panelMessagesView.Controls.Add(this.labelMessageNumber);
            this.panelMessagesView.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelMessagesView.Location = new System.Drawing.Point(0, 0);
            this.panelMessagesView.Name = "panelMessagesView";
            this.panelMessagesView.Size = new System.Drawing.Size(281, 196);
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
            this.listMessages.Size = new System.Drawing.Size(266, 160);
            this.listMessages.TabIndex = 8;
            this.listMessages.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.ListMessagesMessageSelected);
            // 
            // contextMenuMessages
            // 
            this.contextMenuMessages.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuViewSource,
            this.Tasking});
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
            // labelMessageNumber
            // 
            this.labelMessageNumber.Location = new System.Drawing.Point(8, 8);
            this.labelMessageNumber.Name = "labelMessageNumber";
            this.labelMessageNumber.Size = new System.Drawing.Size(136, 16);
            this.labelMessageNumber.TabIndex = 1;
            this.labelMessageNumber.Text = "Messages";
            // 
            // attachmentPanel
            // 
            this.attachmentPanel.Controls.Add(this.listAttachments);
            this.attachmentPanel.Controls.Add(this.labelAttachments);
            this.attachmentPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.attachmentPanel.Location = new System.Drawing.Point(657, 0);
            this.attachmentPanel.Name = "attachmentPanel";
            this.attachmentPanel.Size = new System.Drawing.Size(208, 196);
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
            this.listAttachments.Size = new System.Drawing.Size(192, 160);
            this.listAttachments.TabIndex = 10;
            this.listAttachments.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.listAttachments_AfterSelect);
            // 
            // labelAttachments
            // 
            this.labelAttachments.Location = new System.Drawing.Point(12, 8);
            this.labelAttachments.Name = "labelAttachments";
            this.labelAttachments.Size = new System.Drawing.Size(136, 16);
            this.labelAttachments.TabIndex = 3;
            this.labelAttachments.Text = "Attachments";
            // 
            // saveFile
            // 
            this.saveFile.Title = "Save Attachment";
            // 
            // totalMessagesTextBox
            // 
            this.totalMessagesTextBox.Location = new System.Drawing.Point(8, 37);
            this.totalMessagesTextBox.Name = "totalMessagesTextBox";
            this.totalMessagesTextBox.Size = new System.Drawing.Size(44, 20);
            this.totalMessagesTextBox.TabIndex = 10;
            // 
            // labelTotalMessages
            // 
            this.labelTotalMessages.Location = new System.Drawing.Point(5, 11);
            this.labelTotalMessages.Name = "labelTotalMessages";
            this.labelTotalMessages.Size = new System.Drawing.Size(100, 23);
            this.labelTotalMessages.TabIndex = 11;
            this.labelTotalMessages.Text = "Total Messages";
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(7, 172);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(360, 18);
            this.progressBar.TabIndex = 11;
            // 
            // TestForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(865, 444);
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
                pop3Client.Authenticate(loginTextBox.Text, passwordTextBox.Text);
                int count = pop3Client.GetMessageCount();
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

                if(f > 0)
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
            catch(LoginDelayException)
            {
                MessageBox.Show(this, "You tried to log in too quickly between two logins.");
            }
            catch(Exception e)
            {
                MessageBox.Show(this, "Error ocurred retrieveing mail " + e.Message);
            }
            finally
            {

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

		private static int GetMessageNumberFromSelectedNode(TreeNode node)
		{
			if (node == null)
				throw new ArgumentNullException("node");

			if(node.Tag is int)
			{
				return (int) node.Tag;
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

        private void Tasking_Click(object sender, EventArgs e)
        {
            //int messageNumber = GetMessageNumberFromSelectedNode(listMessages.SelectedNode);
            //Message message = messages[messageNumber];
            //string text = 

            //if (listMessages.SelectedNode != null)
            //{

            //    FileInfo file = new FileInfo("tasking.docx");


            //    message.Save(file);

            //}

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
}
    }