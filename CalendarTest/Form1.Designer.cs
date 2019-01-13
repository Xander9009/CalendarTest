namespace CalendarTest
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.EventsList = new System.Windows.Forms.ListView();
            this.btnNewEvent = new System.Windows.Forms.Button();
            this.Description = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.When = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.UserSelectionPanel = new System.Windows.Forms.Panel();
            this.UserListView = new System.Windows.Forms.ListView();
            this.btnLoadUser = new System.Windows.Forms.Button();
            this.btnNewUser = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnChangeUser = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblCurrentUser = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblEmail = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblURL = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.trayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.notificationIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.UserSelectionPanel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // EventsList
            // 
            this.EventsList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.EventsList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Description,
            this.When});
            this.EventsList.Location = new System.Drawing.Point(3, 215);
            this.EventsList.Name = "EventsList";
            this.EventsList.Size = new System.Drawing.Size(794, 179);
            this.EventsList.TabIndex = 0;
            this.EventsList.UseCompatibleStateImageBehavior = false;
            this.EventsList.View = System.Windows.Forms.View.Details;
            // 
            // btnNewEvent
            // 
            this.btnNewEvent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNewEvent.Location = new System.Drawing.Point(722, 400);
            this.btnNewEvent.Name = "btnNewEvent";
            this.btnNewEvent.Size = new System.Drawing.Size(75, 23);
            this.btnNewEvent.TabIndex = 1;
            this.btnNewEvent.Text = "New Event";
            this.btnNewEvent.UseVisualStyleBackColor = true;
            this.btnNewEvent.Click += new System.EventHandler(this.btnNewEvent_Click);
            // 
            // Description
            // 
            this.Description.Text = "Description";
            this.Description.Width = 248;
            // 
            // When
            // 
            this.When.Text = "When";
            this.When.Width = 167;
            // 
            // UserSelectionPanel
            // 
            this.UserSelectionPanel.Controls.Add(this.btnNewUser);
            this.UserSelectionPanel.Controls.Add(this.btnLoadUser);
            this.UserSelectionPanel.Controls.Add(this.UserListView);
            this.UserSelectionPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.UserSelectionPanel.Location = new System.Drawing.Point(0, 24);
            this.UserSelectionPanel.Name = "UserSelectionPanel";
            this.UserSelectionPanel.Size = new System.Drawing.Size(800, 426);
            this.UserSelectionPanel.TabIndex = 2;
            // 
            // UserListView
            // 
            this.UserListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserListView.Location = new System.Drawing.Point(13, 13);
            this.UserListView.Name = "UserListView";
            this.UserListView.Size = new System.Drawing.Size(693, 401);
            this.UserListView.TabIndex = 0;
            this.UserListView.UseCompatibleStateImageBehavior = false;
            this.UserListView.View = System.Windows.Forms.View.List;
            // 
            // btnLoadUser
            // 
            this.btnLoadUser.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoadUser.Enabled = false;
            this.btnLoadUser.Location = new System.Drawing.Point(713, 13);
            this.btnLoadUser.Name = "btnLoadUser";
            this.btnLoadUser.Size = new System.Drawing.Size(75, 23);
            this.btnLoadUser.TabIndex = 1;
            this.btnLoadUser.Text = "Load User";
            this.btnLoadUser.UseVisualStyleBackColor = true;
            this.btnLoadUser.Click += new System.EventHandler(this.btnLoadUser_Click);
            // 
            // btnNewUser
            // 
            this.btnNewUser.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNewUser.Location = new System.Drawing.Point(713, 42);
            this.btnNewUser.Name = "btnNewUser";
            this.btnNewUser.Size = new System.Drawing.Size(75, 23);
            this.btnNewUser.TabIndex = 2;
            this.btnNewUser.Text = "New User";
            this.btnNewUser.UseVisualStyleBackColor = true;
            this.btnNewUser.Click += new System.EventHandler(this.btnNewUser_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.lblURL);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.lblEmail);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.lblCurrentUser);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnChangeUser);
            this.panel1.Controls.Add(this.EventsList);
            this.panel1.Controls.Add(this.btnNewEvent);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 24);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(800, 426);
            this.panel1.TabIndex = 3;
            // 
            // btnChangeUser
            // 
            this.btnChangeUser.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChangeUser.Location = new System.Drawing.Point(639, 400);
            this.btnChangeUser.Name = "btnChangeUser";
            this.btnChangeUser.Size = new System.Drawing.Size(77, 23);
            this.btnChangeUser.TabIndex = 2;
            this.btnChangeUser.Text = "Change User";
            this.btnChangeUser.UseVisualStyleBackColor = true;
            this.btnChangeUser.Click += new System.EventHandler(this.btnChangeUser_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 199);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(184, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Upcoming Events in Google Calendar";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Current User";
            // 
            // lblCurrentUser
            // 
            this.lblCurrentUser.AutoSize = true;
            this.lblCurrentUser.Location = new System.Drawing.Point(78, 12);
            this.lblCurrentUser.Name = "lblCurrentUser";
            this.lblCurrentUser.Size = new System.Drawing.Size(0, 13);
            this.lblCurrentUser.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 34);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Email";
            // 
            // lblEmail
            // 
            this.lblEmail.AutoSize = true;
            this.lblEmail.Location = new System.Drawing.Point(78, 34);
            this.lblEmail.Name = "lblEmail";
            this.lblEmail.Size = new System.Drawing.Size(0, 13);
            this.lblEmail.TabIndex = 7;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 56);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 13);
            this.label6.TabIndex = 8;
            this.label6.Text = "URL";
            // 
            // lblURL
            // 
            this.lblURL.AutoSize = true;
            this.lblURL.Location = new System.Drawing.Point(78, 56);
            this.lblURL.Name = "lblURL";
            this.lblURL.Size = new System.Drawing.Size(0, 13);
            this.lblURL.TabIndex = 9;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.trayToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // trayToolStripMenuItem
            // 
            this.trayToolStripMenuItem.Name = "trayToolStripMenuItem";
            this.trayToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.trayToolStripMenuItem.Text = "Tray";
            this.trayToolStripMenuItem.Click += new System.EventHandler(this.trayToolStripMenuItem_Click);
            // 
            // notificationIcon
            // 
            this.notificationIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notificationIcon.Icon")));
            this.notificationIcon.Text = "notifyIcon1";
            this.notificationIcon.Visible = true;
            this.notificationIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notificationIcon_MouseDoubleClick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.UserSelectionPanel);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.UserSelectionPanel.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView EventsList;
        private System.Windows.Forms.Button btnNewEvent;
        private System.Windows.Forms.ColumnHeader Description;
        private System.Windows.Forms.ColumnHeader When;
        private System.Windows.Forms.Panel UserSelectionPanel;
        private System.Windows.Forms.Button btnNewUser;
        private System.Windows.Forms.Button btnLoadUser;
        private System.Windows.Forms.ListView UserListView;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblURL;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblEmail;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblCurrentUser;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnChangeUser;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem trayToolStripMenuItem;
        private System.Windows.Forms.NotifyIcon notificationIcon;
    }
}

