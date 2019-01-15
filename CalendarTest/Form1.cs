using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalendarTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            notificationIcon.Visible = false;
        }

        private void btnNewUser_Click(object sender, EventArgs e)
        {
            Program.NewUser();
        }

        private bool maxedWhenTrayed = false;
        private void trayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            maxedWhenTrayed = this.WindowState == FormWindowState.Maximized;
            this.WindowState = FormWindowState.Minimized;
            notificationIcon.Visible = true;
            this.ShowInTaskbar = false;
        }

        private void notificationIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (maxedWhenTrayed)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
            notificationIcon.Visible = false;
            this.ShowInTaskbar = true;
        }

        private void UserListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            Program.LoadUser();
        }

        private void lblURL_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://" + (sender as Control).Text);
        }
    }
}
