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

        private void btnNewEvent_Click(object sender, EventArgs e)
        {
            Program.NewEvent();
        }

        private void btnLoadUser_Click(object sender, EventArgs e)
        {
            Program.LoadUser();
        }

        private void btnNewUser_Click(object sender, EventArgs e)
        {
            Program.NewUser();
        }

        private void btnChangeUser_Click(object sender, EventArgs e)
        {
            Program.ChangeUser();
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
    }
}
