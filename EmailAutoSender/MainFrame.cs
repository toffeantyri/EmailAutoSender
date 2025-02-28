using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EmailAutoSender
{
    public partial class MainFrame : Form
    {



        public MainFrame()
        {
            InitializeComponent();
        }

        private void MainFrame_Load(object sender, EventArgs e)
        {

        }

        private void LoadSettings() 
        {
            txtEmail.Text = Properties.Settings.Default.email;
            txtPassword.Text = Properties.Settings.Default.password;
            txtSmtpServer.Text = Properties.Settings.Default.smtpServer;
            txtPort.Text = Properties.Settings.Default.smtpPort.ToString();
            txtExcelPath.Text = Properties.Settings.Default.filePath;
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
