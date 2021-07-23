using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CabbageChaosBot;

namespace TwitchBotForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label2.Text = "INACTIVE";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Program.bot.InitiateEndingSequence();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            bool currentStatus = Program.bot.ToggleKillSwitch();

            if (currentStatus)
            {
                label2.Text = "ACTIVE";
            }
            else
            {
                label2.Text = "INACTIVE";
            }
        }
    }
}
