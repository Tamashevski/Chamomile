using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Chamomile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void раздел1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            contentPanel.Controls.Clear();
            var toeControl = new TOE1Control();
            toeControl.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(toeControl);
            


        }

        private void раздел2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            contentPanel.Controls.Clear();
            var toeControl = new TOE2Control();
            toeControl.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(toeControl);
        }

        private void раздел3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            contentPanel.Controls.Clear();
            var toeControl = new TOE3Control();
            toeControl.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(toeControl);
        }

        private void диагностикаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            contentPanel.Controls.Clear();
            var DiagnosticsControl = new DiagnosticsControl();
            DiagnosticsControl.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(DiagnosticsControl);
        }

        private void экономикаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            contentPanel.Controls.Clear();
            var EconomyControl = new EconomyControl();
            EconomyControl.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(EconomyControl);
        }
    }
}
