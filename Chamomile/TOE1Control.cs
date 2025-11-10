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
    public partial class TOE1Control : UserControl
    {
        public TOE1Control()
        {
            InitializeComponent();
        }

        private void calculation_Click(object sender, EventArgs e)
        {
            
        }

        private void label19_Click(object sender, EventArgs e)
        {
            if (label19.Text == "-")
                label19.Text = "+";
            else
                label19.Text = "-";
        }

        private void label16_Click(object sender, EventArgs e)
        {
            if (label16.Text == "-")
                label16.Text = "+";
            else
                label16.Text = "-";
        }

        private void label20_Click(object sender, EventArgs e)
        {
            if (label20.Text == "-")
                label20.Text = "+";
            else
                label20.Text = "-";
        }
    }
}
