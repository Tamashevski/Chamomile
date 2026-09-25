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
    public partial class EconomyCourseWork : UserControl
    {
        public EconomyCourseWork()
        {
            InitializeComponent();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox1.Checked = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox2.Checked = false;
            }
        }

        private void textBox5_Enter(object sender, EventArgs e)
        {
            textBox5.Text = "";
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void textBox16_Enter(object sender, EventArgs e)
        {
            textBox16.Text = "";
        }

        private void textBox15_Enter(object sender, EventArgs e)
        {
            textBox15.Text = "";
        }

        private void textBox6_Enter(object sender, EventArgs e)
        {
            textBox6.Text = "";
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            textBox4.Text = "";
        }

        private void textBox18_Enter(object sender, EventArgs e)
        {
            textBox18.Text = "";
        }

        private void textBox17_Enter(object sender, EventArgs e)
        {
            textBox17.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (checkBox1.Checked == true)
            //{
            //    textBox104.Text = textBox8.Text;
            //    textBox103.Text = textBox7.Text;
            //    textBox102.Text = textBox10.Text;
            //    textBox101.Text = textBox12.Text;
            //    textBox100.Text = textBox11.Text;
            //    textBox99.Text = textBox9.Text;
            //    textBox98.Text = textBox22.Text;
            //    textBox97.Text = textBox21.Text;
            //    textBox96.Text = textBox20.Text;
            //    textBox95.Text = textBox19.Text;
            //    textBox94.Text = textBox14.Text;
            //    textBox93.Text = textBox13.Text;
            //    textBox92.Text = textBox28.Text;
            //    textBox91.Text = textBox27.Text;
            //    textBox90.Text = textBox26.Text;
            //    textBox89.Text = textBox25.Text;
            //}
            //else
            //{
            //    textBox104.Text = textBox42.Text;
            //    textBox103.Text = textBox41.Text;
            //    textBox102.Text = textBox40.Text;
            //    textBox101.Text = textBox39.Text;
            //    textBox100.Text = textBox38.Text;
            //    textBox99.Text = textBox37.Text;
            //    textBox98.Text = textBox36.Text;
            //    textBox97.Text = textBox35.Text;
            //    textBox96.Text = textBox34.Text;
            //    textBox95.Text = textBox33.Text;
            //    textBox94.Text = textBox32.Text;
            //    textBox93.Text = textBox31.Text;
            //    textBox92.Text = textBox30.Text;
            //    textBox91.Text = textBox29.Text;
            //    textBox90.Text = textBox24.Text;
            //    textBox89.Text = textBox23.Text;
            //}

            textBox58.Text = (Math.Round(double.Parse(textBox5.Text) * double.Parse(textBox104.Text), 2)).ToString();

            textBox57.Text = (Math.Round(double.Parse(textBox103.Text) * double.Parse(textBox1.Text), 2)).ToString();

            textBox56.Text = (Math.Round(double.Parse(textBox102.Text) * double.Parse(textBox3.Text), 2)).ToString();

            textBox55.Text = (Math.Round(double.Parse(textBox101.Text) * double.Parse(textBox2.Text), 2)).ToString();

            textBox54.Text = (Math.Round(double.Parse(textBox100.Text) * double.Parse(textBox15.Text), 2)).ToString();

            textBox53.Text = (Math.Round(double.Parse(textBox58.Text) * double.Parse(textBox16.Text) / 1000, 2)).ToString();

            textBox52.Text = (Math.Round(double.Parse(textBox57.Text) + double.Parse(textBox56.Text) + double.Parse(textBox55.Text) + double.Parse(textBox54.Text) + double.Parse(textBox53.Text), 2)).ToString();

            textBox51.Text = (Math.Round(double.Parse(textBox52.Text) * double.Parse(textBox99.Text) / 100, 2)).ToString();

            //2 Раздел

            textBox50.Text = (Math.Round(double.Parse(textBox98.Text) * double.Parse(textBox57.Text), 2)).ToString();

            textBox49.Text = (Math.Round(double.Parse(textBox97.Text) * double.Parse(textBox56.Text), 2)).ToString();

            textBox48.Text = (Math.Round(double.Parse(textBox96.Text) * (double.Parse(textBox55.Text)+ double.Parse(textBox54.Text)), 2)).ToString();

            textBox47.Text = (Math.Round(double.Parse(textBox95.Text) * double.Parse(textBox53.Text), 2)).ToString();

            textBox46.Text = (Math.Round(double.Parse(textBox47.Text) + double.Parse(textBox48.Text) + double.Parse(textBox49.Text) + double.Parse(textBox50.Text), 2)).ToString();

            textBox45.Text = (Math.Round(double.Parse(textBox46.Text) * 0.4, 2)).ToString();

            textBox44.Text = (Math.Round(double.Parse(textBox46.Text) + double.Parse(textBox45.Text), 2)).ToString();

            textBox88.Text = (Math.Round(double.Parse(textBox44.Text) * 0.2, 2)).ToString();

            textBox43.Text = (Math.Round(double.Parse(textBox88.Text) + double.Parse(textBox44.Text), 2)).ToString();

            textBox62.Text = (Math.Round(double.Parse(textBox43.Text) * double.Parse(textBox94.Text), 2)).ToString();

            textBox61.Text = (Math.Round(double.Parse(textBox95.Text) * (double.Parse(textBox51.Text)), 2)).ToString();

            textBox60.Text = (Math.Round(double.Parse(textBox61.Text) * 0.4, 2)).ToString();

            textBox59.Text = (Math.Round(double.Parse(textBox60.Text) + double.Parse(textBox61.Text), 2)).ToString();

            textBox63.Text = (Math.Round(double.Parse(textBox59.Text) * 0.2, 2)).ToString();

            textBox65.Text = (Math.Round(double.Parse(textBox63.Text) + double.Parse(textBox59.Text), 2)).ToString();

            textBox64.Text = (Math.Round(double.Parse(textBox65.Text) * double.Parse(textBox94.Text), 2)).ToString();

            textBox67.Text = (Math.Round(double.Parse(textBox93.Text) * double.Parse(textBox103.Text), 2)).ToString();

            textBox66.Text = (Math.Round(double.Parse(textBox92.Text) * double.Parse(textBox102.Text), 2)).ToString();

            textBox68.Text = (Math.Round(double.Parse(textBox91.Text) * double.Parse(textBox101.Text), 2)).ToString();

            textBox73.Text = (Math.Round(double.Parse(textBox90.Text) * double.Parse(textBox58.Text) / 1000, 2)).ToString();

            textBox72.Text = (Math.Round(double.Parse(textBox73.Text) + double.Parse(textBox68.Text) + double.Parse(textBox66.Text) + double.Parse(textBox67.Text), 2)).ToString();

            textBox71.Text = (Math.Round(double.Parse(textBox89.Text) * double.Parse(textBox58.Text) / 1000, 2)).ToString();

            textBox70.Text = (Math.Round(double.Parse(textBox72.Text) * 0.05, 2)).ToString();

            textBox69.Text = (Math.Round(double.Parse(textBox71.Text) * 0.1, 2)).ToString();

            //3 Раздел

            textBox77.Text = (Math.Round(double.Parse(textBox44.Text) * double.Parse(textBox6.Text) / 100, 2)).ToString();

            textBox76.Text = (Math.Round(double.Parse(textBox44.Text) * double.Parse(textBox4.Text) / 100, 2)).ToString();

            textBox75.Text = (Math.Round(double.Parse(textBox59.Text) * double.Parse(textBox6.Text) / 100, 2)).ToString();

            textBox74.Text = (Math.Round(double.Parse(textBox59.Text) * double.Parse(textBox4.Text) / 100, 2)).ToString();

            //4 Раздел

            textBox80.Text = (Math.Round(double.Parse(textBox71.Text) + double.Parse(textBox72.Text) + double.Parse(textBox44.Text) + double.Parse(textBox88.Text) + double.Parse(textBox62.Text) + double.Parse(textBox77.Text) + double.Parse(textBox76.Text), 2)).ToString();

            textBox79.Text = (Math.Round(double.Parse(textBox80.Text) * double.Parse(textBox18.Text) / 100, 2)).ToString();

            textBox78.Text = (Math.Round(double.Parse(textBox79.Text) + double.Parse(textBox80.Text), 2)).ToString();

            //5 Раздел

            textBox84.Text = (Math.Round(double.Parse(textBox69.Text) + double.Parse(textBox70.Text) + double.Parse(textBox59.Text) + double.Parse(textBox63.Text) + double.Parse(textBox64.Text) + double.Parse(textBox75.Text) + double.Parse(textBox74.Text), 2)).ToString();

            textBox83.Text = (Math.Round(double.Parse(textBox84.Text) * double.Parse(textBox18.Text) / 100, 2)).ToString();

            textBox82.Text = (Math.Round(double.Parse(textBox83.Text) + double.Parse(textBox84.Text), 2)).ToString();

            textBox81.Text = (Math.Round(double.Parse(textBox82.Text) - double.Parse(textBox69.Text) - double.Parse(textBox70.Text), 2)).ToString();

            textBox87.Text = (Math.Round(double.Parse(textBox81.Text) * double.Parse(textBox17.Text) / 100, 2)).ToString();

            textBox86.Text = (Math.Round((double.Parse(textBox87.Text) + double.Parse(textBox81.Text)) * 20 / 100, 2)).ToString();

            textBox85.Text = (Math.Round(double.Parse(textBox81.Text) + double.Parse(textBox87.Text) + double.Parse(textBox86.Text), 2)).ToString();

        }

    }
}
