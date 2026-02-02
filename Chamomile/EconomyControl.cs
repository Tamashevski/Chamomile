using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Chamomile
{
    public partial class EconomyControl : UserControl
    {
        

        private readonly string TemplateFileName = @"L:\Chamomile\example.docx";

        public EconomyControl()
        {
            InitializeComponent();
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
        }

        private void textBox5_Click(object sender, EventArgs e)
        {
            textBox5.Text = "";
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox6.Text = (Math.Round(double.Parse(textBox5.Text),2)).ToString();
            textBox8.Text = (Math.Round((double.Parse(textBox5.Text)+ double.Parse(textBox6.Text)), 2)).ToString();
            textBox7.Text = (Math.Round((double.Parse(textBox8.Text) * 0.34), 2)).ToString();
            textBox10.Text = (Math.Round((double.Parse(textBox8.Text) * 0.006), 2)).ToString();
            textBox14.Text = (Math.Round((double.Parse(textBox8.Text) * 0.8), 2)).ToString();
            textBox9.Text = (Math.Round((double.Parse(textBox8.Text) + double.Parse(textBox7.Text) + double.Parse(textBox10.Text)+ double.Parse(textBox14.Text)), 2)).ToString();
            textBox12.Text = (Math.Round((double.Parse(textBox9.Text) * 0.3), 2)).ToString();
            textBox11.Text = (Math.Round(((double.Parse(textBox12.Text) + double.Parse(textBox9.Text)) * 0.2), 2)).ToString();
            textBox13.Text = (Math.Round((double.Parse(textBox11.Text) + double.Parse(textBox12.Text) + double.Parse(textBox9.Text)), 2)).ToString();

            if (checkBox1.Checked == true)
            {
               
                var wordApp = new Word.Application();
                wordApp.Visible = true;
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);

                    ReplaceWordStub("{name}", textBox4.Text, wordDocument);
                    ReplaceWordStub("{comp1}", textBox2.Text, wordDocument);
                    ReplaceWordStub("{comp2}", textBox1.Text, wordDocument);
                    ReplaceWordStub("{org}", textBox3.Text, wordDocument);
                    ReplaceWordStub("{CHTS}", textBox5.Text, wordDocument);
                    ReplaceWordStub("{P}", textBox6.Text, wordDocument);
                    ReplaceWordStub("{ZP}", textBox8.Text, wordDocument);
                    ReplaceWordStub("{Ofszn}", textBox7.Text, wordDocument);
                    ReplaceWordStub("{Obgs}", textBox10.Text, wordDocument);
                    ReplaceWordStub("{Pnakl}", textBox14.Text, wordDocument);
                    ReplaceWordStub("{Snch}", textBox9.Text, wordDocument);
                    ReplaceWordStub("{Pr}", textBox12.Text, wordDocument);
                    ReplaceWordStub("{NDS}", textBox11.Text, wordDocument);
                    ReplaceWordStub("{S1nch}", textBox13.Text, wordDocument);

                    ReplaceWordStub("{mat11}", textBox15.Text, wordDocument);
                    ReplaceWordStub("{mat12}", textBox16.Text, wordDocument);

                    string s1 = textBox15.Text;
                    s1 = char.ToUpper(s1[0]) + s1.Substring(1);
                    string s2 = textBox16.Text;
                    s2 = char.ToUpper(s2[0]) + s2.Substring(1);

                    ReplaceWordStub("{mat11U}", s1, wordDocument);
                    ReplaceWordStub("{mat12U}", s2, wordDocument);
                    ReplaceWordStub("{ediz11}", comboBox1.Text, wordDocument);
                    ReplaceWordStub("{ediz12}", comboBox2.Text, wordDocument);

                    ReplaceWordStub("{norm11}", textBox19.Text, wordDocument);
                    ReplaceWordStub("{norm12}", textBox22.Text, wordDocument);
                    ReplaceWordStub("{cena11}", textBox20.Text, wordDocument);
                    ReplaceWordStub("{cena12}", textBox21.Text, wordDocument);
                    ReplaceWordStub("{summa11}", (Math.Round(double.Parse(textBox19.Text) * double.Parse(textBox20.Text), 2)).ToString(), wordDocument);
                    ReplaceWordStub("{summa12}", (Math.Round(double.Parse(textBox22.Text) * double.Parse(textBox21.Text), 2)).ToString(), wordDocument);
                    ReplaceWordStub("{summa1}", (Math.Round(double.Parse(textBox22.Text) * double.Parse(textBox21.Text) + double.Parse(textBox19.Text) * double.Parse(textBox20.Text), 2)).ToString(), wordDocument);

                    double summa1t = Math.Round((double.Parse(textBox22.Text) * double.Parse(textBox21.Text) + double.Parse(textBox19.Text) * double.Parse(textBox20.Text)) * 1.1, 2);

                    ReplaceWordStub("{summa1t}", summa1t.ToString(), wordDocument);

                    ReplaceWordStub("{mat21}", textBox18.Text, wordDocument);
                    ReplaceWordStub("{mat22}", textBox17.Text, wordDocument);

                    string s3 = textBox18.Text;
                    s3 = char.ToUpper(s3[0]) + s3.Substring(1);
                    string s4 = textBox17.Text;
                    s4 = char.ToUpper(s4[0]) + s4.Substring(1);

                    ReplaceWordStub("{mat21U}", s3, wordDocument);
                    ReplaceWordStub("{mat22U}", s4, wordDocument);
                    ReplaceWordStub("{ediz21}", comboBox3.Text, wordDocument);
                    ReplaceWordStub("{ediz22}", comboBox4.Text, wordDocument);

                    ReplaceWordStub("{norm21}", textBox24.Text, wordDocument);
                    ReplaceWordStub("{norm22}", textBox26.Text, wordDocument);
                    ReplaceWordStub("{cena21}", textBox23.Text, wordDocument);
                    ReplaceWordStub("{cena22}", textBox25.Text, wordDocument);
                    ReplaceWordStub("{summa21}", (Math.Round(double.Parse(textBox24.Text) * double.Parse(textBox23.Text), 2)).ToString(), wordDocument);
                    ReplaceWordStub("{summa22}", (Math.Round(double.Parse(textBox26.Text) * double.Parse(textBox25.Text), 2)).ToString(), wordDocument);
                    ReplaceWordStub("{summa2}", (Math.Round(double.Parse(textBox24.Text) * double.Parse(textBox23.Text) + double.Parse(textBox26.Text) * double.Parse(textBox25.Text), 2)).ToString(), wordDocument);

                    double summa2t = Math.Round((double.Parse(textBox24.Text) * double.Parse(textBox23.Text) + double.Parse(textBox26.Text) * double.Parse(textBox25.Text)) * 1.1, 2);

                    ReplaceWordStub("{summa2t}", summa2t.ToString(), wordDocument);

                    ReplaceWordStub("{det1}", textBox27.Text, wordDocument);
                    string s5 = textBox27.Text;
                    s5 = char.ToUpper(s5[0]) + s5.Substring(1);
                    ReplaceWordStub("{det1U}", s5, wordDocument);
                    ReplaceWordStub("{count1}", textBox29.Text, wordDocument);
                    ReplaceWordStub("{Stoim1}", textBox28.Text, wordDocument);
                    ReplaceWordStub("{sum1}", (Math.Round(double.Parse(textBox29.Text) * double.Parse(textBox28.Text), 2)).ToString(), wordDocument);

                    double sum1t = Math.Round(1.1 * double.Parse(textBox29.Text) * double.Parse(textBox28.Text), 2);

                    ReplaceWordStub("{sum1t}", sum1t.ToString(), wordDocument);

                    ReplaceWordStub("{det2}", textBox32.Text, wordDocument);
                    string s6 = textBox32.Text;
                    s6 = char.ToUpper(s6[0]) + s6.Substring(1);
                    ReplaceWordStub("{det2U}", s6, wordDocument);
                    ReplaceWordStub("{count2}", textBox31.Text, wordDocument);
                    ReplaceWordStub("{Stoim2}", textBox30.Text, wordDocument);
                    ReplaceWordStub("{sum2}", (Math.Round(double.Parse(textBox30.Text) * double.Parse(textBox31.Text), 2)).ToString(), wordDocument);

                    double sum2t = Math.Round(1.1 * double.Parse(textBox30.Text) * double.Parse(textBox31.Text), 2);

                    ReplaceWordStub("{sum2t}", sum2t.ToString(), wordDocument);

                    ReplaceWordStub("{t11}", textBox38.Text, wordDocument);
                    ReplaceWordStub("{t12}", textBox39.Text, wordDocument);
                    ReplaceWordStub("{t13}", textBox41.Text, wordDocument);
                    ReplaceWordStub("{t14}", textBox40.Text, wordDocument);
                    ReplaceWordStub("{t15}", textBox43.Text, wordDocument);
                    ReplaceWordStub("{t1}", (double.Parse(textBox38.Text) + double.Parse(textBox39.Text) + double.Parse(textBox40.Text) + double.Parse(textBox41.Text) + double.Parse(textBox43.Text)).ToString() , wordDocument);
                    ReplaceWordStub("{sodoper11}", textBox33.Text, wordDocument);
                    ReplaceWordStub("{sodoper12}", textBox34.Text, wordDocument);
                    ReplaceWordStub("{sodoper13}", textBox35.Text, wordDocument);
                    ReplaceWordStub("{sodoper14}", textBox36.Text, wordDocument);
                    ReplaceWordStub("{sodoper15}", textBox37.Text, wordDocument);

                    ReplaceWordStub("{t21}", textBox47.Text, wordDocument);
                    ReplaceWordStub("{t22}", textBox46.Text, wordDocument);
                    ReplaceWordStub("{t23}", textBox45.Text, wordDocument);
                    ReplaceWordStub("{t24}", textBox44.Text, wordDocument);
                    ReplaceWordStub("{t25}", textBox42.Text, wordDocument);
                    ReplaceWordStub("{t2}", (double.Parse(textBox47.Text) + double.Parse(textBox46.Text) + double.Parse(textBox45.Text) + double.Parse(textBox44.Text) + double.Parse(textBox42.Text)).ToString(), wordDocument);
                    ReplaceWordStub("{sodoper21}", textBox52.Text, wordDocument);
                    ReplaceWordStub("{sodoper22}", textBox51.Text, wordDocument);
                    ReplaceWordStub("{sodoper23}", textBox50.Text, wordDocument);
                    ReplaceWordStub("{sodoper24}", textBox49.Text, wordDocument);
                    ReplaceWordStub("{sodoper25}", textBox48.Text, wordDocument);

                    double Sr1 = Math.Round(double.Parse(textBox13.Text) * (double.Parse(textBox38.Text) + double.Parse(textBox39.Text) + double.Parse(textBox40.Text) + double.Parse(textBox41.Text) + double.Parse(textBox43.Text)), 2);
                    double Sr2 = Math.Round(double.Parse(textBox13.Text) * (double.Parse(textBox47.Text) + double.Parse(textBox46.Text) + double.Parse(textBox45.Text) + double.Parse(textBox44.Text) + double.Parse(textBox42.Text)), 2);

                    ReplaceWordStub("{Sr1}", Sr1.ToString(), wordDocument);
                    ReplaceWordStub("{Sr2}", Sr2.ToString(), wordDocument);

                    ReplaceWordStub("{all1}",(Sr1 + summa1t + sum1t).ToString(), wordDocument);
                    ReplaceWordStub("{all2}", (Sr2 + summa2t + sum2t).ToString(), wordDocument);

                    wordDocument.SaveAs(@"L:\Chamomile\result.docx");
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("ErrorWord");
                }
            }
        }

        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Word.WdReplace.wdReplaceAll);
        }

        private void textBox15_Click(object sender, EventArgs e)
        {
            textBox15.Text = "";
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
        }

        private void textBox19_Click(object sender, EventArgs e)
        {
            textBox19.Text = "";
        }

        private void textBox20_Click(object sender, EventArgs e)
        {
            textBox20.Text = "";
        }

        private void textBox16_Click(object sender, EventArgs e)
        {
            textBox16.Text = "";
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            comboBox2.Text = "";
        }

        private void textBox22_Click(object sender, EventArgs e)
        {
            textBox22.Text = "";
        }

        private void textBox21_Click(object sender, EventArgs e)
        {
            textBox21.Text = "";
        }

        private void textBox27_Click(object sender, EventArgs e)
        {
            textBox27.Text = "";
        }

        private void textBox29_Click(object sender, EventArgs e)
        {
            textBox29.Text = "";
        }

        private void textBox28_Click(object sender, EventArgs e)
        {
            textBox28.Text = "";
        }

        private void textBox33_Click(object sender, EventArgs e)
        {
            textBox33.Text = "";
        }

        private void textBox38_Click(object sender, EventArgs e)
        {
            textBox38.Text = "";
        }

        private void textBox34_Click(object sender, EventArgs e)
        {
            textBox34.Text = "";
        }

        private void textBox39_Click(object sender, EventArgs e)
        {
            textBox39.Text = "";
        }

        private void textBox35_Click(object sender, EventArgs e)
        {
            textBox35.Text = "";
        }

        private void textBox41_Click(object sender, EventArgs e)
        {
            textBox41.Text = "";
        }

        private void textBox36_Click(object sender, EventArgs e)
        {
            textBox36.Text = "";
        }

        private void textBox40_Click(object sender, EventArgs e)
        {
            textBox40.Text = "";
        }

        private void textBox37_Click(object sender, EventArgs e)
        {
            textBox37.Text = "";
        }

        private void textBox43_Click(object sender, EventArgs e)
        {
            textBox43.Text = "";
        }

        private void textBox32_Click(object sender, EventArgs e)
        {
            textBox32.Text = "";
        }

        private void textBox31_Click(object sender, EventArgs e)
        {
            textBox31.Text = "";
        }

        private void textBox30_Click(object sender, EventArgs e)
        {
            textBox30.Text = "";
        }

        private void textBox18_Click(object sender, EventArgs e)
        {
            textBox18.Text = "";
        }

        private void comboBox3_Click(object sender, EventArgs e)
        {
           comboBox3.Text = "";
        }

        private void textBox24_Click(object sender, EventArgs e)
        {
            textBox24.Text = "";
        }

        private void textBox23_Click(object sender, EventArgs e)
        {
            textBox23.Text = "";
        }

        private void textBox17_Click(object sender, EventArgs e)
        {
            textBox17.Text = "";
        }

        private void comboBox4_Click(object sender, EventArgs e)
        {
           comboBox4.Text = "";
        }

        private void textBox26_Click(object sender, EventArgs e)
        {
            textBox26.Text = "";
        }

        private void textBox25_Click(object sender, EventArgs e)
        {
            textBox25.Text = "";
        }

        private void textBox52_Click(object sender, EventArgs e)
        {
            textBox52.Text = "";
        }

        private void textBox47_Click(object sender, EventArgs e)
        {
            textBox47.Text = "";
        }

        private void textBox51_Click(object sender, EventArgs e)
        {
            textBox51.Text = "";
        }

        private void textBox46_Click(object sender, EventArgs e)
        {
            textBox46.Text = "";
        }

        private void textBox50_Click(object sender, EventArgs e)
        {
            textBox50.Text = "";
        }

        private void textBox45_Click(object sender, EventArgs e)
        {
            textBox45.Text = "";
        }

        private void textBox49_Click(object sender, EventArgs e)
        {
            textBox49.Text = "";
        }

        private void textBox44_Click(object sender, EventArgs e)
        {
            textBox44.Text = "";
        }

        private void textBox48_Click(object sender, EventArgs e)
        {
            textBox48.Text = "";
        }

        private void textBox42_Click(object sender, EventArgs e)
        {
            textBox42.Text = "";
        }
    }
}
