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
    public partial class DiagnosticsControl : UserControl
    {
        private readonly string TemplateFileName = @"L:\Chamomile\example.docx";
        public DiagnosticsControl()
        {
            InitializeComponent();
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
        }

        private void calculation_Click(object sender, EventArgs e)
        {
            //ПРоверка
            //textBox2.Text = "27";
            //comboBox1.Text = "Битумоминеральное";
            //comboBox2.Text = "Большой город";
            //comboBox3.Text = "Холмистый";
            //comboBox4.Text = "Умеренный";
            //textBox3.Text = "105";
            //textBox4.Text = "69";
            //comboBox5.Text = "1,8-3,5";
            //comboBox6.Text = "Осень-зима";

            if (comboBox5.Text == "1,2-1,8")
            {
                textBox6.Text = "150";
                textBox8.Text = "0,4";
                textBox7.Text = "2,6";
                textBox9.Text = "10,2";
                textBox10.Text = "3,4";

            }
            else 
            {
                textBox6.Text = "350";
                textBox8.Text = "0,3";
                textBox7.Text = "6";
                textBox9.Text = "16,9";
                textBox10.Text = "3";
            }
            

            //Заполнение коэффициентов

            textBox50.Text = "10";
            textBox5.Text = "20";

            double k1 = 0;

            int operatingCategory = 0;//Категория условий эксплуатации
            int roadSurface = 0, terrain = 0, drivingConditions = 0;

            if (comboBox1.Text == "Цементобетон" || comboBox1.Text == "Асфальтобетон" || comboBox1.Text == "Брусчатка" || comboBox1.Text == "Мозаика")
            {
                roadSurface = 1;
            }
            else if (comboBox1.Text == "Битумоминеральное")
            {
                roadSurface = 2;
            }
            else if (comboBox1.Text == "Щебень")
            {
                roadSurface = 3;
            }
            else if (comboBox1.Text == "Булыжник")
            {
                roadSurface = 4;
            }


            if (comboBox3.Text == "Равнинный")
            {
                terrain = 1;
            }
            else if (comboBox3.Text == "Слабохолмистый")
            {
                terrain = 2;
            }
            else terrain = 3;


            if (comboBox2.Text == "За городом")
            {
                drivingConditions = 1;
            }
            else if (comboBox2.Text == "Малый город")
            {
                drivingConditions = 2;
            }
            else drivingConditions = 3;



            if (drivingConditions == 1)
            {
                if (roadSurface == 1 && (terrain == 1 || terrain == 2))
                {
                    operatingCategory = 1;
                }
                else if ((roadSurface == 1 && terrain == 3) || (roadSurface == 2 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 3 && (terrain == 1 || terrain == 2 || terrain == 3)))
                {
                    operatingCategory = 2;
                }
                else if (roadSurface == 4 && (terrain == 1 || terrain == 2 || terrain == 3))
                {
                    operatingCategory = 3;
                }
                else if (roadSurface == 5 && (terrain == 1 || terrain == 2 || terrain == 3))
                {
                    operatingCategory = 4;
                }
            }
            else if (drivingConditions == 2)
            {
                if ((roadSurface == 1 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 2 && terrain == 1))
                {
                    operatingCategory = 2;
                }
                else if (roadSurface == 2 && (terrain == 2 || terrain == 3) || (roadSurface == 3 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 4 && (terrain == 1 || terrain == 2 || terrain == 3)))
                {
                    operatingCategory = 3;
                }
                else if (roadSurface == 5 && (terrain == 1 || terrain == 2 || terrain == 3))
                {
                    operatingCategory = 4;
                }
            }
            else if (drivingConditions == 3)
            {
                if ((roadSurface == 1 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 2 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 3 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 4 && terrain == 1))
                {
                    operatingCategory = 3;
                }
                else if ((roadSurface == 5 && (terrain == 1 || terrain == 2 || terrain == 3)) || (roadSurface == 4 && (terrain == 2 || terrain == 3)))
                {
                    operatingCategory = 4;
                }
            }

            switch (operatingCategory)
            {
                case 1:
                    k1 = 1;
                    break;
                case 2:
                    k1 = 0.9;
                    break;
                case 3:
                    k1 = 0.8;
                    break;
                case 4:
                    k1 = 0.7;
                    break;
            }

            textBox12.Text = k1.ToString();
            textBox11.Text = "1";

            if (comboBox4.Text == "Умеренный" || comboBox4.Text == "Умеренно теплый" || comboBox4.Text == "Умеренно влажный" || comboBox4.Text == "Теплый влажный")
            {
                textBox13.Text = "1";
            }
            else if (comboBox4.Text == "Жаркий сухой" || comboBox4.Text == "Очень жаркий сухой" || comboBox4.Text == "Умеренно холодный" || comboBox4.Text == "Холодный")
            {
                textBox13.Text = "0.9";
            }
            else if (comboBox4.Text == "Очень холодный")
            {
                textBox13.Text = "0.8";
            }

            double x = double.Parse(textBox4.Text) / double.Parse(textBox6.Text);
            textBox51.Text = Convert.ToString(Math.Round(x, 2));

            if (x < 0.5)
            {
                textBox14.Text = "0,7";
                textBox15.Text = "1";
            }
            else if (x < 0.75)
            {
                textBox14.Text = "1";
                textBox15.Text = "1,4";
            }
            else if (x < 1)
            {
                textBox14.Text = "1,3";
                textBox15.Text = "1,5";
            }
            else if (x < 1.25)
            {
                textBox14.Text = "1,4";
                textBox15.Text = "1,6";
            }
            else if (x < 1.5)
            {
                textBox14.Text = "1,4";
                textBox15.Text = "2";
            }
            else if (x < 1.75)
            {
                textBox14.Text = "1,4";
                textBox15.Text = "2,2";
            }
            else if (x < 2)
            {
                textBox14.Text = "1,4";
                textBox15.Text = "2,5";
            }
            else if (x >= 2)
            {
                textBox14.Text = "1,4";
                textBox15.Text = "2,7";
            }

            if (int.Parse(textBox2.Text) < 26)
            {
                textBox16.Text = "1,5";
            }
            else if (int.Parse(textBox2.Text) < 51)
            {
                textBox16.Text = "1,27";
            }
            else if (int.Parse(textBox2.Text) < 101)
            {
                textBox16.Text = "1,15";
            }
            else if (int.Parse(textBox2.Text) < 201)
            {
                textBox16.Text = "1";
            }
            else if (int.Parse(textBox2.Text) < 301)
            {
                textBox16.Text = "0,95";
            }
            else if (int.Parse(textBox2.Text) < 600)
            {
                textBox16.Text = "0,85";
            }
            else if (int.Parse(textBox2.Text) < 601)
            {
                textBox16.Text = "0,75";
            }

            if (comboBox6.Text == "Осень-зима")
            {
                textBox44.Text = "1,25";
            }
            else
            {
                textBox44.Text = "1";
            }
            //Заполнение коэффициентов 

            //Расчеты

            textBox17.Text = (Math.Round(1000 * double.Parse(textBox6.Text) * double.Parse(textBox12.Text) * double.Parse(textBox11.Text) * double.Parse(textBox13.Text), 2)).ToString();
            textBox18.Text = (Math.Round(1000 * double.Parse(textBox50.Text) * double.Parse(textBox12.Text) * double.Parse(textBox13.Text), 2)).ToString();
            textBox19.Text = (Math.Round(1000 * double.Parse(textBox5.Text) * double.Parse(textBox12.Text) * double.Parse(textBox13.Text), 2)).ToString();

            textBox20.Text = (Math.Round(double.Parse(textBox18.Text) / double.Parse(textBox3.Text))).ToString();
            textBox21.Text = (double.Parse(textBox3.Text) * double.Parse(textBox20.Text)).ToString();
            textBox22.Text = (Math.Round(double.Parse(textBox19.Text) / double.Parse(textBox21.Text))).ToString();
            textBox23.Text = (double.Parse(textBox21.Text) * double.Parse(textBox22.Text)).ToString();
            textBox24.Text = (Math.Round(double.Parse(textBox17.Text) / double.Parse(textBox23.Text))).ToString();
            textBox25.Text = (double.Parse(textBox23.Text) * double.Parse(textBox24.Text)).ToString();

            textBox26.Text = (Math.Round(double.Parse(textBox2.Text) * double.Parse(textBox3.Text) *365, 2)).ToString();

            textBox27.Text = (Math.Round(double.Parse(textBox26.Text) / (double.Parse(textBox6.Text)*1000), 2)).ToString();
            textBox39.Text = (Math.Round(double.Parse(textBox26.Text) / double.Parse(textBox23.Text) - double.Parse(textBox27.Text), 2)).ToString();
            textBox29.Text = (Math.Round(double.Parse(textBox26.Text) / double.Parse(textBox21.Text) - double.Parse(textBox27.Text) - double.Parse(textBox39.Text), 2)).ToString();
            textBox30.Text = (Math.Round(double.Parse(textBox26.Text) / double.Parse(textBox3.Text), 2)).ToString();
            textBox42.Text = (Math.Round(1.6 * (double.Parse(textBox29.Text) + double.Parse(textBox39.Text)), 2)).ToString();

            //Дальше для диплома

            textBox49.Text = (Math.Round(1.1 * double.Parse(textBox29.Text) + double.Parse(textBox39.Text), 2)).ToString();
            textBox32.Text = (Math.Round(1.2 * double.Parse(textBox39.Text), 2)).ToString();
            textBox33.Text = (Math.Round((double.Parse(textBox30.Text)/ 357), 2)).ToString();
            textBox34.Text = (Math.Round((double.Parse(textBox29.Text) / 255), 2)).ToString();
            textBox45.Text = (Math.Round((double.Parse(textBox39.Text) / 255), 2)).ToString();

            textBox46.Text = (Math.Round((double.Parse(textBox8.Text) * double.Parse(textBox11.Text) * double.Parse(textBox16.Text) * double.Parse(textBox44.Text) * 0.85), 2)).ToString();

            textBox35.Text = (Math.Round((double.Parse(textBox7.Text) * double.Parse(textBox11.Text) * double.Parse(textBox15.Text) * double.Parse(textBox16.Text) * double.Parse(textBox44.Text)), 2)).ToString();
            textBox36.Text = (Math.Round((double.Parse(textBox9.Text) * double.Parse(textBox11.Text) * double.Parse(textBox15.Text) * double.Parse(textBox16.Text) * double.Parse(textBox44.Text)), 2)).ToString();
            textBox47.Text = (Math.Round((double.Parse(textBox10.Text) * double.Parse(textBox11.Text) * double.Parse(textBox12.Text) * double.Parse(textBox13.Text) * double.Parse(textBox14.Text) * double.Parse(textBox16.Text) * double.Parse(textBox44.Text)), 2)).ToString();

            textBox37.Text = (Math.Round((double.Parse(textBox30.Text) * double.Parse(textBox46.Text)), 2)).ToString();
            textBox38.Text = (Math.Round((20 * double.Parse(textBox46.Text)), 2)).ToString();
            textBox48.Text = (Math.Round((double.Parse(textBox42.Text) * double.Parse(textBox38.Text)), 2)).ToString();
            textBox28.Text = (Math.Round((double.Parse(textBox29.Text) * double.Parse(textBox35.Text)), 2)).ToString();
            textBox40.Text = (Math.Round((double.Parse(textBox39.Text) * double.Parse(textBox36.Text)), 2)).ToString();

            textBox41.Text = (Math.Round((double.Parse(textBox28.Text) + 0.15), 2)).ToString();
            textBox43.Text = (Math.Round((double.Parse(textBox40.Text) + 0.15), 2)).ToString();

            textBox31.Text = (Math.Round(((double.Parse(textBox26.Text) / 1000) * double.Parse(textBox47.Text)), 2)).ToString();

            //Расчеты



            //Формировнаие отчета
            if (exportToWord.Checked == true)
            {
                var wordApp = new Word.Application();
                wordApp.Visible = true;
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);

                    ReplaceWordStub("{model}", textBox1.Text, wordDocument);
                    ReplaceWordStub("{count}", textBox2.Text, wordDocument);
                    ReplaceWordStub("{road}", comboBox1.Text, wordDocument);
                    ReplaceWordStub("{conditions}", comboBox2.Text, wordDocument);
                    ReplaceWordStub("{relief}", comboBox3.Text, wordDocument);
                    ReplaceWordStub("{climate}", comboBox4.Text, wordDocument);
                    ReplaceWordStub("{Lsp}", textBox3.Text, wordDocument);
                    ReplaceWordStub("{L}", textBox4.Text, wordDocument);

                    ReplaceWordStub("{Lnto1}", textBox50.Text, wordDocument);
                    ReplaceWordStub("{Lnto2}", textBox5.Text, wordDocument);
                    ReplaceWordStub("{Lnsp}", textBox6.Text, wordDocument);

                    ReplaceWordStub("{tneo}", textBox8.Text, wordDocument);
                    ReplaceWordStub("{tnto1}", textBox7.Text, wordDocument);
                    ReplaceWordStub("{tnto2}", textBox9.Text, wordDocument);
                    ReplaceWordStub("{tntr}", textBox10.Text, wordDocument);

                    ReplaceWordStub("{X}", textBox51.Text, wordDocument);

                    ReplaceWordStub("{k1}", textBox12.Text, wordDocument);
                    ReplaceWordStub("{k2}", textBox11.Text, wordDocument);
                    ReplaceWordStub("{k3}", textBox13.Text, wordDocument);
                    ReplaceWordStub("{k4}", textBox14.Text, wordDocument);
                    ReplaceWordStub("{k42}", textBox15.Text, wordDocument);
                    ReplaceWordStub("{k5}", textBox16.Text, wordDocument);
                    ReplaceWordStub("{k6}", textBox44.Text, wordDocument);

                    ReplaceWordStub("{L`sp}", textBox17.Text, wordDocument);
                    ReplaceWordStub("{L`to1}", textBox18.Text, wordDocument);
                    ReplaceWordStub("{L`to2}", textBox19.Text, wordDocument);

                    ReplaceWordStub("{L`to1/Lsp}", textBox20.Text, wordDocument);
                    ReplaceWordStub("{Lto1}", textBox21.Text, wordDocument);
                    ReplaceWordStub("{L`to2/Lto1}", textBox22.Text, wordDocument);
                    ReplaceWordStub("{Lto2}", textBox23.Text, wordDocument);
                    ReplaceWordStub("{L`sp/Lto2}", textBox24.Text, wordDocument);
                    ReplaceWordStub("{Lcp}", textBox25.Text, wordDocument);

                    ReplaceWordStub("{Lg}", textBox26.Text, wordDocument);

                    ReplaceWordStub("{Ngkr}", textBox27.Text, wordDocument);
                    ReplaceWordStub("{Ngto2}", textBox39.Text, wordDocument);
                    ReplaceWordStub("{Ngto1}", textBox29.Text, wordDocument);
                    ReplaceWordStub("{Ngeos}", textBox30.Text, wordDocument);

                    ReplaceWordStub("{Ngeot}", textBox42.Text, wordDocument);
                    ReplaceWordStub("{Ngd1}", textBox49.Text, wordDocument);
                    ReplaceWordStub("{Ngd2}", textBox32.Text, wordDocument);

                    ReplaceWordStub("{Nseos}", textBox33.Text, wordDocument);
                    ReplaceWordStub("{Nsto1}", textBox34.Text, wordDocument);
                    ReplaceWordStub("{Nsto2}", textBox45.Text, wordDocument);

                    ReplaceWordStub("{teo}", textBox46.Text, wordDocument);
                    ReplaceWordStub("{tto1}", textBox35.Text, wordDocument);
                    ReplaceWordStub("{tto2}", textBox36.Text, wordDocument);
                    ReplaceWordStub("{ttr}", textBox47.Text, wordDocument);

                    ReplaceWordStub("{teom}", textBox38.Text, wordDocument);
                    ReplaceWordStub("{Tgeos}", textBox37.Text, wordDocument);
                    ReplaceWordStub("{Tgeom}", textBox48.Text, wordDocument);
                    ReplaceWordStub("{Tgto1}", textBox28.Text, wordDocument);

                    ReplaceWordStub("{Tgto2}", textBox40.Text, wordDocument);
                    ReplaceWordStub("{Tgto1rmt}", textBox41.Text, wordDocument);
                    ReplaceWordStub("{Tgto2rmt}", textBox43.Text, wordDocument);
                    ReplaceWordStub("{Tgtr}", textBox31.Text, wordDocument);


                    wordDocument.SaveAs(@"L:\Chamomile\result.docx");
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("ErrorWord");
                }
            }

            //Формировнаие отчета
        }

        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Word.WdReplace.wdReplaceAll);
        }

        private void exportToWord_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox50.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox8.Text = "";
            textBox7.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox12.Text = "";
            textBox51.Text = "";
            textBox11.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox44.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";
            textBox24.Text = "";
            textBox25.Text = "";
            textBox26.Text = "";
            textBox27.Text = "";
            textBox39.Text = "";
            textBox29.Text = "";
            textBox30.Text = "";
            textBox42.Text = "";
            textBox49.Text = "";
            textBox32.Text = "";
            textBox33.Text = "";
            textBox34.Text = "";
            textBox45.Text = "";
            textBox46.Text = "";
            textBox35.Text = "";
            textBox36.Text = "";
            textBox47.Text = "";
            textBox37.Text = "";
            textBox38.Text = "";
            textBox48.Text = "";
            textBox28.Text = "";
            textBox40.Text = "";
            textBox41.Text = "";
            textBox31.Text = "";
            textBox43.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            comboBox6.Text = "";
        }
    }
}
