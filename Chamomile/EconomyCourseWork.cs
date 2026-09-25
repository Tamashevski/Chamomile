using MathNet.Numerics.Distributions;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

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

        private readonly string TemplateFileName = @"L:\Chamomile\Экономика курсовая\EconomicsCourseWorkTemplate.docx";
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




            if (checkBox3.Checked == true)
            {

                var wordApp = new Word.Application();
                wordApp.Visible = true;
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);

                    if (comboBox2.Text.Length > 0)
                    {
                        ReplaceWordStub("{name}", comboBox2.Text, wordDocument);
                    }

                    ReplaceWordStub("{theme}", textBox107.Text, wordDocument);
                    ReplaceInHeadersFooters("{theme}", textBox107.Text, wordDocument);

                    if (comboBox2.Text.Length > 0)
                    {
                        ReplaceWordStub("{group}", comboBox2.Text, wordDocument);
                        ReplaceInHeadersFooters("{group}", comboBox2.Text, wordDocument);
                        ReplaceInHeadersFooters("{group}", comboBox2.Text, wordDocument);
                    }

                    if (comboBox3.Text.Length > 0)
                    {
                        ReplaceWordStub("{option}", comboBox3.Text, wordDocument);
                        ReplaceInHeadersFooters("{option}", comboBox3.Text, wordDocument);
                        ReplaceInHeadersFooters("{option}", comboBox3.Text, wordDocument);
                    }

                    if (textBox105.Text.Length > 6)
                    {
                        string SNF = textBox105.Text.Substring(6) + " " + textBox105.Text.Substring(0, 6);
                        ReplaceWordStub("{student}", SNF, wordDocument);
                        ReplaceInHeadersFooters("{student1}", textBox105.Text, wordDocument);
                    }

                    if (comboBox1.Text.Length > 6)
                    {
                        string SNF = comboBox1.Text.Substring(6) + " " + comboBox1.Text.Substring(0, 6);
                        ReplaceWordStub("{teacher}", SNF, wordDocument);
                        ReplaceInHeadersFooters("{teacher1}", comboBox1.Text, wordDocument);
                        ReplaceInHeadersFooters("{teacher1}", comboBox1.Text, wordDocument);
                    }

                    ReplaceWordStub("{Lc}", textBox5.Text, wordDocument);
                    ReplaceWordStub("{D}", textBox104.Text, wordDocument);
                    ReplaceWordStub("{L}", textBox58.Text, wordDocument);
                    ReplaceWordStub("{Neo}", textBox103.Text, wordDocument);
                    ReplaceWordStub("{teo}", textBox1.Text, wordDocument);
                    ReplaceWordStub("{Teo1}", textBox57.Text, wordDocument);
                    ReplaceWordStub("{Tto1}", textBox56.Text, wordDocument);
                    ReplaceWordStub("{Tto2}", textBox55.Text, wordDocument);
                    ReplaceWordStub("{Nto1}", textBox102.Text, wordDocument);
                    ReplaceWordStub("{Nto2}", textBox101.Text, wordDocument);
                    ReplaceWordStub("{tto1}", textBox3.Text, wordDocument);
                    ReplaceWordStub("{tto2}", textBox2.Text, wordDocument);
                    ReplaceWordStub("{Nso}", textBox100.Text, wordDocument);
                    ReplaceWordStub("{tso}", textBox15.Text, wordDocument);
                    ReplaceWordStub("{Tso1}", textBox54.Text, wordDocument);
                    ReplaceWordStub("{Ttr1}", textBox53.Text, wordDocument);
                    ReplaceWordStub("{ttr}", textBox16.Text, wordDocument);
                    ReplaceWordStub("{L}", textBox58.Text, wordDocument);
                    ReplaceWordStub("{Tob}", textBox52.Text, wordDocument);
                    ReplaceWordStub("{Ppn}", textBox99.Text, wordDocument);
                    ReplaceWordStub("{Tpn}", textBox51.Text, wordDocument);

                    //2 Раздел

                    ReplaceWordStub("{CHTSeo}", textBox98.Text, wordDocument);
                    ReplaceWordStub("{CHTSto1}", textBox97.Text, wordDocument);
                    ReplaceWordStub("{CHTSto2}", textBox96.Text, wordDocument);
                    ReplaceWordStub("{CHTStr}", textBox95.Text, wordDocument);
                    ReplaceWordStub("{ZPeo}", textBox50.Text, wordDocument);
                    ReplaceWordStub("{ZPto1}", textBox49.Text, wordDocument);
                    ReplaceWordStub("{ZPto2}", textBox48.Text, wordDocument);
                    ReplaceWordStub("{ZPtr}", textBox47.Text, wordDocument);
                    ReplaceWordStub("{ZPpovr}", textBox46.Text, wordDocument);
                    ReplaceWordStub("{P}", textBox45.Text, wordDocument);
                    ReplaceWordStub("{ZPosn}", textBox44.Text, wordDocument);
                    ReplaceWordStub("{ZPodop}", textBox88.Text, wordDocument);
                    ReplaceWordStub("{ZPg}", textBox43.Text, wordDocument);
                    ReplaceWordStub("{Hosn}", textBox94.Text, wordDocument);
                    ReplaceWordStub("{Osn}", textBox62.Text, wordDocument);
                    ReplaceWordStub("{ZPrn}", textBox61.Text, wordDocument);
                    ReplaceWordStub("{Prn}", textBox60.Text, wordDocument);
                    ReplaceWordStub("{ZPosnr}", textBox59.Text, wordDocument);
                    ReplaceWordStub("{ZPdopr}", textBox63.Text, wordDocument);
                    ReplaceWordStub("{ZPgr}", textBox65.Text, wordDocument);
                    ReplaceWordStub("{Osnr}", textBox64.Text, wordDocument);
                    ReplaceWordStub("{Heo}", textBox93.Text, wordDocument);
                    ReplaceWordStub("{Meo}", textBox67.Text, wordDocument);
                    ReplaceWordStub("{Hto1}", textBox92.Text, wordDocument);
                    ReplaceWordStub("{Mto1}", textBox66.Text, wordDocument);
                    ReplaceWordStub("{Hto2}", textBox91.Text, wordDocument);
                    ReplaceWordStub("{Mto2}", textBox68.Text, wordDocument);
                    ReplaceWordStub("{Htr}", textBox90.Text, wordDocument);
                    ReplaceWordStub("{Mtr}", textBox73.Text, wordDocument);
                    ReplaceWordStub("{Zm}", textBox72.Text, wordDocument);
                    ReplaceWordStub("{Hzch}", textBox89.Text, wordDocument);
                    ReplaceWordStub("{Zzch}", textBox71.Text, wordDocument);
                    ReplaceWordStub("{Zmr}", textBox70.Text, wordDocument);
                    ReplaceWordStub("{Zzchr}", textBox69.Text, wordDocument);

                    //3 Раздел

                    ReplaceWordStub("{Peopr}", textBox6.Text, wordDocument);
                    ReplaceWordStub("{Popr}", textBox77.Text, wordDocument);
                    ReplaceWordStub("{Peohr}", textBox4.Text, wordDocument);
                    ReplaceWordStub("{Pohr}", textBox76.Text, wordDocument);
                    ReplaceWordStub("{Poprr}", textBox75.Text, wordDocument);
                    ReplaceWordStub("{Pohrr}", textBox74.Text, wordDocument);

                    //4 Раздел

                    ReplaceWordStub("{Cpr}", textBox80.Text, wordDocument);
                    ReplaceWordStub("{Hk}", textBox18.Text, wordDocument);
                    ReplaceWordStub("{Pk}", textBox79.Text, wordDocument);
                    ReplaceWordStub("{Cp}", textBox78.Text, wordDocument);

                    //5 Раздел

                    ReplaceWordStub("{Cprr}", textBox84.Text, wordDocument);
                    ReplaceWordStub("{Pkp}", textBox83.Text, wordDocument);
                    ReplaceWordStub("{Cpr3}", textBox82.Text, wordDocument);
                    ReplaceWordStub("{Cpr1}", textBox81.Text, wordDocument);
                    ReplaceWordStub("{R}", textBox17.Text, wordDocument);
                    ReplaceWordStub("{P1}", textBox87.Text, wordDocument);
                    ReplaceWordStub("{NDS}", textBox86.Text, wordDocument);
                    ReplaceWordStub("{T}", textBox85.Text, wordDocument);

                    double p11 = Math.Round(double.Parse(textBox72.Text) * 100/ double.Parse(textBox78.Text), 2);
                    double p21 = Math.Round(double.Parse(textBox44.Text) * 100 / double.Parse(textBox78.Text), 2);
                    double p31 = Math.Round(double.Parse(textBox77.Text) * 100 / double.Parse(textBox78.Text), 2);
                    double p41 = Math.Round(double.Parse(textBox76.Text) * 100 / double.Parse(textBox78.Text), 2);
                    ReplaceWordStub("{p11}", p11.ToString(), wordDocument);
                    ReplaceWordStub("{p21}", p21.ToString(), wordDocument);
                    ReplaceWordStub("{p31}", p31.ToString(), wordDocument);
                    ReplaceWordStub("{p41}", p41.ToString(), wordDocument);


                    wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    string path;
                    if (textBox105.Text.Length > 6)
                    {
                        path = @"L:\Chamomile\Экономика курсовая\EconomicsCourseWork" + textBox105.Text.Trim().Substring(6) + ".docx";
                    }
                    else
                    {
                        path = @"L:\Chamomile\Экономика курсовая\EconomicsCourseWork.docx";
                    }

                    if (File.Exists(path))
                    {
                        try
                        {
                            File.Delete(path);
                        }
                        catch 
                        {
                           
                        }
                    }    

                    wordDocument.SaveAs(path);
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

        private static void ReplaceInHeadersFooters(string stubToReplace, string text, Word.Document wordDocument)
        {
            object findText = stubToReplace;
            object replaceWith = text;
            object replace = 1; // WdReplace.wdReplaceOne (заменить только одно первое вхождение)
            object forward = true;
            object wrap = 1;    // WdFindWrap.wdFindContinue
            object tm = Type.Missing;

            // Проверяем только верхние колонтитулы (основной и первую страницу)
            Word.WdHeaderFooterIndex[] hfIndices = {
        Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
        Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage
    };

            foreach (Word.Section section in wordDocument.Sections)
            {
                foreach (var index in hfIndices)
                {
                    var header = section.Headers[index];
                    if (header.Exists)
                    {
                        if (section.Index > 1 && header.LinkToPrevious)
                            header.LinkToPrevious = false;

                        // 1. Ищем и заменяем в тексте колонтитула. Если нашли — сразу выходим.
                        if (header.Range.Find.Execute(ref findText, ref tm, ref tm, ref tm, ref tm, ref tm, ref forward, ref wrap, ref tm, ref replaceWith, ref replace, ref tm, ref tm, ref tm, ref tm))
                        {
                            return;
                        }

                        // 2. Ищем и заменяем в фигурах/штампе. Если нашли — сразу выходим.
                        if (ProcessShapesOnce(header.Shapes, findText, replaceWith, replace, forward, wrap, tm))
                        {
                            return;
                        }
                    }
                }
            }
        }

        // Поиск по фигурам с мгновенным выходом после первой же успешной замены
        private static bool ProcessShapesOnce(Word.Shapes shapes, object findText, object replaceWith, object replace, object forward, object wrap, object tm)
        {
            foreach (Word.Shape shape in shapes)
            {
                try
                {
                    if (shape.TextFrame.HasText != 0)
                    {
                        if (shape.TextFrame.TextRange.Find.Execute(
                            ref findText, ref tm, ref tm, ref tm, ref tm, ref tm,
                            ref forward, ref wrap, ref tm, ref replaceWith, ref replace,
                            ref tm, ref tm, ref tm, ref tm))
                        {
                            return true;
                        }
                    }
                }
                catch { }

                try
                {
                    dynamic dynShape = shape;
                    if (dynShape.Type == 6) // msoGroup (группа фигур)
                    {
                        foreach (Word.Shape subShape in dynShape.GroupItems)
                        {
                            if (subShape.TextFrame.HasText != 0)
                            {
                                if (subShape.TextFrame.TextRange.Find.Execute(
                                    ref findText, ref tm, ref tm, ref tm, ref tm, ref tm,
                                    ref forward, ref wrap, ref tm, ref replaceWith, ref replace,
                                    ref tm, ref tm, ref tm, ref tm))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            return false;
        }


    }
}
