using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;
using Visio = Microsoft.Office.Interop.Visio;
using System.Numerics;

namespace Chamomile
{
    public partial class TOE2Control : UserControl
    {
       
        private readonly string TemplateFileName = @"L:\Chamomile\TOE2.docx";

        public TOE2Control()
        {
            InitializeComponent();
        }

        private void Calculation_Click(object sender, EventArgs e)
        {
            // исходные данные

            double valueU, valueR1, valueR2, valueR3, valueL = 0, valueC = 0;
            double valueXL, valueXC;
            
            try
            {
                string boof = U.Text;
                if (boof.Contains('.'))
                {
                    boof = boof.Replace('.', ',');
                }
                U.Text = boof;
                valueU = double.Parse(U.Text);

                boof = R1.Text;
                if (boof.Contains('.'))
                {
                    boof = boof.Replace('.', ',');
                }
                R1.Text = boof;
                valueR1 = double.Parse(R1.Text);

                boof = R2.Text;
                if (boof.Contains('.'))
                {
                    boof = boof.Replace('.', ',');
                }
                R2.Text = boof;
                valueR2 = double.Parse(R2.Text);

                
                if (R3.Text != "")
                {
                    boof = R3.Text;
                    if (boof.Contains('.'))
                    {
                        boof = boof.Replace('.', ',');
                    }
                    R3.Text = boof;
                    valueR3 = double.Parse(R3.Text);
                }
                else
                {
                    valueR3 = 0; 
                }

                boof = L.Text;
                if (boof.Contains('.'))
                {
                    boof = boof.Replace('.', ',');
                }
                L.Text = boof;
                valueL = double.Parse(L.Text);

                boof = C.Text;
                if (boof.Contains('.'))
                {
                    boof = boof.Replace('.', ',');
                }
                C.Text = boof;
                valueC = double.Parse(C.Text);
  
            }
            catch 
            {
                MessageBox.Show("Проверьте введенные данные", "Ошибка",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // исходные данные

            // исходная схема
            try
            {
                if (X1.Text == "" && X2.Text == "" && X3.Text == "" && X4.Text == "" && X5.Text == "" && X6.Text == "" && X7.Text == "" && X8.Text == "" && X9.Text == "" && X10.Text == "" && X11.Text == "" && X12.Text == "")
                {
                    throw new Exception("Все поля пустые");
                }
                if (X1.Text == "" && X2.Text == "" && X3.Text == "")
                {
                    throw new Exception("Все поля пустые");
                }
                if (X4.Text == "" && X5.Text == "" && X6.Text == "")
                {
                    throw new Exception("Все поля пустые");
                }
                if (X7.Text == "" && X8.Text == "" && X9.Text == "")
                {
                    throw new Exception("Все поля пустые");
                }
            }
            catch 
            {
                MessageBox.Show("Проверьте cхему", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // исходная схема


            //рассчитываем реактивные сопротивления
            valueXL = Math.Round(2 * Math.PI * 50 * valueL * 0.001, 2);
            valueXC = Math.Round(1000000 / (2 * Math.PI * 50 * valueC), 2);
            XL.Text = Convert.ToString(valueXL);
            XC.Text = Convert.ToString(valueXC);
            //рассчитываем реактивные сопротивления

            // заполяняем массив расположения элементов
            string[,] TI = new string[3, 4];
            TI[0, 0] = X1.Text;  //I1
            TI[1, 0] = X2.Text;  //I1
            TI[2, 0] = X3.Text;  //I1
            TI[0, 1] = X4.Text;  //I2
            TI[1, 1] = X5.Text;  //I2
            TI[2, 1] = X6.Text;  //I2
            TI[0, 2] = X7.Text;  //I3
            TI[1, 2] = X8.Text;  //I3
            TI[2, 2] = X9.Text;  //I3
            TI[0, 3] = X10.Text; //I4
            TI[1, 3] = X11.Text; //I4
            TI[2, 3] = X12.Text; //I4
                                 // заполяняем массив расположения элементов

            // заполняем массив расположения элементов числами
            int[,] TI1 = new int[3, 4];
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    switch (TI[i, j])
                    {
                        case "L":
                            TI1[i, j] = 1;
                            break;
                        case "C":
                            TI1[i, j] = 2;
                            break;
                        case "R1":
                            TI1[i, j] = 4;
                            break;
                        case "R2":
                            TI1[i, j] = 8;
                            break;
                        case "R3":
                            TI1[i, j] = 15;
                            break;
                        default:
                            TI1[i, j] = 0;
                            break;
                    }
                }
            }
            // заполняем массив расположения элементов числами

            // находим суммы в ветвях
            int[] sums = new int[4];
            sums[0] = TI1[0, 0] + TI1[1, 0] + TI1[2, 0];  //  I1
            sums[1] = TI1[0, 1] + TI1[1, 1] + TI1[2, 1];  //  I2
            sums[2] = TI1[0, 2] + TI1[1, 2] + TI1[2, 2];  //  I3
            sums[3] = TI1[0, 3] + TI1[1, 3] + TI1[2, 3];  //  I4
            // находим суммы в ветвях

            // определяем количество ветвей
            int n = 4;
            if (sums[3] == 0)
                n = 3;
            // определяем количество ветвей

            // рассчитываем сопротивления каждой ветви
            Complex[] zs = new Complex[6]; // комплексные значения cопротивлений
            String[] z0s = new String[4];  // 5 + 5 + j5
            String[] z1s = new String[4];  // 10 + j5
            String[] z3s = new String[4];  // L C R1
            for (int i = 0; i < n; i++)
            {
                switch (sums[i])
                {
                    case 1:  //L
                        zs[i] = new Complex(0, valueXL);
                        z0s[i] = "";
                        z1s[i] = "j" + Math.Round(zs[i].Imaginary, 2);
                        z3s[i] = "jXL";
                        break;
                    case 2:  //C
                        zs[i] = new Complex(0, -valueXC);
                        z0s[i] = "";
                        z1s[i] = "‒j" + Math.Round(-zs[i].Imaginary, 2);
                        z3s[i] = "‒jXC";
                        break;
                    case 3:  //LC
                        zs[i] = new Complex(0, valueXL - valueXC);
                        z0s[i] = "j" + valueXL + " ‒ j" + valueXC;
                        if (valueXL > valueXC)
                        {
                            z1s[i] = "j" + Math.Round(zs[i].Imaginary, 2);
                        }
                        else
                        {
                            z1s[i] = "-j" + Math.Round(-zs[i].Imaginary, 2);
                        }
                        z3s[i] = "jXL ‒ jXC";
                        break;
                    case 4:  //R1 
                        zs[i] = new Complex(valueR1, 0);
                        z0s[i] = "";
                        z1s[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                        z3s[i] = "R1";
                        break;
                    case 5:  //R1L 
                        zs[i] = new Complex(valueR1, valueXL);
                        z0s[i] = "";
                        z1s[i] = Math.Round(zs[i].Real, 2) + " + j" + Math.Round(zs[i].Imaginary, 2);
                        z3s[i] = "R1 + jXL";
                        break;
                    case 6:  //R1C 
                        zs[i] = new Complex(valueR1, -valueXC);
                        z0s[i] = "";
                        z1s[i] = Math.Round(zs[i].Real, 2) + " ‒ j" + Math.Round(-zs[i].Imaginary, 2);
                        z3s[i] = "R1 ‒ jXC";
                        break;
                    case 8:  //R2 
                        zs[i] = new Complex(valueR2, 0);
                        z0s[i] = "";
                        z1s[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                        z3s[i] = "R2";
                        break;
                    case 9:  //R2L
                        zs[i] = new Complex(valueR2, valueXL);
                        z0s[i] = "";
                        z1s[i] = Math.Round(zs[i].Real, 2) + " + j" + Math.Round(zs[i].Imaginary, 2);
                        z3s[i] = "R2 + jXL";
                        break;
                    case 10:  //R2C 
                        zs[i] = new Complex(valueR2, -valueXC);
                        z0s[i] = "";
                        z1s[i] = Math.Round(zs[i].Real, 2) + " ‒ j" + Math.Round(-zs[i].Imaginary, 2);
                        z3s[i] = "R2 ‒ jXC";
                        break;
                    case 11:  //R2CL
                        zs[i] = new Complex(valueR2, valueXL - valueXC);
                        z0s[i] = Math.Round(zs[i].Real, 2) + " + j" + Math.Round(valueXL, 2) + " ‒ j" + Math.Round(valueXC, 2);
                        if (valueXL > valueXC)
                        {
                            z1s[i] = zs[i].Real + " + j" + zs[i].Imaginary;
                        }
                        else
                        {
                            z1s[i] = zs[i].Real + " ‒ j" + -zs[i].Imaginary;
                        }
                        z3s[i] = "R2 + jXL ‒ jXC";
                        break;
                    case 12:  //R1R2 
                        zs[i] = new Complex(valueR1 + valueR2, 0);
                        z0s[i] = valueR1 + " + " + valueR2;
                        z1s[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                        z3s[i] = "R1 + R2";
                        break;
                    case 13:  //R1R2L
                        zs[i] = new Complex(valueR1 + valueR2, valueXL);
                        z0s[i] = valueR1 + " + " + valueR2 + " + j" + zs[i].Imaginary;
                        z1s[i] = Math.Round(zs[i].Real, 2) + " + j" + Math.Round(zs[i].Imaginary, 2);
                        z3s[i] = "R1 + R2 + jXL";
                        break;
                    case 14:  //R1R2C
                        zs[i] = new Complex(valueR1 + valueR2, -valueXC);
                        z0s[i] = valueR1 + " + " + valueR2 + " ‒ j" + -zs[i].Imaginary;
                        z1s[i] = Math.Round(zs[i].Real, 2) + " ‒ j" + Math.Round(-zs[i].Imaginary, 2);
                        z3s[i] = "R1 + R2 ‒ jXC";
                        break;
                    case 15:  //R3
                        zs[i] = new Complex(valueR3, 0);
                        z0s[i] = "";
                        z1s[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                        z3s[i] = "R3";
                        break;
                    case 16:  //R3L
                        zs[i] = new Complex(valueR3, valueXL);
                        z0s[i] = "";
                        z1s[i] = Math.Round(zs[i].Real, 2) + " + j" + Math.Round(zs[i].Imaginary, 2);
                        z3s[i] = "R3 + jXL";
                        break;
                    case 17:  //R3C
                        zs[i] = new Complex(valueR3, -valueXC);
                        z0s[i] = "";
                        z1s[i] = Math.Round(zs[i].Real, 2) + " ‒ j" + Math.Round(-zs[i].Imaginary, 2);
                        z3s[i] = "R3 ‒ jXC";
                        break;
                    case 23:  //R2R3 
                        zs[i] = new Complex(valueR3 + valueR2, 0);
                        z0s[i] = valueR2 + " + " + valueR3;
                        z1s[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                        z3s[i] = "R2 + R3";
                        break;
                    case 25:  //R2R3C 
                        zs[i] = new Complex(valueR2 + valueR3, -valueXC);
                        z0s[i] = valueR2 + " + " + valueR3 + " ‒ j" + -zs[i].Imaginary;
                        z1s[i] = Math.Round(zs[i].Real, 2) + " ‒ j" + Math.Round(-zs[i].Imaginary, 2);
                        z3s[i] = "R2 + R3 ‒ jXC";
                        break;
                }
            }
            // рассчитываем сопротивления каждой ветви

            // находим Zобщ
            zs[4] = (zs[1] * zs[2]) / (zs[1] + zs[2]); //Z23
            zs[5] = zs[4] + zs[0] + zs[3];             // Zобщ
            // находим Zобщ

            // вычисляем токи
            Complex[] I = new Complex[3];
            Complex U1 = new Complex(valueU, 0); // напряжение U
            I[0] = U1 / zs[5];                   //Iобщ
            Complex UAB = I[0] * zs[4];          // напряжение между АВ
            I[1] = UAB / zs[1];                  // I2
            I[2] = UAB / zs[2];                  // I3
            // вычисляем токи

            //  комплексная мощность цепи
            Complex IS = new Complex(I[0].Real, -I[0].Imaginary);  // Iобщ*
            Complex S = IS * U1;                                   // Sист
            double P = 0;
            int IR1 = 0;
            int IR2 = 0;
            int IR3 = 0;
            if (TI[0, 0] == "R1" ^ TI[1, 0] == "R1" ^ TI[2, 0] == "R1")
            {
                P += I[0].Magnitude * I[0].Magnitude * valueR1; // UR1
                IR1 = 1;
            }
            if (TI[0, 1] == "R1" ^ TI[1, 1] == "R1" ^ TI[2, 1] == "R1")
            {
                P += I[1].Magnitude * I[1].Magnitude * valueR1; // UR1
                IR1 = 2;
            }
            if (TI[0, 2] == "R1" ^ TI[1, 2] == "R1" ^ TI[2, 2] == "R1")
            {
                P += I[2].Magnitude * I[2].Magnitude * valueR1; // UR1
                IR1 = 3;
            }
            if (TI[0, 3] == "R1" ^ TI[1, 3] == "R1" ^ TI[2, 3] == "R1")
            {
                P += I[0].Magnitude * I[0].Magnitude * valueR1; // UR1
                IR1 = 4;
            }

            if (TI[0, 0] == "R2" ^ TI[1, 0] == "R2" ^ TI[2, 0] == "R2")
            {
                P += I[0].Magnitude * I[0].Magnitude * valueR2; // UR2
                IR2 = 1;
            }
            if (TI[0, 1] == "R2" ^ TI[1, 1] == "R2" ^ TI[2, 1] == "R2")
            {
                P += I[1].Magnitude * I[1].Magnitude * valueR2; // UR2
                IR2 = 2;
            }
            if (TI[0, 2] == "R2" ^ TI[1, 2] == "R2" ^ TI[2, 2] == "R2")
            {
                P += I[2].Magnitude * I[2].Magnitude * valueR2; // UR2
                IR2 = 3;
            }
            if (TI[0, 3] == "R2" ^ TI[1, 3] == "R2" ^ TI[2, 3] == "R2")
            {
                P += I[0].Magnitude * I[0].Magnitude * valueR2; // UR2
                IR2 = 4;
            }
            if (valueR3 != 0)
            {
                if (TI[0, 0] == "R3" ^ TI[1, 0] == "R3" ^ TI[2, 0] == "R3")
                {
                    P += I[0].Magnitude * I[0].Magnitude * valueR3; // UR3
                    IR3 = 1;
                }
                if (TI[0, 1] == "R3" ^ TI[1, 1] == "R3" ^ TI[2, 1] == "R3")
                {
                    P += I[1].Magnitude * I[1].Magnitude * valueR3; // UR3
                    IR3 = 2;
                }
                if (TI[0, 2] == "R3" ^ TI[1, 2] == "R3" ^ TI[2, 2] == "R3")
                {
                    P += I[2].Magnitude * I[2].Magnitude * valueR3; // UR3
                    IR3 = 3;
                }
                if (TI[0, 3] == "R3" ^ TI[1, 3] == "R3" ^ TI[2, 3] == "R3")
                {
                    P += I[0].Magnitude * I[0].Magnitude * valueR3; // UR3
                    IR3 = 4;
                }
            }
            double Q = 0;
            int IXC = 0;
            int IXL = 0;
            if (TI[0, 0] == "C" ^ TI[1, 0] == "C" ^ TI[2, 0] == "C")
            {
                Q -= I[0].Magnitude * I[0].Magnitude * valueXC; // XC
                IXC = 1;
            }
            if (TI[0, 1] == "C" ^ TI[1, 1] == "C" ^ TI[2, 1] == "C")
            {
                Q -= I[1].Magnitude * I[1].Magnitude * valueXC; // XC
                IXC = 2;
            }
            if (TI[0, 2] == "C" ^ TI[1, 2] == "C" ^ TI[2, 2] == "C")
            {
                Q -= I[2].Magnitude * I[2].Magnitude * valueXC; // XC
                IXC = 3;
            }
            if (TI[0, 3] == "C" ^ TI[1, 3] == "C" ^ TI[2, 3] == "C")
            {
                Q -= I[0].Magnitude * I[0].Magnitude * valueXC; // XC
                IXC = 4;
            }

            if (TI[0, 0] == "L" ^ TI[1, 0] == "L" ^ TI[2, 0] == "L")
            {
                Q += I[0].Magnitude * I[0].Magnitude * valueXL; // XL
                IXL = 1;
            }
            if (TI[0, 1] == "L" ^ TI[1, 1] == "L" ^ TI[2, 1] == "L")
            {
                Q += I[1].Magnitude * I[1].Magnitude * valueXL; // XL
                IXL = 2;
            }
            if (TI[0, 2] == "L" ^ TI[1, 2] == "L" ^ TI[2, 2] == "L")
            {
                Q += I[2].Magnitude * I[2].Magnitude * valueXL; // XL
                IXL = 3;
            }
            if (TI[0, 3] == "L" ^ TI[1, 3] == "L" ^ TI[2, 3] == "L")
            {
                Q += I[0].Magnitude * I[0].Magnitude * valueXL; // XL
                IXL = 4;
            }
            double Sp = Math.Sqrt(P * P + Q * Q);
            //  комплексная мощность цепи

            //  напряжения на элементах
            double[] UV = new double[6];
            UV[0] = valueU;
            if (TI[0, 0] == "L" ^ TI[1, 0] == "L" ^ TI[2, 0] == "L")
                UV[1] = I[0].Magnitude * valueXL; // UXL
            if (TI[0, 1] == "L" ^ TI[1, 1] == "L" ^ TI[2, 1] == "L")
                UV[1] = I[1].Magnitude * valueXL; // UXL
            if (TI[0, 2] == "L" ^ TI[1, 2] == "L" ^ TI[2, 2] == "L")
                UV[1] = I[2].Magnitude * valueXL; // UXL
            if (TI[0, 3] == "L" ^ TI[1, 3] == "L" ^ TI[2, 3] == "L")
                UV[1] = I[0].Magnitude * valueXL; // UXL

            if (TI[0, 0] == "C" ^ TI[1, 0] == "C" ^ TI[2, 0] == "C")
                UV[2] = I[0].Magnitude * valueXC; // UXC
            if (TI[0, 1] == "C" ^ TI[1, 1] == "C" ^ TI[2, 1] == "C")
                UV[2] = I[1].Magnitude * valueXC; // UXC
            if (TI[0, 2] == "C" ^ TI[1, 2] == "C" ^ TI[2, 2] == "C")
                UV[2] = I[2].Magnitude * valueXC; // UXC
            if (TI[0, 3] == "C" ^ TI[1, 3] == "C" ^ TI[2, 3] == "C")
                UV[2] = I[0].Magnitude * valueXC; // UXC

            if (TI[0, 0] == "R1" ^ TI[1, 0] == "R1" ^ TI[2, 0] == "R1")
                UV[3] = I[0].Magnitude * valueR1; // UR1
            if (TI[0, 1] == "R1" ^ TI[1, 1] == "R1" ^ TI[2, 1] == "R1")
                UV[3] = I[1].Magnitude * valueR1; // UR1
            if (TI[0, 2] == "R1" ^ TI[1, 2] == "R1" ^ TI[2, 2] == "R1")
                UV[3] = I[2].Magnitude * valueR1; // UR1
            if (TI[0, 3] == "R1" ^ TI[1, 3] == "R1" ^ TI[2, 3] == "R1")
                UV[3] = I[0].Magnitude * valueR1; // UR1

            if (TI[0, 0] == "R2" ^ TI[1, 0] == "R2" ^ TI[2, 0] == "R2")
                UV[4] = I[0].Magnitude * valueR2; // UR2
            if (TI[0, 1] == "R2" ^ TI[1, 1] == "R2" ^ TI[2, 1] == "R2")
                UV[4] = I[1].Magnitude * valueR2; // UR2
            if (TI[0, 2] == "R2" ^ TI[1, 2] == "R2" ^ TI[2, 2] == "R2")
                UV[4] = I[2].Magnitude * valueR2; // UR2
            if (TI[0, 3] == "R2" ^ TI[1, 3] == "R2" ^ TI[2, 3] == "R2")
                UV[4] = I[0].Magnitude * valueR2; // UR2
            if (valueR3 != 0)
            {
                if (TI[0, 0] == "R3" ^ TI[1, 0] == "R3" ^ TI[2, 0] == "R3")
                    UV[5] = I[0].Magnitude * valueR3; // UR3
                if (TI[0, 1] == "R3" ^ TI[1, 1] == "R3" ^ TI[2, 1] == "R3")
                    UV[5] = I[1].Magnitude * valueR3; // UR3
                if (TI[0, 2] == "R3" ^ TI[1, 2] == "R3" ^ TI[2, 2] == "R3")
                    UV[5] = I[2].Magnitude * valueR3; // UR3
                if (TI[0, 3] == "R3" ^ TI[1, 3] == "R3" ^ TI[2, 3] == "R3")
                    UV[5] = I[0].Magnitude * valueR3; // UR3
            }
            //  напряжения на элементах

            //  Масштаб по току и напряжению
            float valueMI, valueMU;
            valueMI = Convert.ToSingle(MI.Text);
            valueMU = Convert.ToSingle(MU.Text);
            //  Масштаб по току и напряжению

            //  Длины векторов
            double[] LUV = new double[9];  // длины векторов
            LUV[0] = I[0].Magnitude / valueMI;  // LIобщ, I1
            LUV[1] = I[1].Magnitude / valueMI;  // LI2
            LUV[2] = I[2].Magnitude / valueMI;  // LI3
            LUV[3] = UV[0] / valueMU;           // LU
            LUV[4] = UV[1] / valueMU;           // LUXL
            LUV[5] = UV[2] / valueMU;           // LUXC
            LUV[6] = UV[3] / valueMU;           // LUR1
            LUV[7] = UV[4] / valueMU;           // LUR2
            if (valueR3 != 0)
                LUV[8] = UV[5] / valueMU;       // LUR3
            //  Длины векторов

            // Векторная
            textBox39.Text = X1.Text + X2.Text + X3.Text;
            textBox40.Text = X4.Text + X5.Text + X6.Text;
            textBox41.Text = X7.Text + X8.Text + X8.Text;
            textBox42.Text = X9.Text + X10.Text + X11.Text;
            // Векторная

            //Заполнение textbox
            XL.Text = (valueXL).ToString();
            XC.Text = (valueXC).ToString();
            
            Z1real.Text = (Math.Round(zs[0].Real, 2)).ToString();
            Z1imaginary.Text = (Math.Round(zs[0].Imaginary, 2)).ToString();
            Z1z.Text = (Math.Round(zs[0].Magnitude, 2)).ToString();
            Z1a.Text = (Math.Round((zs[0].Phase * 180) / Math.PI)).ToString();

            Z2real.Text = (Math.Round(zs[1].Real, 2)).ToString();
            Z2imaginary.Text = (Math.Round(zs[1].Imaginary, 2)).ToString();
            Z2z.Text = (Math.Round(zs[1].Magnitude, 2)).ToString();
            Z2a.Text = (Math.Round((zs[1].Phase * 180) / Math.PI)).ToString();

            Z3real.Text = (Math.Round(zs[2].Real, 2)).ToString();
            Z3imaginary.Text = (Math.Round(zs[2].Imaginary, 2)).ToString();
            Z3z.Text = (Math.Round(zs[2].Magnitude, 2)).ToString();
            Z3a.Text = (Math.Round((zs[2].Phase * 180) / Math.PI)).ToString();

            if (n == 4)
            {
                Z4real.Text = (Math.Round(zs[3].Real, 2)).ToString();
                Z4imaginary.Text = (Math.Round(zs[3].Imaginary, 2)).ToString();
                Z4z.Text = (Math.Round(zs[3].Magnitude, 2)).ToString();
                Z4a.Text = (Math.Round((zs[3].Phase * 180) / Math.PI)).ToString();
            }

            Z23real.Text = (Math.Round(zs[4].Real, 2)).ToString();
            Z23imaginary.Text = (Math.Round(zs[4].Imaginary, 2)).ToString();
            Z23z.Text = (Math.Round(zs[4].Magnitude, 2)).ToString();
            Z23a.Text = (Math.Round((zs[4].Phase * 180) / Math.PI)).ToString();

            Zreal.Text = (Math.Round(zs[5].Real, 2)).ToString();
            Zimaginary.Text = (Math.Round(zs[5].Imaginary, 2)).ToString();
            Zz.Text = (Math.Round(zs[5].Magnitude, 2)).ToString();
            Za.Text = (Math.Round((zs[5].Phase * 180) / Math.PI)).ToString();

            Uz.Text = (Math.Round(U1.Magnitude, 2)).ToString();
            Ua.Text = (Math.Round(U1.Phase * 180) / Math.PI).ToString();
            Iz.Text = (Math.Round(I[0].Magnitude, 2)).ToString();
            Ia.Text = (Math.Round((I[0].Phase * 180) / Math.PI)).ToString();
            UABz.Text = (Math.Round(UAB.Magnitude, 2)).ToString();
            UABa.Text = (Math.Round((UAB.Phase * 180) / Math.PI)).ToString();
            I2z.Text = (Math.Round(I[1].Magnitude, 2)).ToString();
            I2a.Text = (Math.Round((I[1].Phase * 180) / Math.PI)).ToString();
            I3z.Text = (Math.Round(I[2].Magnitude, 2)).ToString();
            I3a.Text = (Math.Round((I[2].Phase * 180) / Math.PI)).ToString();

            Ssz.Text = (Math.Round(S.Magnitude, 2)).ToString();
            Ssa.Text = (Math.Round((S.Phase * 180) / Math.PI)).ToString();
            Ssreal.Text = (Math.Round(S.Real, 2)).ToString();
            Ssimaginary.Text = (Math.Round(S.Imaginary, 2)).ToString();
            Ps.Text = Ssreal.Text;
            Qs.Text = Ssimaginary.Text;
            Pc.Text = (Math.Round(P, 2)).ToString();
            Qc.Text = (Math.Round(Q, 2)).ToString();
            Sc.Text = (Math.Round(Sp, 2)).ToString();

            UXL.Text = (Math.Round(UV[1], 2)).ToString();
            UXC.Text = (Math.Round(UV[2], 2)).ToString();
            UR1.Text = (Math.Round(UV[3], 2)).ToString();
            UR2.Text = (Math.Round(UV[4], 2)).ToString();
            if (R3.Text != "")
                UR3.Text = (Math.Round(UV[5], 2)).ToString();

            LI.Text = (Math.Round(LUV[0], 2)).ToString();
            LI2.Text = (Math.Round(LUV[1], 2)).ToString();
            LI3.Text = (Math.Round(LUV[2], 2)).ToString();
            LU.Text = (Math.Round(LUV[3], 2)).ToString();
            LUXL.Text = (Math.Round(LUV[4], 2)).ToString();
            LUXC.Text = (Math.Round(LUV[5], 2)).ToString();
            LUR1.Text = (Math.Round(LUV[6], 2)).ToString();
            LUR2.Text = (Math.Round(LUV[7], 2)).ToString();
            if (R3.Text != "")
                LUR3.Text = (Math.Round(LUV[8], 2)).ToString();

            angleI.Text = (Math.Round((I[0].Phase * 180) / Math.PI)).ToString();
            angleI2.Text = (Math.Round((I[1].Phase * 180) / Math.PI)).ToString();
            angleI3.Text = (Math.Round((I[2].Phase * 180) / Math.PI)).ToString();
            angleU.Text = (Math.Round((U1.Phase * 180) / Math.PI)).ToString();
            //Заполнение textbox

           
            //Формирование отчета
            if (CalculationWithReport.Checked == true)
            {
                progressBar1.Value = 0;
                progressBar1.Visible = true;
                progressBar1.Value++;
                var wordApp = new Word.Application();
                wordApp.Visible = false;
                
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);
                    
                    ReplaceWordStub("{U}", U.Text, wordDocument);
                    ReplaceWordStub("{R1}", R1.Text, wordDocument);
                    ReplaceWordStub("{R2}", R2.Text, wordDocument);
                    progressBar1.Value++;
                    if (R3.Text != "")
                        ReplaceWordStub("{R3}", R3.Text, wordDocument);
                    else ReplaceWordStub("{R3}", "delete", wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{L}", L.Text, wordDocument);
                    ReplaceWordStub("{C}", C.Text, wordDocument);
                    ReplaceWordStub("{XL}", XL.Text, wordDocument);
                    ReplaceWordStub("{XC}", XC.Text, wordDocument);
                    progressBar1.Value++;
                    // Z1-Z4
                    ReplaceWordStub("{Z1e}", z3s[0], wordDocument);
                    ReplaceWordStub("{Z2e}", z3s[1], wordDocument);
                    ReplaceWordStub("{Z3e}", z3s[2], wordDocument);
                    progressBar1.Value++;
                    string[] Zer = new string[4];
                    for (int i = 0; i < 4; i++)
                    {
                        if (z0s[i] == "")
                            Zer[i] = z1s[i];
                        else Zer[i] = z0s[i] + " = " + z1s[i];
                    }
                    ReplaceWordStub("{Z1er}", Zer[0], wordDocument);
                    ReplaceWordStub("{Z2er}", Zer[1], wordDocument);
                    ReplaceWordStub("{Z3er}", Zer[2], wordDocument);
                    progressBar1.Value++;
                    string[] ZX = new string[4];
                    string[] ZY = new string[4];
                    string[] ZU = new string[4];
                    for (int i = 0; i < n; i++)
                    {
                        if (sums[i] == 5 ^ sums[i] == 9 ^ sums[i] == 13 ^ sums[i] == 16 ^ (sums[i] == 11 & valueXL > valueXC))
                        {
                            ZX[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                            ZY[i] = Convert.ToString(Math.Round(zs[i].Imaginary, 2));
                            ZU[i] = ZY[i];
                        }
                        else if (sums[i] == 6 ^ sums[i] == 10 ^ sums[i] == 14 ^ sums[i] == 17 ^ sums[i] == 25 ^ (sums[i] == 11 & valueXL < valueXC))
                        {
                            ZX[i] = Convert.ToString(Math.Round(zs[i].Real, 2));
                            ZY[i] = "(‒" + Math.Round(-zs[i].Imaginary, 2) + ")";
                            ZU[i] = "‒" + Math.Round(-zs[i].Imaginary, 2) + "";
                        }
                        else
                        {
                            ZX[i] = "delete";
                            ZY[i] = "delete";
                            ZU[i] = "delete";
                        }
                    }
                    ReplaceWordStub("{Z1X}", ZX[0], wordDocument);
                    ReplaceWordStub("{Z2X}", ZX[1], wordDocument);
                    ReplaceWordStub("{Z3X}", ZX[2], wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{Z4X}", ZX[3], wordDocument);
                    ReplaceWordStub("{Z1Y}", ZY[0], wordDocument);
                    ReplaceWordStub("{Z2Y}", ZY[1], wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{Z3Y}", ZY[2], wordDocument);
                    ReplaceWordStub("{Z4Y}", ZY[3], wordDocument);
                    ReplaceWordStub("{Z1U}", ZU[0], wordDocument);
                    ReplaceWordStub("{Z2U}", ZU[1], wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{Z3U}", ZU[2], wordDocument);
                    ReplaceWordStub("{Z4U}", ZU[3], wordDocument);

                    ReplaceWordStub("{Z1}", Convert.ToString(Math.Round(zs[0].Magnitude, 2)), wordDocument);
                    ReplaceWordStub("{Z2}", Convert.ToString(Math.Round(zs[1].Magnitude, 2)), wordDocument);
                    ReplaceWordStub("{Z3}", Convert.ToString(Math.Round(zs[2].Magnitude, 2)), wordDocument);
                    progressBar1.Value++;
                    string[] Zangle = new string[4];
                    for (int i = 0; i < 4; i++)
                    {
                        if (zs[i].Phase < 0)
                            Zangle[i] = "‒j" + Math.Round((-zs[i].Phase * 180) / Math.PI);
                        else Zangle[i] = "j" + Math.Round((zs[i].Phase * 180) / Math.PI);
                    }
                    ReplaceWordStub("{Z1angle}", Zangle[0], wordDocument);
                    ReplaceWordStub("{Z2angle}", Zangle[1], wordDocument);
                    ReplaceWordStub("{Z3angle}", Zangle[2], wordDocument);
                    progressBar1.Value++;
                    if (n == 3)
                    {
                        ReplaceWordStub("{Z4e}", "delete", wordDocument);
                        ReplaceWordStub("{Z4er}", "delete", wordDocument);
                        ReplaceWordStub("{Z4}", "delete", wordDocument);
                        ReplaceWordStub("{Z4angle}", "delete", wordDocument);
                    }
                    else
                    {
                        ReplaceWordStub("{Z4e}", z3s[3], wordDocument);
                        ReplaceWordStub("{Z4er}", Zer[3], wordDocument);
                        ReplaceWordStub("{Z4}", Convert.ToString(Math.Round(zs[3].Magnitude, 2)), wordDocument);
                        ReplaceWordStub("{Z4angle}", Zangle[3], wordDocument);
                    }
                    // Z1-Z4
                    progressBar1.Value++;
                    // Z23
                    string Z23plus = "";
                    if ((sums[2] == 2) ^ (sums[2] == 3 & valueXL < valueXC))
                        Z23plus = z1s[1] + " " + z1s[2];
                    else Z23plus = z1s[1] + " + " + z1s[2];
                    ReplaceWordStub("{Z2+Z3}", Z23plus, wordDocument);
                    progressBar1.Value++;
                    Complex Z23pr = zs[1] * zs[2];
                    ReplaceWordStub("{Z231}", Convert.ToString(Math.Round(Z23pr.Magnitude, 2)), wordDocument);

                    string Z23a = "";
                    if (Z23pr.Phase < 0)
                        Z23a = "‒j" + Math.Round((-Z23pr.Phase * 180) / Math.PI);
                    else Z23a = "j" + Math.Round((Z23pr.Phase * 180) / Math.PI);
                    ReplaceWordStub("{Z23a}", Z23a, wordDocument);
                    progressBar1.Value++;
                    Complex Z23s = zs[1] + zs[2];
                    string Z23ss = "";
                    double Z23p = Math.Round((Z23s.Phase * 180) / Math.PI);
                    if (Z23p == -90)
                        Z23ss = "‒j" + Math.Round(-Z23s.Imaginary, 2);
                    else if (Z23p == 90)
                        Z23ss = "j" + Math.Round(Z23s.Imaginary, 2);
                    else if (Z23p == 0)
                        Z23ss = "" + Math.Round(Z23s.Real, 2);
                    else if (Z23p < 0)
                        Z23ss = Math.Round(Z23s.Real, 2) + " ‒ j" + Math.Round(-Z23s.Imaginary, 2);
                    else if (Z23p > 0)
                        Z23ss = Math.Round(Z23s.Real, 2) + " + j" + Math.Round(Z23s.Imaginary, 2);
                    ReplaceWordStub("{Z23p}", Z23ss, wordDocument);

                    ReplaceWordStub("{Z23pr}", Convert.ToString(Math.Round(Z23s.Magnitude, 2)), wordDocument);
                    progressBar1.Value++;
                    string Z23pra = "";
                    if (Z23s.Phase < 0)
                        Z23pra = "‒j" + Math.Round((-Z23s.Phase * 180) / Math.PI);
                    else Z23pra = "j" + Math.Round((Z23s.Phase * 180) / Math.PI);
                    ReplaceWordStub("{Z23pra}", Z23pra, wordDocument);

                    ReplaceWordStub("{Z23}", Convert.ToString(Math.Round(zs[4].Magnitude, 2)), wordDocument);
                    progressBar1.Value++;
                    string Z23oa = "";
                    if (zs[4].Phase < 0)
                        Z23oa = "‒j" + Math.Round((-zs[4].Phase * 180) / Math.PI);
                    else Z23oa = "j" + Math.Round((zs[4].Phase * 180) / Math.PI);
                    ReplaceWordStub("{Z23oa}", Z23oa, wordDocument);

                    string Z23I = "";
                    double Z23q = Math.Round((zs[4].Phase * 180) / Math.PI);
                    if (Z23q == -90)
                        Z23I = "‒j" + Math.Round(-zs[4].Imaginary, 2);
                    else if (Z23q == 90)
                        Z23I = "j" + Math.Round(zs[4].Imaginary, 2);
                    else if (Z23q == 0)
                        Z23I = "" + Math.Round(zs[4].Real, 2);
                    else if (Z23q < 0)
                        Z23I = Math.Round(zs[4].Real, 2) + " ‒ j" + Math.Round(-zs[4].Imaginary, 2);
                    else if (Z23q > 0)
                        Z23I = Math.Round(zs[4].Real, 2) + " + j" + Math.Round(zs[4].Imaginary, 2);
                    ReplaceWordStub("{Z23I}", Z23I, wordDocument);
                    progressBar1.Value++;
                    string Z23y = "";
                    if (zs[4].Phase < 0)
                        Z23y = "‒" + Math.Round((-zs[4].Phase * 180) / Math.PI);
                    else Z23y = Math.Round((zs[4].Phase * 180) / Math.PI) + "";
                    ReplaceWordStub("{Z23y}", Z23y, wordDocument);
                    // Z23

                    // Zобщ
                    string Zo = "";
                    if (n == 3)
                        if (Z23q == -90)
                            Zo = Zer[0] + " " + Z23I;
                        else Zo = Zer[0] + " + " + Z23I;
                    if (n == 4)
                    {
                        if (Z23q == -90 & (sums[3] == 2) || (sums[3] == 3 & valueXL < valueXC))
                            Zo = Zer[0] + " " + Z23I + " " + Zer[3];
                        else if (Z23q == -90 & !((sums[3] == 2) || (sums[3] == 3 & valueXL < valueXC)))
                            Zo = Zer[0] + " " + Z23I + " + " + Zer[3];
                        else if (Z23q != -90 & !(sums[3] != 2 & (sums[3] != 3 & valueXL < valueXC)))
                            Zo = Zer[0] + " + " + Z23I + " + " + Zer[3];
                        else if (Z23q != -90 & sums[3] != 2 & (sums[3] != 3 & valueXL < valueXC))
                            Zo = Zer[0] + " + " + Z23I + " " + Zer[3];
                    }
                    ReplaceWordStub("{Zo}", Zo, wordDocument);
                    progressBar1.Value++;
                    string Zo1 = "";
                    double Zo11 = Math.Round((zs[5].Phase * 180) / Math.PI);
                    if (Zo11 == -90)
                        Zo1 = "‒j" + Math.Round(-zs[5].Imaginary, 2);
                    else if (Zo11 == 90)
                        Zo1 = "j" + Math.Round(zs[5].Imaginary, 2);
                    else if (Zo11 == 0)
                        Zo1 = "" + Math.Round(zs[5].Real, 2);
                    else if (Zo11 < 0)
                        Zo1 = Math.Round(zs[5].Real, 2) + " ‒ j" + Math.Round(-zs[5].Imaginary, 2);
                    else if (Zo11 > 0)
                        Zo1 = Math.Round(zs[5].Real, 2) + " + j" + Math.Round(zs[5].Imaginary, 2);
                    ReplaceWordStub("{Zo1}", Zo1, wordDocument);

                    string Zo2 = "";
                    if (zs[5].Real < 0)
                        Zo2 = "(" + Math.Round(zs[5].Real, 2) + ")";
                    else Zo2 = Convert.ToString(Math.Round(zs[5].Real, 2));
                    ReplaceWordStub("{Zo2}", Zo2, wordDocument);
                    progressBar1.Value++;
                    string Zo3 = "";
                    if (zs[5].Imaginary < 0)
                        Zo3 = "(‒" + Math.Round(-zs[5].Imaginary, 2) + ")";
                    else Zo3 = Convert.ToString(Math.Round(zs[5].Imaginary, 2));
                    ReplaceWordStub("{Zo3}", Zo3, wordDocument);

                    double Zo4 = Math.Round(zs[5].Real, 2);
                    string Zo41 = "";
                    if (Zo4 > 0)
                        Zo41 = Convert.ToString(Zo4);
                    else Zo41 = "‒" + -Zo4;
                    ReplaceWordStub("{Zo4}", Convert.ToString(Zo4), wordDocument);

                    double Zo5 = Math.Round(zs[5].Imaginary, 2);
                    string Zo51 = "";
                    if (Zo5 > 0)
                        Zo51 = Convert.ToString(Zo5);
                    else Zo51 = "‒" + -Zo5;
                    ReplaceWordStub("{Zo5}", Zo51, wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{Zo6}", Convert.ToString(Math.Round(zs[5].Magnitude, 2)), wordDocument);

                    string Zo7 = "";
                    if (zs[5].Phase < 0)
                        Zo7 = "‒j" + Math.Round((-zs[5].Phase * 180) / Math.PI);
                    else Zo7 = "j" + Math.Round((zs[5].Phase * 180) / Math.PI);
                    ReplaceWordStub("{Zo7}", Zo7, wordDocument);
                    // Zобщ
                    progressBar1.Value++;
                    // Токи
                    ReplaceWordStub("{I1}", Convert.ToString(Math.Round(I[0].Magnitude, 2)), wordDocument);

                    string Iangle = "";
                    if (I[0].Phase < 0)
                        Iangle = "‒j" + Math.Round((-I[0].Phase * 180) / Math.PI);
                    else Iangle = "j" + Math.Round((I[0].Phase * 180) / Math.PI);
                    ReplaceWordStub("{angleI}", Iangle, wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{UAB}", Convert.ToString(Math.Round(UAB.Magnitude, 2)), wordDocument);

                    string UABangle = "";
                    if (UAB.Phase < 0)
                        UABangle = "‒j" + Math.Round((-UAB.Phase * 180) / Math.PI);
                    else UABangle = "j" + Math.Round((UAB.Phase * 180) / Math.PI);
                    ReplaceWordStub("{UABangle}", UABangle, wordDocument);


                    ReplaceWordStub("{I2}", Convert.ToString(Math.Round(I[1].Magnitude, 2)), wordDocument);

                    string I2angle = "";
                    if (I[1].Phase < 0)
                        I2angle = "‒j" + Math.Round((-I[1].Phase * 180) / Math.PI);
                    else I2angle = "j" + Math.Round((I[1].Phase * 180) / Math.PI);
                    ReplaceWordStub("{angleI2}", I2angle, wordDocument);

                    ReplaceWordStub("{I3}", Convert.ToString(Math.Round(I[2].Magnitude, 2)), wordDocument);

                    string I3angle = "";
                    if (I[2].Phase < 0)
                        I3angle = "‒j" + Math.Round((-I[2].Phase * 180) / Math.PI);
                    else I3angle = "j" + Math.Round((I[2].Phase * 180) / Math.PI);
                    ReplaceWordStub("{angleI3}", I3angle, wordDocument);
                    // Токи
                    progressBar1.Value++;
                    // Баланс
                    string IIangle = "";
                    if (I[0].Phase < 0)
                        IIangle = "j" + Math.Round((-I[0].Phase * 180) / Math.PI);
                    else IIangle = "‒j" + Math.Round((I[0].Phase * 180) / Math.PI);
                    ReplaceWordStub("{IYangle}", IIangle, wordDocument);

                    ReplaceWordStub("{Ss}", Convert.ToString(Math.Round(S.Magnitude, 2)), wordDocument);

                    string SSangle = "";
                    if (S.Phase < 0)
                        SSangle = "‒j" + Math.Round((-S.Phase * 180) / Math.PI);
                    else SSangle = "j" + Math.Round((S.Phase * 180) / Math.PI);
                    ReplaceWordStub("{Ssangle}", SSangle, wordDocument);
                    progressBar1.Value++;
                    string Sangle = "";
                    if (S.Phase < 0)
                        Sangle = "‒" + Math.Round((-S.Phase * 180) / Math.PI);
                    else Sangle = "" + Math.Round((S.Phase * 180) / Math.PI);
                    ReplaceWordStub("{Sangle}", Sangle, wordDocument);

                    string Ssp = "";
                    double Sspq = Math.Round((S.Phase * 180) / Math.PI);
                    if (Sspq == -90)
                        Ssp = "‒j" + Math.Round(-S.Imaginary, 2);
                    else if (Sspq == 90)
                        Ssp = "j" + Math.Round(S.Imaginary, 2);
                    else if (Sspq == 0)
                        Ssp = "" + Math.Round(S.Real, 2);
                    else if (Sspq < 0)
                        Ssp = Math.Round(S.Real, 2) + " ‒ j" + Math.Round(-S.Imaginary, 2);
                    else if (Sspq > 0)
                        Ssp = Math.Round(S.Real, 2) + " + j" + Math.Round(S.Imaginary, 2);
                    ReplaceWordStub("{Ssp}", Ssp, wordDocument);
                    progressBar1.Value++;
                    string Ps = "";
                    if (S.Real < 0)
                        Ps = "‒ " + Math.Round(-S.Real, 2);
                    else Ps = Convert.ToString(Math.Round(S.Real, 2));
                    ReplaceWordStub("{Ps}", Ps, wordDocument);

                    string Qs = "";
                    if (S.Imaginary < 0)
                        Qs = "‒ " + Math.Round(-S.Imaginary, 2);
                    else Qs = Convert.ToString(Math.Round(S.Imaginary, 2));
                    ReplaceWordStub("{Qs}", Qs, wordDocument);

                    ReplaceWordStub("{Sp}", Convert.ToString(Math.Round(Sp, 2)), wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{Pp}", Convert.ToString(Math.Round(P, 2)), wordDocument);

                    string Qss = "";
                    if (Q < 0)
                        Qss = "‒ " + Math.Round(-Q, 2);
                    else Qss = Convert.ToString(Math.Round(Q, 2));
                    ReplaceWordStub("{Qp}", Qss, wordDocument);

                    string Pp1 = "";
                    if (P < 0)
                        Pp1 = "(‒ " + Math.Round(-P, 2) + ")";
                    else Pp1 = Convert.ToString(Math.Round(P, 2));
                    ReplaceWordStub("{Pp1}", Pp1, wordDocument);
                    progressBar1.Value++;
                    string Qq1 = "";
                    if (Q < 0)
                        Qq1 = "(‒ " + Math.Round(-Q, 2) + ")";
                    else Qq1 = Convert.ToString(Math.Round(Q, 2));
                    ReplaceWordStub("{Qq1}", Qq1, wordDocument);

                    string IXLS = Convert.ToString(IXL);
                    string IXCS = Convert.ToString(IXC);
                    string IR1S = Convert.ToString(IR1);
                    string IR2S = Convert.ToString(IR2);
                    string IR3S;
                    if (R3.Text != "")
                        IR3S = Convert.ToString(IR3);
                    else IR3S = "none";

                    ReplaceWordStub("{IXL}", IXLS, wordDocument);
                    ReplaceWordStub("{IXC}", IXCS, wordDocument);
                    ReplaceWordStub("{IR1}", IR1S, wordDocument);
                    ReplaceWordStub("{IR2}", IR2S, wordDocument);
                    ReplaceWordStub("{IR3}", IR3S, wordDocument);
                    progressBar1.Value++;
                    string IR1Z = "";
                    string IR2Z = "";
                    string IR3Z = "";
                    string IXLZ = "";
                    string IXCZ = "";
                    if (IXLS == "1" ^ IXLS == "4")
                        IXLZ = Convert.ToString(Math.Round(I[0].Magnitude, 2));
                    else if (IXLS == "2")
                        IXLZ = Convert.ToString(Math.Round(I[1].Magnitude, 2));
                    else if (IXLS == "3")
                        IXLZ = Convert.ToString(Math.Round(I[2].Magnitude, 2));

                    if (IXCS == "1" ^ IXCS == "4")
                        IXCZ = Convert.ToString(Math.Round(I[0].Magnitude, 2));
                    else if (IXCS == "2")
                        IXCZ = Convert.ToString(Math.Round(I[1].Magnitude, 2));
                    else if (IXCS == "3")
                        IXCZ = Convert.ToString(Math.Round(I[2].Magnitude, 2));

                    if (IR1S == "1" ^ IR1S == "4")
                        IR1Z = Convert.ToString(Math.Round(I[0].Magnitude, 2));
                    else if (IR1S == "2")
                        IR1Z = Convert.ToString(Math.Round(I[1].Magnitude, 2));
                    else if (IR1S == "3")
                        IR1Z = Convert.ToString(Math.Round(I[2].Magnitude, 2));

                    if (IR2S == "1" ^ IR2S == "4")
                        IR2Z = Convert.ToString(Math.Round(I[0].Magnitude, 2));
                    else if (IR2S == "2")
                        IR2Z = Convert.ToString(Math.Round(I[1].Magnitude, 2));
                    else if (IR2S == "3")
                        IR2Z = Convert.ToString(Math.Round(I[2].Magnitude, 2));

                    if (R3.Text != "")
                    {
                        if (IR3S == "1" ^ IR3S == "4")
                            IR3Z = Convert.ToString(Math.Round(I[0].Magnitude, 2));
                        else if (IR3S == "2")
                            IR3Z = Convert.ToString(Math.Round(I[1].Magnitude, 2));
                        else if (IR3S == "3")
                            IR3Z = Convert.ToString(Math.Round(I[2].Magnitude, 2));
                    }
                    else IR3Z = "none";

                    ReplaceWordStub("{IXLZ}", IXLZ, wordDocument);
                    ReplaceWordStub("{IXCZ}", IXCZ, wordDocument);
                    ReplaceWordStub("{IR1Z}", IR1Z, wordDocument);
                    ReplaceWordStub("{IR2Z}", IR2Z, wordDocument);
                    ReplaceWordStub("{IR3Z}", IR3Z, wordDocument);
                    // Баланс
                    progressBar1.Value++;
                    // Векторная
                    ReplaceWordStub("{UXL}", Convert.ToString(Math.Round(UV[1], 2)), wordDocument);
                    ReplaceWordStub("{UXC}", Convert.ToString(Math.Round(UV[2], 2)), wordDocument);
                    ReplaceWordStub("{UR1}", Convert.ToString(Math.Round(UV[3], 2)), wordDocument);
                    ReplaceWordStub("{UR2}", Convert.ToString(Math.Round(UV[4], 2)), wordDocument);
                    if (R3.Text != "")
                        ReplaceWordStub("{UR3}", Convert.ToString(Math.Round(UV[5], 2)), wordDocument);
                    else ReplaceWordStub("{UR3}", "delete", wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{MI}", MI.Text, wordDocument);
                    ReplaceWordStub("{MU}", MU.Text, wordDocument);
                    ReplaceWordStub("{LI}", Convert.ToString(Math.Round(LUV[0], 2)), wordDocument);
                    ReplaceWordStub("{LI2}", Convert.ToString(Math.Round(LUV[1], 2)), wordDocument);
                    progressBar1.Value++;
                    ReplaceWordStub("{LI3}", Convert.ToString(Math.Round(LUV[2], 2)), wordDocument);
                    ReplaceWordStub("{LU}", Convert.ToString(Math.Round(LUV[3], 2)), wordDocument);
                    ReplaceWordStub("{LUXL}", Convert.ToString(Math.Round(LUV[4], 2)), wordDocument);
                    ReplaceWordStub("{LUXC}", Convert.ToString(Math.Round(LUV[5], 2)), wordDocument);
                    ReplaceWordStub("{LUR1}", Convert.ToString(Math.Round(LUV[6], 2)), wordDocument);
                    ReplaceWordStub("{LUR2}", Convert.ToString(Math.Round(LUV[7], 2)), wordDocument);
                    if (R3.Text != "")
                        ReplaceWordStub("{LUR3}", Convert.ToString(Math.Round(LUV[8], 2)), wordDocument);
                    else ReplaceWordStub("{LUR3}", "delete", wordDocument);
                    string Angle1 = "";
                    if (I[0].Phase < 0)
                        Angle1 = "‒" + Math.Round((-I[0].Phase * 180 / Math.PI));
                    else Angle1 = Math.Round((I[0].Phase * 180 / Math.PI)) + "";
                    ReplaceWordStub("{1angleI}", Angle1, wordDocument);

                    string Angle2 = "";
                    if (I[1].Phase < 0)
                        Angle2 = "‒" + Math.Round((-I[1].Phase * 180 / Math.PI));
                    else Angle2 = Math.Round((I[1].Phase * 180 / Math.PI)) + "";
                    ReplaceWordStub("{1angleI2}", Angle2, wordDocument);
                    string Angle3 = "";
                    if (I[2].Phase < 0)
                        Angle3 = "‒" + Math.Round((-I[2].Phase * 180 / Math.PI));
                    else Angle3 = Math.Round((I[2].Phase * 180 / Math.PI)) + "";
                    ReplaceWordStub("{1angleI3}", Angle3, wordDocument);

                    string rez = "";
                    if (R3.Text == "")
                        rez = "два";
                    else rez = "три";
                    ReplaceWordStub("{rez}", rez, wordDocument);

                    if (R3.Text != "")
                        ReplaceWordStub("{dot}", ";", wordDocument);
                    else ReplaceWordStub("{dot}", ".", wordDocument);

                    if (n == 3)
                        ReplaceWordStub("{dot1}", ".", wordDocument);
                    else ReplaceWordStub("{dot1}", ";", wordDocument);

                    // Векторная
                    wordDocument.SaveAs(@"D:\Work\ТОЭ\2 Раздел\TOE2.docx");
                    wordApp.Visible = true;
   
                }
                catch
                {
                    MessageBox.Show("ErrorWord");
                }
                
                string message;
                if (n == 3 && R3.Text == "")
                {
                    message = "Удалить выделенное зеленым и бирюзовым";
                }
                else if (n == 4 && R3.Text != "")
                {
                    message = "Убрать выделение";
                }
                else if (n == 4)
                {
                    message = "Убрать выделенное бирюзовым";
                }
                else message = "Убрать выделенное зеленым";

                MessageBox.Show(message, "Финальные правки Word", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Вывод в висио
                    Visio.Application app = new Visio.Application();
                    app.Visible = false;
                    Visio.Document doc = app.Documents.Open(@"L:\Chamomile\Example.vsdm");

                    try
                    {
                        Visio.Page visioPage = app.Application.ActivePage;
                        /// Диаграмма

                        // начальная координата(0,0)
                        double X0 = 7.767715 - 12 / 2.54;
                        double Y0 = 11.192912 - 16 / 2.54;
                        // начальная координата(0,0)

                        // расчет
                        double valueX1 = X0 + (LUV[0] * Math.Cos(I[0].Phase)) / 2.54;// конец Iобщ
                        double Y1 = Y0 + (LUV[0] * Math.Sin(I[0].Phase)) / 2.54;// конец Iобщ

                        double valueX2 = X0 + (LUV[1] * Math.Cos(I[1].Phase)) / 2.54;// конец I2
                        double Y2 = Y0 + (LUV[1] * Math.Sin(I[1].Phase)) / 2.54;// конец I2

                        double valueX3 = X0 + (LUV[2] * Math.Cos(I[2].Phase)) / 2.54;// конец I3
                        double Y3 = Y0 + (LUV[2] * Math.Sin(I[2].Phase)) / 2.54;// конец I3

                        string[] vector = new string[12];
                        vector[0] = X1.Text;    //I1
                        vector[1] = X2.Text;    //I1
                        vector[2] = X3.Text;    //I1
                        vector[3] = X4.Text;    //I2
                        vector[4] = X5.Text;    //I2
                        vector[5] = X6.Text;    //I2
                        vector[6] = X7.Text;    //I3
                        vector[7] = X8.Text;    //I3
                        vector[8] = X9.Text;    //I3
                        vector[9] = X10.Text;   //I4
                        vector[10] = X11.Text;  //I4
                        vector[11] = X11.Text;  //I4


                        int j = 0;

                        double[] XS = new double[6];
                        double[] YS = new double[6];
                        double[] XE = new double[6];
                        double[] YE = new double[6];
                        string[] VU = new string[6];
                        XS[0] = X0;
                        YS[0] = Y0;

                        // I1
                        for (int i = 0; i < 3; i++)
                        {
                            switch (vector[i])
                            {
                                case "":
                                    break;
                                case "R1":
                                    XE[j] = XS[j] + (LUV[6] * Math.Cos(I[0].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[6] * Math.Sin(I[0].Phase)) / 2.54;
                                    VU[j] = "UR1";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "R2":
                                    XE[j] = XS[j] + (LUV[7] * Math.Cos(I[0].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[7] * Math.Sin(I[0].Phase)) / 2.54;
                                    VU[j] = "UR2";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "R3":
                                    XE[j] = XS[j] + (LUV[8] * Math.Cos(I[0].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[8] * Math.Sin(I[0].Phase)) / 2.54;
                                    VU[j] = "UR3";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "L":
                                    XE[j] = XS[j] + (LUV[4] * Math.Cos(I[0].Phase + 90 * (Math.PI / 180))) / 2.54;
                                    YE[j] = YS[j] + (LUV[4] * Math.Sin(I[0].Phase + 90 * (Math.PI / 180))) / 2.54;
                                    VU[j] = "UXL";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "C":
                                    XE[j] = XS[j] + (LUV[5] * Math.Cos(I[0].Phase - 90 * (Math.PI / 180))) / 2.54;
                                    YE[j] = YS[j] + (LUV[5] * Math.Sin(I[0].Phase - 90 * (Math.PI / 180))) / 2.54;
                                    VU[j] = "UXC";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                            }

                        }
                        // I1
                        double Xprom = XS[j];
                        double Yprom = YS[j];
                        // I2
                        for (int i = 3; i < 6; i++)
                        {
                            switch (vector[i])
                            {
                                case "":
                                    break;
                                case "R1":
                                    XE[j] = XS[j] + (LUV[6] * Math.Cos(I[1].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[6] * Math.Sin(I[1].Phase)) / 2.54;
                                    VU[j] = "UR1";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "R2":
                                    XE[j] = XS[j] + (LUV[7] * Math.Cos(I[1].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[7] * Math.Sin(I[1].Phase)) / 2.54;
                                    VU[j] = "UR2";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "R3":
                                    XE[j] = XS[j] + (LUV[8] * Math.Cos(I[1].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[8] * Math.Sin(I[1].Phase)) / 2.54;
                                    VU[j] = "UR3";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "L":
                                    XE[j] = XS[j] + (LUV[4] * Math.Cos(I[1].Phase + 90 * (Math.PI / 180))) / 2.54;
                                    YE[j] = YS[j] + (LUV[4] * Math.Sin(I[1].Phase + 90 * (Math.PI / 180))) / 2.54;
                                    VU[j] = "UXL";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "C":
                                    XE[j] = XS[j] + (LUV[5] * Math.Cos(I[1].Phase - 90 * (Math.PI / 180))) / 2.54;
                                    YE[j] = YS[j] + (LUV[5] * Math.Sin(I[1].Phase - 90 * (Math.PI / 180))) / 2.54;
                                    VU[j] = "UXC";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                            }

                        }
                        // I2
                        XS[j] = Xprom;
                        YS[j] = Yprom;

                        // I3
                        for (int i = 6; i < 9; i++)
                        {
                            switch (vector[i])
                            {
                                case "":
                                    break;
                                case "R1":
                                    XE[j] = XS[j] + (LUV[6] * Math.Cos(I[2].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[6] * Math.Sin(I[2].Phase)) / 2.54;
                                    VU[j] = "UR1";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "R2":
                                    XE[j] = XS[j] + (LUV[7] * Math.Cos(I[2].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[7] * Math.Sin(I[2].Phase)) / 2.54;
                                    VU[j] = "UR2";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "R3":
                                    XE[j] = XS[j] + (LUV[8] * Math.Cos(I[2].Phase)) / 2.54;
                                    YE[j] = YS[j] + (LUV[8] * Math.Sin(I[2].Phase)) / 2.54;
                                    VU[j] = "UR3";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "L":
                                    XE[j] = XS[j] + (LUV[4] * Math.Cos(I[2].Phase + 90 * (Math.PI / 180))) / 2.54;
                                    YE[j] = YS[j] + (LUV[4] * Math.Sin(I[2].Phase + 90 * (Math.PI / 180))) / 2.54;
                                    VU[j] = "UXL";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                                case "C":
                                    XE[j] = XS[j] + (LUV[5] * Math.Cos(I[2].Phase - 90 * (Math.PI / 180))) / 2.54;
                                    YE[j] = YS[j] + (LUV[5] * Math.Sin(I[2].Phase - 90 * (Math.PI / 180))) / 2.54;
                                    VU[j] = "UXC";
                                    XS[j + 1] = XE[j];
                                    YS[j + 1] = YE[j];
                                    j++;
                                    break;
                            }

                        }
                        // I3

                        // I4
                        if (n == 4)
                        {
                            for (int i = 9; i < 12; i++)
                            {
                                switch (vector[i])
                                {
                                    case "":
                                        break;
                                    case "R1":
                                        XE[j] = XS[j] + (LUV[6] * Math.Cos(I[0].Phase)) / 2.54;
                                        YE[j] = YS[j] + (LUV[6] * Math.Sin(I[0].Phase)) / 2.54;
                                        VU[j] = "UR1";
                                        XS[j + 1] = XE[j];
                                        YS[j + 1] = YE[j];
                                        j++;
                                        break;
                                    case "R2":
                                        XE[j] = XS[j] + (LUV[7] * Math.Cos(I[0].Phase)) / 2.54;
                                        YE[j] = YS[j] + (LUV[7] * Math.Sin(I[0].Phase)) / 2.54;
                                        VU[j] = "UR2";
                                        XS[j + 1] = XE[j];
                                        YS[j + 1] = YE[j];
                                        j++;
                                        break;
                                    case "R3":
                                        XE[j] = XS[j] + (LUV[8] * Math.Cos(I[0].Phase)) / 2.54;
                                        YE[j] = YS[j] + (LUV[8] * Math.Sin(I[0].Phase)) / 2.54;
                                        VU[j] = "UR3";
                                        XS[j + 1] = XE[j];
                                        YS[j + 1] = YE[j];
                                        j++;
                                        break;
                                    case "L":
                                        XE[j] = XS[j] + (LUV[4] * Math.Cos(I[0].Phase + 90 * (Math.PI / 180))) / 2.54;
                                        YE[j] = YS[j] + (LUV[4] * Math.Sin(I[0].Phase + 90 * (Math.PI / 180))) / 2.54;
                                        VU[j] = "UXL";
                                        XS[j + 1] = XE[j];
                                        YS[j + 1] = YE[j];
                                        j++;
                                        break;
                                    case "C":
                                        XE[j] = XS[j] + (LUV[5] * Math.Cos(I[0].Phase - 90 * (Math.PI / 180))) / 2.54;
                                        YE[j] = YS[j] + (LUV[5] * Math.Sin(I[0].Phase - 90 * (Math.PI / 180))) / 2.54;
                                        VU[j] = "UXC";
                                        XS[j + 1] = XE[j];
                                        YS[j + 1] = YE[j];
                                        j++;
                                        break;
                                }

                            }
                        }
                        // I4
                        double Xmin = X0;
                        double Xmax = X0;
                        double Ymin = Y0;
                        double Ymax = Y0;

                        if (valueX1 < Xmin)
                            Xmin = valueX1;
                        if (valueX2 < Xmin)
                            Xmin = valueX2;
                        if (valueX3 < Xmin)
                            Xmin = valueX3;
                        if ((valueX3 - X0 + valueX2) < Xmin)
                            Xmin = (valueX3 - X0 + valueX2);
                        if ((X0 + LUV[3] / 2.54) < Xmin)
                            Xmin = X0 + LUV[3] / 2.54;
                        //
                        if (valueX1 > Xmax)
                            Xmax = valueX1;
                        if (valueX2 > Xmax)
                            Xmax = valueX2;
                        if (valueX3 > Xmax)
                            Xmax = valueX3;
                        if ((valueX3 - X0 + valueX2) > Xmax)
                            Xmax = (valueX3 - X0 + valueX2);
                        if ((X0 + LUV[3] / 2.54) > Xmax)
                            Xmax = X0 + LUV[3] / 2.54;
                        //
                        if (Y1 < Ymin)
                            Ymin = Y1;
                        if (Y2 < Ymin)
                            Ymin = Y2;
                        if (Y3 < Ymin)
                            Ymin = Y3;
                        if ((Y3 - Y0 + Y2) < Ymin)
                            Ymin = (Y3 - Y0 + Y2);

                        //
                        if (Y1 > Ymax)
                            Ymax = Y1;
                        if (Y2 > Ymax)
                            Ymax = Y2;
                        if (Y3 > Ymax)
                            Ymax = Y3;
                        if ((Y3 - Y0 + Y2) > Ymax)
                            Ymax = (Y3 - Y0 + Y2);


                        int G = 4;
                        if (R3.Text != "")
                            G++;
                        for (int i = 0; i < G; i++)
                        {
                            if (Xmin > XS[i])
                                Xmin = XS[i];
                            if (Xmin > XE[i])
                                Xmin = XE[i];
                            if (Xmax < XS[i])
                                Xmax = XS[i];
                            if (Xmax < XE[i])
                                Xmax = XE[i];
                            /////
                            if (Ymin > YS[i])
                                Ymin = YS[i];
                            if (Ymin > YE[i])
                                Ymin = YE[i];
                            if (Ymax < YS[i])
                                Ymax = YS[i];
                            if (Ymax < YE[i])
                                Ymax = YE[i];
                        }
                        double X00 = (Xmax + Xmin) / 2;
                        double Y00 = (Ymax + Ymin) / 2;
                        double Y01 = Y00 - 0.05;
                        double Y02 = Y00 + 0.05;
                        double X01 = X00 - (18.3333 / 2) / 2.54;
                        double X02 = X00 + (18.3333 / 2) / 2.54;
                        // расчет

                        // вставка
                        Visio.Shape LineII = visioPage.DrawLine(X01, Y01, X01, Y02);

                        Visio.Shape LineIII = visioPage.DrawLine(X02, Y01, X02, Y02);


                        Visio.Shape LineP1 = visioPage.DrawLine(Xmin - 1 / 2.54, Y0, Xmax + 1 / 2.54, Y0);
                        LineP1.Text = "+1";

                        Visio.Shape LineJ = visioPage.DrawLine(X0, Ymin - 1 / 2.54, X0, Ymax + 1 / 2.54);
                        LineJ.Text = "j";

                        Visio.Shape LineI1 = visioPage.DrawLine(X0, Y0, valueX1, Y1);
                        LineI1.Text = "IОбщ,I1";


                        Visio.Shape LineI2 = visioPage.DrawLine(X0, Y0, valueX2, Y2);
                        LineI2.Text = "I2";

                        Visio.Shape LineI3 = visioPage.DrawLine(X0, Y0, valueX3, Y3);
                        LineI3.Text = "I3";

                        Visio.Shape LineI22 = visioPage.DrawLine(valueX3, Y3, valueX3 - X0 + valueX2, Y3 - Y0 + Y2);
                        LineI22.Text = "I2";

                        Visio.Shape LineU = visioPage.DrawLine(X0, Y0, X0 + LUV[3] / 2.54, Y0);
                        LineU.Text = "U";

                        Visio.Shape LineU1 = visioPage.DrawLine(XS[0], YS[0], XE[0], YE[0]);
                        LineU1.Text = VU[0];

                        Visio.Shape LineU2 = visioPage.DrawLine(XS[1], YS[1], XE[1], YE[1]);
                        LineU2.Text = VU[1];

                        Visio.Shape LineU3 = visioPage.DrawLine(XS[2], YS[2], XE[2], YE[2]);
                        LineU3.Text = VU[2];

                        Visio.Shape LineU4 = visioPage.DrawLine(XS[3], YS[3], XE[3], YE[3]);
                        LineU4.Text = VU[3];

                        if (R3.Text != "")
                        {
                            Visio.Shape LineU5 = visioPage.DrawLine(XS[4], YS[4], XE[4], YE[4]);
                            LineU5.Text = VU[4];
                        }
                        // вставка
                        /// Диаграмма

                        doc.SaveAs(@"L:\Chamomile\Result.vsdm");


                    if (n == 3 && R3.Text == "")
                    {
                        message = "Запустите макрос R2I";
                    }
                    else if (n == 4 && R3.Text != "")
                    {
                        message = "Запустите макрос R2I4";
                    }
                    else if (n == 4)
                    {
                        message = "Запустите макрос R3I4";
                    }
                    else message = "Запустите макрос R3I";

                    MessageBox.Show(message, "финальные правки Visio", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    progressBar1.Visible = false;
                    }
                    catch
                    {
                        MessageBox.Show("ErrorVisio");
                    }
                // Вывод в висио
            }
            //Формировнаие отчета
        }

        
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Word.WdReplace.wdReplaceAll);
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            XL.Clear();
            XC.Clear();
            Z1a.Clear();
            Z1imaginary.Clear();
            Z1real.Clear();
            Z1z.Clear();
            Z2a.Clear();
            Z2imaginary.Clear();
            Z2real.Clear();
            Z2z.Clear();
            Z3a.Clear();
            Z3imaginary.Clear();
            Z3real.Clear();
            Z3z.Clear();
            Z4a.Clear();
            Z4imaginary.Clear();
            Z4real.Clear();
            Z4z.Clear();
            Z23a.Clear();
            Z23imaginary.Clear();
            Z23real.Clear();
            Z23z.Clear();
            Za.Clear();
            Zimaginary.Clear();
            Zreal.Clear();
            Zz.Clear();
            Uz.Clear();
            Ua.Clear();
            Iz.Clear();
            Ia.Clear();
            UABa.Clear();
            UABz.Clear();
            I2a.Clear();
            I2z.Clear();
            I3a.Clear();
            I3z.Clear();
            Ssa.Clear();
            Ssimaginary.Clear();
            Ssreal.Clear();
            Ssz.Clear();
            Ps.Clear();
            Qs.Clear();
            Pc.Clear();
            Qc.Clear();
            Sc.Clear();
            UXL.Clear();
            UXC.Clear();
            UR1.Clear();
            UR2.Clear();
            UR3.Clear();
            LI.Clear();
            LI2.Clear();
            LI3.Clear();
            LU.Clear();
            LUXL.Clear();
            LUXC.Clear();
            LUR1.Clear();
            LUR2.Clear();
            LUR3.Clear();
            angleI.Clear();
            angleI2.Clear();
            angleI3.Clear();
            angleU.Clear();
            textBox39.Clear();
            textBox40.Clear();
            textBox41.Clear();
            textBox42.Clear();
        }
    }
}
