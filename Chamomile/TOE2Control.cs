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
using System.Numerics;

namespace Chamomile
{
    public partial class TOE2Control : UserControl
    {
       
        private readonly string TemplateFileName = @"L:\Chamomile\exampleR2.docx";

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
            for (int i = 0; i < n; i++)
            {
                switch (sums[i])
                {
                    case 1:  //L
                        zs[i] = new Complex(0, valueXL);
                        break;
                    case 2:  //C
                        zs[i] = new Complex(0, -valueXC);
                        break;
                    case 3:  //LC
                        zs[i] = new Complex(0, valueXL - valueXC);
                        break;
                    case 4:  //R1 
                        zs[i] = new Complex(valueR1, 0);
                        break;
                    case 5:  //R1L 
                        zs[i] = new Complex(valueR1, valueXL);
                        break;
                    case 6:  //R1C 
                        zs[i] = new Complex(valueR1, -valueXC);
                        break;
                    case 8:  //R2 
                        zs[i] = new Complex(valueR2, 0);
                        break;
                    case 9:  //R2L
                        zs[i] = new Complex(valueR2, valueXL);
                        break;
                    case 10:  //R2C 
                        zs[i] = new Complex(valueR2, -valueXC);
                        break;
                    case 11:  //R2CL
                        zs[i] = new Complex(valueR2, valueXL - valueXC);
                        break;
                    case 12:  //R1R2 
                        zs[i] = new Complex(valueR1 + valueR2, 0);
                        break;
                    case 13:  //R1R2L
                        zs[i] = new Complex(valueR1 + valueR2, valueXL);
                        break;
                    case 14:  //R1R2C
                        zs[i] = new Complex(valueR1 + valueR2, -valueXC);
                        break;
                    case 15:  //R3
                        zs[i] = new Complex(valueR3, 0);
                        break;
                    case 16:  //R3L
                        zs[i] = new Complex(valueR3, valueXL);
                        break;
                    case 17:  //R3C
                        zs[i] = new Complex(valueR3, -valueXC);
                        break;
                    case 23:  //R2R3 
                        zs[i] = new Complex(valueR3 + valueR2, 0);
                        break;
                    case 25:  //R2R3C 
                        zs[i] = new Complex(valueR2 + valueR3, -valueXC);
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
                wordApp.Visible = true;
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);

                    wordDocument.SaveAs(@"L:\Chamomile\result12.docx");
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("ErrorWord");
                }
            }
            //Формировнаие отчета
            progressBar1.Visible = false;
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
