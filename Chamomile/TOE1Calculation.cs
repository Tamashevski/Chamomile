using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MathNet.Numerics.LinearAlgebra;

namespace Chamomile
{
    public class TOE1Calculation
    {
        public static void Calculate(TOE1Control toeControl)
        {
            double R1, R2, R3, R4, R5, R6, E1, E2, E3;

            double[] resistors = new double[6];
            double[] EMF = new double[3];

            resistors[0] = double.Parse(toeControl.TextBox4.Text);
            resistors[1] = double.Parse(toeControl.TextBox5.Text);
            resistors[2] = double.Parse(toeControl.TextBox7.Text);
            resistors[3] = double.Parse(toeControl.TextBox6.Text);
            resistors[4] = double.Parse(toeControl.TextBox9.Text);
            resistors[5] = double.Parse(toeControl.TextBox8.Text);

            if (toeControl.TextBox1.Text != "Введите")
            {
                EMF[0] = double.Parse(toeControl.TextBox1.Text);
            }

            if (toeControl.TextBox2.Text != "Введите")
            {
                EMF[1] = double.Parse(toeControl.TextBox2.Text);
            }

            if (toeControl.TextBox3.Text != "Введите")
            {
                EMF[2] = double.Parse(toeControl.TextBox3.Text);
            }

            //Method1(toeControl, resistors, EMF);
            Method2(toeControl, resistors, EMF);

        }

        public static void Method1(TOE1Control toeControl, double[] resistors, double[] EMF)
        {
            int boof;
            var mainMatrix = Matrix<double>.Build.DenseOfArray(new double[6, 7]);

            boof = int.Parse(toeControl.ComboBox4.Text);
            mainMatrix[0, boof - 1] = 1;
            if (toeControl.Label19.Text == "-")
                mainMatrix[0, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox5.Text);
            mainMatrix[0, boof - 1] = 1;
            if (toeControl.Label16.Text == "-")
                mainMatrix[0, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox6.Text);
            mainMatrix[0, boof - 1] = 1;
            if (toeControl.Label20.Text == "-")
                mainMatrix[0, boof - 1] *= -1;

            mainMatrix[0, 6] = 0;

            boof = int.Parse(toeControl.ComboBox9.Text);
            mainMatrix[1, boof - 1] = 1;
            if (toeControl.Label26.Text == "-")
                mainMatrix[1, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox8.Text);
            mainMatrix[1, boof - 1] = 1;
            if (toeControl.Label24.Text == "-")
                mainMatrix[1, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox7.Text);
            mainMatrix[1, boof - 1] = 1;
            if (toeControl.Label22.Text == "-")
                mainMatrix[1, boof - 1] *= -1;

            mainMatrix[1, 6] = 0;

            boof = int.Parse(toeControl.ComboBox12.Text);
            mainMatrix[2, boof - 1] = 1;
            if (toeControl.Label33.Text == "-")
                mainMatrix[2, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox11.Text);
            mainMatrix[2, boof - 1] = 1;
            if (toeControl.Label31.Text == "-")
                mainMatrix[2, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox10.Text);
            mainMatrix[2, boof - 1] = 1;
            if (toeControl.Label29.Text == "-")
                mainMatrix[2, boof - 1] *= -1;

            mainMatrix[2, 6] = 0;

            boof = int.Parse(toeControl.ComboBox15.Text);
            mainMatrix[3, boof - 1] = resistors[boof - 1];
            if (toeControl.Label40.Text == "-")
                mainMatrix[3, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox14.Text);
            mainMatrix[3, boof - 1] = resistors[boof - 1];
            if (toeControl.Label37.Text == "-")
                mainMatrix[3, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox18.Text);
            mainMatrix[3, boof - 1] = resistors[boof - 1];
            if (toeControl.Label44.Text == "-")
                mainMatrix[3, boof - 1] *= -1;

            if (toeControl.ComboBox20.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox20.Text);
                if (toeControl.Label48.Text == "+")
                    mainMatrix[3, 6] += EMF[boof - 1];
                else mainMatrix[3, 6] -= EMF[boof - 1];
            }

            if (toeControl.ComboBox19.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox19.Text);
                if (toeControl.Label46.Text == "+")
                    mainMatrix[3, 6] += EMF[boof - 1];
                else mainMatrix[3, 6] -= EMF[boof - 1];
            }

            ////////////////////////////////////

            boof = int.Parse(toeControl.ComboBox22.Text);
            mainMatrix[4, boof - 1] = resistors[boof - 1];
            if (toeControl.Label61.Text == "-")
                mainMatrix[4, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox16.Text);
            mainMatrix[4, boof - 1] = resistors[boof - 1];
            if (toeControl.Label54.Text == "-")
                mainMatrix[4, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox13.Text);
            mainMatrix[4, boof - 1] = resistors[boof - 1];
            if (toeControl.Label51.Text == "-")
                mainMatrix[4, boof - 1] *= -1;

            if (toeControl.ComboBox21.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox21.Text);
                if (toeControl.Label58.Text == "+")
                    mainMatrix[4, 6] += EMF[boof - 1];
                else mainMatrix[4, 6] -= EMF[boof - 1];
            }

            if (toeControl.ComboBox17.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox17.Text);
                if (toeControl.Label56.Text == "+")
                    mainMatrix[4, 6] += EMF[boof - 1];
                else mainMatrix[4, 6] -= EMF[boof - 1];
            }

            //////////////////////////////////

            boof = int.Parse(toeControl.ComboBox27.Text);
            mainMatrix[5, boof - 1] = resistors[boof - 1];
            if (toeControl.Label75.Text == "-")
                mainMatrix[5, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox24.Text);
            mainMatrix[5, boof - 1] = resistors[boof - 1];
            if (toeControl.Label68.Text == "-")
                mainMatrix[5, boof - 1] *= -1;

            boof = int.Parse(toeControl.ComboBox23.Text);
            mainMatrix[5, boof - 1] = resistors[boof - 1];
            if (toeControl.Label65.Text == "-")
                mainMatrix[5, boof - 1] *= -1;

            if (toeControl.ComboBox26.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox26.Text);
                if (toeControl.Label72.Text == "+")
                    mainMatrix[5, 6] += EMF[boof - 1];
                else mainMatrix[5, 6] -= EMF[boof - 1];
            }

            if (toeControl.ComboBox25.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox25.Text);
                if (toeControl.Label70.Text == "+")
                    mainMatrix[5, 6] += EMF[boof - 1];
                else mainMatrix[5, 6] -= EMF[boof - 1];
            }

            var matrix = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix[i, j] = mainMatrix[i, j];
                }
            }

            boof = 0;
            var matrix1 = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix1[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 6; i++)
            {
                matrix1[i, boof] = mainMatrix[i, 6];
            }
            boof++;

            var matrix2 = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix2[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 6; i++)
            {
                matrix2[i, boof] = mainMatrix[i, 6];
            }
            boof++;

            var matrix3 = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix3[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 6; i++)
            {
                matrix3[i, boof] = mainMatrix[i, 6];
            }
            boof++;

            var matrix4 = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix4[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 6; i++)
            {
                matrix4[i, boof] = mainMatrix[i, 6];
            }
            boof++;

            var matrix5 = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix5[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 6; i++)
            {
                matrix5[i, boof] = mainMatrix[i, 6];
            }
            boof++;

            var matrix6 = Matrix<double>.Build.DenseOfArray(new double[6, 6]);

            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    matrix6[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 6; i++)
            {
                matrix6[i, boof] = mainMatrix[i, 6];
            }

            double det = matrix.Determinant();
            double det1 = matrix1.Determinant();
            double det2 = matrix2.Determinant();
            double det3 = matrix3.Determinant();
            double det4 = matrix4.Determinant();
            double det5 = matrix5.Determinant();
            double det6 = matrix6.Determinant();

            double I1 = det1 / det;
            double I2 = det2 / det;
            double I3 = det3 / det;
            double I4 = det4 / det;
            double I5 = det5 / det;
            double I6 = det6 / det;

            toeControl.TextBox33.Text = I1.ToString();
            toeControl.TextBox36.Text = I2.ToString();
            toeControl.TextBox38.Text = I3.ToString();
            toeControl.TextBox44.Text = I4.ToString();
            toeControl.TextBox42.Text = I5.ToString();
            toeControl.TextBox40.Text = I6.ToString();

        }

        public static void Method2(TOE1Control toeControl, double[] resistors, double[] EMF)
        {
            int boof = 0;
            var mainMatrix = Matrix<double>.Build.DenseOfArray(new double[3, 4]);

            switch (toeControl.ComboBox60.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[0, boof] = resistors[int.Parse(toeControl.ComboBox57.Text)-1] + resistors[int.Parse(toeControl.ComboBox37.Text)-1] + resistors[int.Parse(toeControl.ComboBox38.Text) - 1];
            if (toeControl.Label117.Text == "-") mainMatrix[0, boof] *= -1;

            switch (toeControl.ComboBox42.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[0, boof] = resistors[int.Parse(toeControl.ComboBox80.Text)-1];
            if (toeControl.Label85.Text == "-") mainMatrix[0, boof] *= -1;

            switch (toeControl.ComboBox40.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[0, boof] = resistors[int.Parse(toeControl.ComboBox85.Text)-1];
            if (toeControl.Label82.Text == "-") mainMatrix[0, boof] *= -1;

            if (toeControl.ComboBox45.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox45.Text);
                if (toeControl.Label89.Text == "+")
                    mainMatrix[0, 3] += EMF[boof - 1];
                else mainMatrix[0, 3] -= EMF[boof - 1];
            }

            if (toeControl.ComboBox43.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox43.Text);
                if (toeControl.Label87.Text == "+")
                    mainMatrix[0, 3] += EMF[boof - 1];
                else mainMatrix[0, 3] -= EMF[boof - 1];
            }

            ///////////////////////////////////////////////

            switch (toeControl.ComboBox52.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[1, boof] = resistors[int.Parse(toeControl.ComboBox51.Text) - 1] + resistors[int.Parse(toeControl.ComboBox50.Text) - 1] + resistors[int.Parse(toeControl.ComboBox49.Text) - 1];
            if (toeControl.Label107.Text == "-") mainMatrix[1, boof] *= -1;

            switch (toeControl.ComboBox47.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[1, boof] = resistors[int.Parse(toeControl.ComboBox81.Text) - 1];
            if (toeControl.Label96.Text == "-") mainMatrix[1, boof] *= -1;

            switch (toeControl.ComboBox44.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[1, boof] = resistors[int.Parse(toeControl.ComboBox84.Text) - 1];
            if (toeControl.Label93.Text == "-") mainMatrix[1, boof] *= -1;

            if (toeControl.ComboBox46.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox46.Text);
                if (toeControl.Label100.Text == "+")
                    mainMatrix[1, 3] += EMF[boof - 1];
                else mainMatrix[1, 3] -= EMF[boof - 1];
            }

            if (toeControl.ComboBox48.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox48.Text);
                if (toeControl.Label98.Text == "+")
                    mainMatrix[1, 3] += EMF[boof - 1];
                else mainMatrix[1, 3] -= EMF[boof - 1];
            }

            ////////////////////////////////////////////////////////

            switch (toeControl.ComboBox62.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[2, boof] = resistors[int.Parse(toeControl.ComboBox61.Text) - 1] + resistors[int.Parse(toeControl.ComboBox59.Text) - 1] + resistors[int.Parse(toeControl.ComboBox58.Text) - 1];
            if (toeControl.Label127.Text == "-") mainMatrix[2, boof] *= -1;

            switch (toeControl.ComboBox55.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[2, boof] = resistors[int.Parse(toeControl.ComboBox82.Text) - 1];
            if (toeControl.Label113.Text == "-") mainMatrix[2, boof] *= -1;

            switch (toeControl.ComboBox53.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
            }

            mainMatrix[2, boof] = resistors[int.Parse(toeControl.ComboBox83.Text) - 1];
            if (toeControl.Label110.Text == "-") mainMatrix[2, boof] *= -1;

            if (toeControl.ComboBox54.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox54.Text);
                if (toeControl.Label120.Text == "+")
                    mainMatrix[2, 3] += EMF[boof - 1];
                else mainMatrix[2, 3] -= EMF[boof - 1];
            }

            if (toeControl.ComboBox56.Text != "Выберите")
            {
                boof = int.Parse(toeControl.ComboBox56.Text);
                if (toeControl.Label115.Text == "+")
                    mainMatrix[2, 3] += EMF[boof - 1];
                else mainMatrix[2, 3] -= EMF[boof - 1];
            }

            var matrix = Matrix<double>.Build.DenseOfArray(new double[3, 3]);

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix[i, j] = mainMatrix[i, j];
                }
            }

            boof = 0;
            var matrix1 = Matrix<double>.Build.DenseOfArray(new double[3, 3]);

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix1[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 3; i++)
            {
                matrix1[i, boof] = mainMatrix[i, 3];
            }
            boof++;

            var matrix2 = Matrix<double>.Build.DenseOfArray(new double[3, 3]);

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix2[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 3; i++)
            {
                matrix2[i, boof] = mainMatrix[i, 3];
            }
            boof++;

            var matrix3 = Matrix<double>.Build.DenseOfArray(new double[3, 3]);

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix3[i, j] = mainMatrix[i, j];
                }
            }
            for (int i = 0; i < 3; i++)
            {
                matrix3[i, boof] = mainMatrix[i, 3];
            }
            boof++;

            /////////////////////
            StringBuilder sb = new StringBuilder();

            // Функция для форматирования матрицы MathNet
            string FormatMatrix(Matrix<double> mat)
            {
                StringBuilder matrixSb = new StringBuilder();
                int rows = mat.RowCount;
                int cols = mat.ColumnCount;

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        // Форматируем числа для красивого вывода
                        matrixSb.Append(mat[i, j].ToString("F2").PadLeft(8));
                    }
                    matrixSb.AppendLine();
                }
                return matrixSb.ToString();
            }

            // Добавляем все матрицы
            sb.AppendLine("mainMatrix[3, 4]:");
            sb.AppendLine(FormatMatrix(mainMatrix));

            sb.AppendLine("matrix[3, 3]:");
            sb.AppendLine(FormatMatrix(matrix));

            sb.AppendLine("matrix1[3, 3]:");
            sb.AppendLine(FormatMatrix(matrix1));

            sb.AppendLine("matrix2[3, 3]:");
            sb.AppendLine(FormatMatrix(matrix2));

            sb.AppendLine("matrix3[3, 3]:");
            sb.AppendLine(FormatMatrix(matrix3));

            // Показываем MessageBox
            MessageBox.Show(sb.ToString(), "Матрицы MathNet", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //////////////////////

            double det = matrix.Determinant();
            double det1 = matrix1.Determinant();
            double det2 = matrix2.Determinant();
            double det3 = matrix3.Determinant();

            double[] Ik = new double[3];

            Ik[0] = det1 / det;
            Ik[1] = det2 / det;
            Ik[2] = det3 / det;
            //////////////////////
            sb = new StringBuilder();

            sb.AppendLine("=== РЕЗУЛЬТАТЫ ВЫЧИСЛЕНИЙ ===");
            sb.AppendLine();

            sb.AppendLine("1. ДЕТЕРМИНАНТЫ МАТРИЦ:");
            sb.AppendLine($"   det(matrix)  = {det:F6}");
            sb.AppendLine($"   det(matrix1) = {det1:F6}");
            sb.AppendLine($"   det(matrix2) = {det2:F6}");
            sb.AppendLine($"   det(matrix3) = {det3:F6}");
            sb.AppendLine();

            sb.AppendLine("2. ВЫЧИСЛЕНИЕ IK:");
            sb.AppendLine($"   Ik[0] = det1 / det = {det1:F6} / {det:F6} = {Ik[0]:F6}");
            sb.AppendLine($"   Ik[1] = det2 / det = {det2:F6} / {det:F6} = {Ik[1]:F6}");
            sb.AppendLine($"   Ik[2] = det3 / det = {det3:F6} / {det:F6} = {Ik[2]:F6}");
            sb.AppendLine();

            sb.AppendLine("3. ИТОГОВЫЕ ЗНАЧЕНИЯ IK:");
            for (int i = 0; i < Ik.Length; i++)
            {
                sb.AppendLine($"   Ik[{i}] = {Ik[i]:F8}");
            }
            sb.AppendLine();

            // Проверка на деление на ноль
            // Показываем MessageBox
            MessageBox.Show(sb.ToString(), "Результаты вычислений",
                           MessageBoxButtons.OK, MessageBoxIcon.Information);
            ////////////////////////


            /////////////////////////
            double I1 = 0;
            double I2 = 0;
            double I3 = 0;
            double I4 = 0;
            double I5 = 0;
            double I6 = 0;

            switch (toeControl.ComboBox63.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }
            if (boof != -1)
            {
                if (toeControl.Label232.Text == "+")
                    I1 += Ik[boof];
                else I1 -= Ik[boof];
            }

            switch (toeControl.ComboBox64.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label234.Text == "+")
                    I1 += Ik[boof];
                else I1 -= Ik[boof];
            }

            switch (toeControl.ComboBox66.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }
            if (boof != -1)
            {
                if (toeControl.Label238.Text == "+")
                    I2 += Ik[boof];
                else I2 -= Ik[boof];
            }

            switch (toeControl.ComboBox65.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label236.Text == "+")
                    I2 += Ik[boof];
                else I2 -= Ik[boof];
            }
            /////
            switch (toeControl.ComboBox73.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label245.Text == "+")
                    I3 += Ik[boof];
                else I3 -= Ik[boof];
            }

            switch (toeControl.ComboBox72.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label243.Text == "+")
                    I3 += Ik[boof];
                else I3 -= Ik[boof];
            }
            //////////
            switch (toeControl.ComboBox75.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label252.Text == "+")
                    I4 += Ik[boof];
                else I4 -= Ik[boof];
            }

            switch (toeControl.ComboBox74.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label250.Text == "+")
                    I4 += Ik[boof];
                else I4 -= Ik[boof];
            }
            ////////////////////////
            switch (toeControl.ComboBox77.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label259.Text == "+")
                    I5 += Ik[boof];
                else I5 -= Ik[boof];
            }

            switch (toeControl.ComboBox76.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label257.Text == "+")
                    I5 += Ik[boof];
                else I5 -= Ik[boof];
            }

            ////////////////////////
            switch (toeControl.ComboBox79.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label266.Text == "+")
                    I6 += Ik[boof];
                else I6 -= Ik[boof];
            }

            switch (toeControl.ComboBox78.Text)
            {
                case "I":
                    boof = 0;
                    break;
                case "II":
                    boof = 1;
                    break;
                case "III":
                    boof = 2;
                    break;
                default:
                    boof = -1;
                    break;
            }

            if (boof != -1)
            {
                if (toeControl.Label264.Text == "+")
                    I6 += Ik[boof];
                else I6 -= Ik[boof];
            }

            toeControl.TextBox34.Text = I1.ToString();
            toeControl.TextBox35.Text = I2.ToString();
            toeControl.TextBox37.Text = I3.ToString();
            toeControl.TextBox43.Text = I4.ToString();
            toeControl.TextBox41.Text = I5.ToString();
            toeControl.TextBox39.Text = I6.ToString();
        }
    }
}
