using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MathNet.Numerics;

namespace Chamomile
{
    public class TOE1Calculation
    {
        public static void Calculate(TOE1Control toeControl)
        {
            double R1, R2, R3, R4, R5, R6, E1, E2, E3;

            R1 = double.Parse(toeControl.TextBox4.Text);
            R2 = double.Parse(toeControl.TextBox5.Text);
            R3 = double.Parse(toeControl.TextBox7.Text);
            R4 = double.Parse(toeControl.TextBox6.Text);
            R5 = double.Parse(toeControl.TextBox9.Text);
            R6 = double.Parse(toeControl.TextBox8.Text);

            if (!string.IsNullOrWhiteSpace(toeControl.TextBox1.Text))
            {
                E1 = double.Parse(toeControl.TextBox1.Text);
            }

            if (!string.IsNullOrWhiteSpace(toeControl.TextBox2.Text))
            {
                E2 = double.Parse(toeControl.TextBox2.Text);
            }

            if (!string.IsNullOrWhiteSpace(toeControl.TextBox3.Text))
            {
                E3 = double.Parse(toeControl.TextBox3.Text);
            }

            Method1(toeControl);

        }

        public static void Method1(TOE1Control toeControl)
        {
            int boof;
            double[,] matrix = new double[6, 7];

            boof = int.Parse(toeControl.ComboBox4.Text);
            matrix[0, boof] = 1;
            if (toeControl.Label19.Text == "-")
                matrix[0, boof] *= -1;

            boof = int.Parse(toeControl.ComboBox5.Text);
            matrix[0, boof] = 1;
            if (toeControl.Label16.Text == "-")
                matrix[0, boof] *= -1;

            boof = int.Parse(toeControl.ComboBox6.Text);
            matrix[0, boof] = 1;
            if (toeControl.Label20.Text == "-")
                matrix[0, boof] *= -1;

            matrix[0, 6] = 0;

            boof = int.Parse(toeControl.ComboBox9.Text);
            matrix[1, boof] = 1;
            if (toeControl.Label26.Text == "-")
                matrix[1, boof] *= -1;

            boof = int.Parse(toeControl.ComboBox8.Text);
            matrix[1, boof] = 1;
            if (toeControl.Label24.Text == "-")
                matrix[1, boof] *= -1;

            boof = int.Parse(toeControl.ComboBox7.Text);
            matrix[1, boof] = 1;
            if (toeControl.Label22.Text == "-")
                matrix[1, boof] *= -1;

            matrix[1, 6] = 0;

            boof = int.Parse(toeControl.ComboBox12.Text);
            matrix[2, boof] = 1;
            if (toeControl.Label33.Text == "-")
                matrix[2, boof] *= -1;

            boof = int.Parse(toeControl.ComboBox11.Text);
            matrix[2, boof] = 1;
            if (toeControl.Label31.Text == "-")
                matrix[2, boof] *= -1;

            boof = int.Parse(toeControl.ComboBox10.Text);
            matrix[2, boof] = 1;
            if (toeControl.Label29.Text == "-")
                matrix[2, boof] *= -1;

            matrix[2, 6] = 0;

        }
    }
}
