using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FilmCalc
{
    public partial class masterCase : Form
    {
        public masterCase()
        {
            InitializeComponent();
            boxPlay.Text = "1/4";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool fillError = errorCheck();
            if (fillError == true)
            {
                MessageBox.Show("Please fill out the missing information.");
            }
            else
            {
                bool isNumeric = numericCheck();
                if (isNumeric == true)
                {
                    double play = fractionSplit(boxPlay.Text);
                    double width = (double.Parse(boxWidthTotal.Text) * ((double.Parse(boxWidthIn.Text)) + (fractionSplit(boxWidthFrac.Text)))) + play;
                    double length = (double.Parse(boxLengthTotal.Text) * ((double.Parse(boxLengthIn.Text)) + (fractionSplit(boxLengthFrac.Text)))) + play;
                    double height = (double.Parse(boxHeightTotal.Text) * ((double.Parse(boxHeightIn.Text)) + (fractionSplit(boxHeightFrac.Text)))) + play;
                    double boxVolume = width * length * height;
                    masterWidth.Text = width.ToString();
                    masterLength.Text = length.ToString();
                    masterHeight.Text = height.ToString();
                    masterVolume.Text = boxVolume.ToString();
                }
            }
        }

        public double fractionSplit(string strFraction)
        {
            string[] fraction = strFraction.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string Fraction1 = fraction[0];
            string Fraction2 = string.Join(" ", fraction.Skip(1).ToList());
            double total = 0;
            int temp;
            bool test = int.TryParse(Fraction2, out temp);
            if (test == false)
            {
                total = double.Parse(Fraction1);
            }
            else
            {
                total = double.Parse(Fraction1) / double.Parse(Fraction2);
            }
            return total;
        }

        /*private void FilmCalc_Load(object sender, EventArgs e)
        {
            boxPlay.Text = "1/4";
        }*/

        public bool errorCheck()
        {
            //checks for empty or improper variables
            ast1.Visible = false;
            ast2.Visible = false;
            ast3.Visible = false;
            ast4.Visible = false;
            ast5.Visible = false;
            ast6.Visible = false;
            ast7.Visible = false;
            ast8.Visible = false;
            ast9.Visible = false;
            ast10.Visible = false;
            bool fillError = false;
            if (String.IsNullOrEmpty(boxWidthIn.Text))
            {
                fillError = true;
                ast1.Visible = true;
            }
            if (String.IsNullOrEmpty(boxLengthIn.Text))
            {
                fillError = true;
                ast2.Visible = true;
            }
            if (String.IsNullOrEmpty(boxHeightIn.Text))
            {
                fillError = true;
                ast3.Visible = true;
            }
            if (String.IsNullOrEmpty(boxPlay.Text))
            {
                fillError = true;
                ast4.Visible = true;
            }
            if (String.IsNullOrEmpty(boxWidthFrac.Text))
            {
                fillError = true;
                ast5.Visible = true;
            }
            if (String.IsNullOrEmpty(boxLengthFrac.Text))
            {
                fillError = true;
                ast6.Visible = true;
            }
            if (String.IsNullOrEmpty(boxHeightFrac.Text))
            {
                fillError = true;
                ast7.Visible = true;
            }
            if (String.IsNullOrEmpty(boxWidthTotal.Text))
            {
                fillError = true;
                ast8.Visible = true;
            }
            if (String.IsNullOrEmpty(boxLengthTotal.Text))
            {
                fillError = true;
                ast9.Visible = true;
            }
            if (String.IsNullOrEmpty(boxHeightTotal.Text))
            {
                fillError = true;
                ast10.Visible = true;
            }
            return fillError;
        }

        public bool numericCheck()
        {
            //checks to make sure that the text boxes have no letters or symbols in them.
            bool numericStatus = true;
            double number;
            if ((double.TryParse(boxWidthIn.Text, out number)) == false)
            {
                numericStatus = false;
                ast1.Visible = true;
            }
            if ((double.TryParse(boxLengthIn.Text, out number)) == false)
            {
                numericStatus = false;
                ast2.Visible = true;
            }
            if ((double.TryParse(boxHeightIn.Text, out number)) == false)
            {
                numericStatus = false;
                ast3.Visible = true;
            }
            if ((double.TryParse(boxWidthTotal.Text, out number)) == false)
            {
                numericStatus = false;
                ast8.Visible = true;
            }
            if ((double.TryParse(boxLengthTotal.Text, out number)) == false)
            {
                numericStatus = false;
                ast9.Visible = true;
            }
            if ((double.TryParse(boxHeightTotal.Text, out number)) == false)
            {
                numericStatus = false;
                ast10.Visible = true;
            }
            if (numericStatus == false)
            {
                MessageBox.Show("Cannot parse the information, please verify that there are no letters or symbols in the textboxes.");
            }
            return numericStatus;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            boxWidthIn.Text = "";
            boxWidthFrac.Text = "";
            boxWidthTotal.Text = "";
            boxLengthIn.Text = "";
            boxLengthFrac.Text = "";
            boxLengthTotal.Text = "";
            boxHeightIn.Text = "";
            boxHeightFrac.Text = "";
            boxHeightTotal.Text = "";
            masterHeight.Text = "";
            masterLength.Text = "";
            masterVolume.Text = "";
            masterWidth.Text = "";
            boxPlay.Text = "1/4";
            ast1.Visible = false;
            ast2.Visible = false;
            ast3.Visible = false;
            ast4.Visible = false;
            ast5.Visible = false;
            ast6.Visible = false;
            ast7.Visible = false;
            ast8.Visible = false;
            ast9.Visible = false;
            ast10.Visible = false;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void masterCase_Load(object sender, EventArgs e)
        {

        }
    }
}
