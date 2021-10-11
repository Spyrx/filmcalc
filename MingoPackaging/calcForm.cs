using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using FilmCalc;
using System.Configuration;

namespace MingoPackaging
{
    public partial class calcForm : Form
    {
        BindingList<string> toolList = new BindingList<string>();
        //BindingSource bs = new BindingSource();
        //List<int> indexList = new List<int>();
        string[,] dataCLDalt = new string[,] {
                {"GC092",   "4.7500",   "4.5000",   "0.6250",   "0.6250"},
                {"GC050",   "5.0000",   "3.5000",   "0.5000",   "0.5000"},
                {"GC100",   "5.0000",   "5.0000",   "0.6250",   "0.6250"},
                {"GC008",   "5.0000",   "5.5420",   "0.6250",   "0.6250"},
                {"GC055",   "5.0000",   "6.0000",   "0.6250",   "0.6250"},
                {"GC007",   "5.0000",   "6.8750",   "0.6250",   "0.6250"},
                {"GC037",   "5.0400",   "2.8250",   "0.5000",   "0.5000"},
                {"GC084",   "5.0400",   "2.9583",   "0.5000",   "0.5000"},
                {"GC105",   "5.1250",   "3.2500",   "0.5000",   "0.5000"},
                {"GC091",   "5.1250",   "5.7500",   "0.6250",   "0.6250"},
                {"GC107",   "5.1250",   "6.0000",   "0.6250",   "0.6250"},
                {"GC098",   "5.1600",   "3.5000",   "0.6250",   "0.6250"},
                {"GC095",   "5.1880",   "3.2500",   "0.5000",   "0.5000"},
                {"GC087",   "5.2500",   "3.5000",   "0.6250",   "0.6250"},
                {"GC005",   "5.2500",   "5.6670",   "0.6250",   "0.6250"},
                {"GC015",   "5.2500",   "6.0000",   "0.6250",   "0.6250"},
                {"GC073",   "5.2500",   "6.3125",   "0.6250",   "0.6250"},
                {"GC061",   "5.2500",   "6.6250",   "0.6250",   "0.6250"},
                {"GC041",   "5.3150",   "5.2500",   "0.6250",   "0.6250"},
                {"GC074",   "5.5000",   "3.7500",   "0.6250",   "0.6250"},
                {"GC063",   "5.5000",   "4.5625",   "0.5000",   "0.6250"},
                {"GC088",   "5.5000",   "5.7500",   "0.6250",   "0.6250"},
                {"GC042",   "5.5000",   "6.0000",   "0.6250",   "0.6250"},
                {"GC060-5", "5.5000",   "6.1667",   "0.5000",   "0.6250"},
                {"GC082",   "5.5000",   "6.2500",   "0.5000",   "0.5000"},
                {"GC064",   "5.5000",   "6.3125",   "0.5000",   "0.6250"},
                {"GC004",   "5.5000",   "6.5625",   "0.6250",   "0.6250"},
                {"GC031",   "5.6000",   "5.3750",   "0.6250",   "0.6250"},
                {"GC097",   "5.6250",   "5.2500",   "0.6250",   "0.6250"},
                {"GC043",   "5.6875",   "5.7500",   "0.6250",   "0.6250"},
                {"GC011",   "5.6875",   "5.9166",   "0.6250",   "0.6250"},
                {"GC016",   "5.7500",   "6.0000",   "0.6250",   "0.6250"},
                {"GC006",   "5.7500",   "6.1670",   "0.6250",   "0.6250"},
                {"GC046",   "5.7500",   "7.3750",   "0.6250",   "0.6250"},
                {"GC009",   "5.7500",   "8.0625",   "0.6250",   "0.6250"},
                {"GC102",   "5.8200",   "6.8750",   "0.6250",   "0.6250"},
                {"GC096",   "5.8260",   "7.7500",   "0.5000",   "0.5000"},
                {"GC027",   "6.0000",   "3.5000",   "0.6250",   "0.6250"},
                {"GC038",   "6.0000",   "3.7813",   "0.6250",   "0.6250"},
                {"GC085",   "6.0000",   "4.0000",   "0.6250",   "0.6250"},
                {"GC017",   "6.0000",   "6.0000",   "0.6250",   "0.6250"},
                {"GC049",   "6.0000",   "6.2500",   "0.6250",   "0.6250"},
                {"GC056",   "6.0000",   "6.2500",   "0.5000",   "0.6250"},
                {"GC051",   "6.0000",   "6.5625",   "0.6250",   "0.6250"},
                {"GC070",   "6.0000",   "7.0000",   "0.6250",   "0.6250"},
                {"GC001",   "6.0000",   "7.6250",   "0.6250",   "0.6250"},
                {"GC044",   "6.1250",   "5.5833",   "0.6250",   "0.6250"},
                {"GC036",   "6.2000",   "6.5625",   "0.6250",   "0.6250"},
                {"GC012",   "6.2500",   "5.7500",   "0.6250",   "0.6250"},
                {"GC090",   "6.2500",   "6.2500",   "0.5000",   "0.6250"},
                {"GC086",   "6.2500",   "6.6250",   "0.5000",   "0.6250"},
                {"GC081",   "6.5000",   "4.5000",   "0.6250",   "0.6250"},
                {"GC093",   "6.5000",   "5.0000",   "0.6250",   "0.6250"},
                {"GC108",   "6.5000",   "5.5000",   "0.6250",   "0.6250"},
                {"GC094",   "6.5000",   "5.7500",   "0.6250",   "0.6250"},
                {"GC026",   "6.5630",   "6.4375",   "0.6250",   "0.6250"},
                {"GC099",   "6.6900",   "5.6667",   "0.5950",   "0.5000"},
                {"GC089",   "6.7500",   "7.0000",   "0.6250",   "0.6250"},
                {"GC054",   "7.0000",   "4.7500",   "0.6250",   "0.6250"},
                {"GC067",   "7.0000",   "5.2500",   "0.6250",   "0.6250"},
                {"GC083",   "7.0000",   "5.6250",   "0.6250",   "0.6250"},
                {"GC106",   "7.0000",   "6.0000",   "0.6250",   "0.6250"},
                {"GC101",   "7.5000",   "4.7500",   "0.6250",   "0.6250"},
                {"GC069",   "7.5000",   "5.5000",   "0.5000",   "0.6250"},
                {"GC025",   "7.7500",   "6.5625",   "0.6250",   "0.6250"},
                {"GC104",   "8.0000",   "7.5000",   "0.5000",   "0.6250"},
                {"GC057",   "8.2500",   "4.7500",   "0.6250",   "0.6250"},
                {"GC077",   "8.2500",   "5.5000",   "0.6250",   "0.6250"},
            };

        public calcForm()
        {
            InitializeComponent();
            txtHeight.Text = "0.5";
            txtFinseal.Text = "0.625";
            txtCrimp.Text = txtFinseal.Text;

            
            //comboBox1.FormattingEnabled = true;
            //bs.DataSource = toolList;
            listToolNumber.DataSource = toolList;
            //comboBox1.DataSource = toolList;

            //string selectedTool = listToolNumber.GetItemText(listToolNumber.SelectedValue);
            //int index = Array.FindIndex<string>(dataCLDalt, selectedTool);
        }

        private void txtWidth_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnCalc.PerformClick();
                MessageBox.Show("Enter");
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {
            //accidental click, cant delete
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //exit button
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //reset button, sets back to default
            txtHeight.Text = "0.5";
            txtFinseal.Text = "0.625";
            txtCrimp.Text = txtFinseal.Text;
            txtWidth.Text = "";
            txtLength.Text = "";
            //txtWeb.Text = "";
            //txtRepeat.Text = "";
            //txtRelm.Text = "";
            //txtCLD.Text = "";
            exportRepeat.Clear();
            exportWeb.Clear();
            exportFinseal.Clear();
            exportEndseal.Clear();
            //comboBox1.ResetText();
            //toolList.Clear();
            //toolList.Add("Select Item:");
            //comboBox1.DataSource = null;
            //comboBox1.SelectedItem = "";
            //comboBox1.ResetText();
            //comboBox1.Text = "Select Item: ";
            //comboBox1.SelectedValue = 0;
            //comboBox1.SelectedIndex = -1;
            //comboBox1.Refresh();
            //listToolNumber.Refresh();
            //comboBox1.DataSource = toolList;
            //comboBox1.SelectedItem = "";
            //comboBox1.Refresh();
            ast1.Visible = false;
            ast2.Visible = false;
            ast3.Visible = false;
            ast4.Visible = false;
            ast6.Visible = false;
            //btnExport.Visible = false;
        }

        public bool errorCheck()
        {
            //checks for empty or improper variables
            ast1.Visible = false;
            ast2.Visible = false;
            ast3.Visible = false;
            ast4.Visible = false;
            ast6.Visible = false;
            bool fillError = false;
            if (String.IsNullOrEmpty(txtWidth.Text))
            {
                fillError = true;
                ast1.Visible = true;
            }
            if (String.IsNullOrEmpty(txtHeight.Text))
            {
                fillError = true;
                ast2.Visible = true;
            }
            if (String.IsNullOrEmpty(txtFinseal.Text))
            {
                fillError = true;
                ast3.Visible = true;
            }
            if (String.IsNullOrEmpty(txtLength.Text))
            {
                fillError = true;
                ast4.Visible = true;
            }
            if (String.IsNullOrEmpty(txtCrimp.Text))
            {
                fillError = true;
                ast6.Visible = true;
            }
            return fillError;
        }

        public bool numericCheck()
        {
            //checks to make sure that the text boxes have no letters or symbols in them.
            bool numericStatus = true;
            double number;
            if ((double.TryParse(txtWidth.Text, out number)) == false)
            {
                numericStatus = false;
                ast1.Visible = true;
            }
            if ((double.TryParse(txtHeight.Text, out number)) == false)
            {
                numericStatus = false;
                ast2.Visible = true;
            }
            if ((double.TryParse(txtFinseal.Text, out number)) == false)
            {
                numericStatus = false;
                ast3.Visible = true;
            }
            if ((double.TryParse(txtLength.Text, out number)) == false)
            {
                numericStatus = false;
                ast4.Visible = true;
            }
            if ((double.TryParse(txtCrimp.Text, out number)) == false)
            {
                numericStatus = false;
                ast6.Visible = true;
            }
            if (numericStatus == false)
            {
                MessageBox.Show("Cannot parse the information, please verify that there are no letters or symbols in the textboxes.");
            }
            return numericStatus;
        }

        /*public string getRelm(double webActual, double repeatActual)
        {
            string[,] dataRelm = new string[,]
            {
                {"GCS0016", "4",    "2.875"},
                {"GCS0020", "4.75", "3.265625"},
                {"GCS0005", "4.75", "3.75"},
                {"GCS0021", "4.875",    "4"},
                {"GCS0007", "5.25", "5.2795"},
                {"GCS0019", "5.25", "6.25"},
                {"GCS0008", "5.375",    "6.125"},
                {"GCS0013", "5.5",  "6"},
                {"GCS0017", "5.5",  "6"},
                {"GCS0018", "5.875",    "6.125"},
                {"GCS0015", "6",    "5.375"},
                {"GCS0010", "6",    "5.5"},
                {"GCS0014", "6",    "6"},
                {"GCS0002", "6.375",    "3.625"},
                {"GCS0004", "6.375",    "5.75"},
                {"GCS0009", "6.9375",   "4.75"},
                {"GCS0001", "6.9375",   "5.5"},
                {"GCS0012", "7",    "5"},
                {"GCS0006", "7.432",    "8.25"},
                {"GCS0011", "7.4375",   "4.375"}
            };

            string toolNumber = "CUSTOM";

            int bound0 = dataRelm.GetUpperBound(0);
            int bound1 = dataRelm.GetUpperBound(1);
            // finds upper bounds for data array

            double xPlay = 0.125;
            double yPlay = 0.125;

            int[,] isMatch = new int[(bound0 + 1), bound1];
            // initializes found match array

            int bound2 = isMatch.GetUpperBound(0);
            int bound3 = isMatch.GetUpperBound(1);
            // finds upper bounds for match array

            for (int i = 0; i <= bound2; i++)
            {
                for (int j = 0; j <= bound3; j++)
                {
                    isMatch[i, j] = 0;
                }
            }
            // fills found match array

            for (int i = 0; i <= bound0; i++)
            {
                if ((double.Parse(dataRelm[i, 1]) >= (webActual - xPlay)) && (double.Parse(dataRelm[i, 1]) <= (webActual + xPlay)))
                {
                    isMatch[i, 0] = 1;
                }
                if ((double.Parse(dataRelm[i, 2]) >= (repeatActual - yPlay)) && (double.Parse(dataRelm[i, 2]) <= (repeatActual + yPlay)))
                {
                    isMatch[i, 1] = 1;
                }
            }
            // finds matches

            double difference = 5555555;
            double location = -1;
            double width;
            double repeat;
            double absval;
            
            for (int i = 0; i <= bound2; i++)
            {
                if (isMatch[i, 0] == 1 && isMatch[i, 1] == 1)
                {
                    width = double.Parse(dataRelm[i, 1]);
                    repeat = double.Parse(dataRelm[i, 2]);
                    absval = Math.Abs(width - webActual);
                    absval += Math.Abs(repeat - repeatActual);
                    if (absval < difference)
                    {
                        difference = absval;
                        location = i;
                    }
                }
            }
            if (location != -1)
            {
                toolNumber = dataRelm[(int)location, 0];
            }

            return toolNumber;
        }*/

        /*public string getCLD(double webActual, double repeatActual)
        {
            string[,] dataCLD = new string[,]
            {
                {"GC035",   "4.1730",   "5.7500",   "0.625",    "0.625"},
                {"GC021",   "4.6250",   "4.3750",   "0.625",    "0.625"},
                {"GC022",   "5.0000",   "4.5313",   "0.625",    "0.625"},
                {"GC008",   "5.0000",   "5.5420",   "0.625",    "0.625"},
                {"GC055",   "5.0000",   "6.0000",   "0.625",    "0.625"},
                {"GC005",   "5.2500",   "5.6670",   "0.625",    "0.625"},
                {"GC015",   "5.2500",   "6.0000",   "0.625",    "0.625"},
                {"GC073",   "5.2500",   "6.3125",   "0.625",    "0.625"},
                {"GC061",   "5.2500",   "6.6250",   "0.625",    "0.625"},
                {"GC041",   "5.3150",   "5.2500",   "0.625",    "0.625"},
                {"GC053",   "5.4720",   "6.9375",   "0.625",    "0.625"},
                {"GC013",   "5.4720",   "7.1875",   "0.625",    "0.625"},
                {"GC074",   "5.5000",   "3.7500",   "0.625",    "0.625"},
                {"GC042",   "5.5000",   "6.0000",   "0.625",    "0.625"},
                {"GC004",   "5.5000",   "6.5625",   "0.625",    "0.625"},
                {"GC076",   "5.5000",   "7.7500",   "0.625",    "0.625"},
                {"GC029",   "5.6000",   "4.0000",   "0.625",    "0.625"},
                {"GC031",   "5.6000",   "5.3750",   "0.625",    "0.625"},
                {"GC032",   "5.6500",   "7.0000",   "0.625",    "0.625"},
                {"GC043",   "5.6875",   "5.7500",   "0.625",    "0.625"},
                {"GC011",   "5.6875",   "5.9166",   "0.625",    "0.625"},
                {"GC003",   "5.6875",   "8.2500",   "0.625",    "0.625"},
                {"GC045",   "5.6875",   "4.0000",   "0.625",    "0.625"},
                {"GC016",   "5.6875",   "6.0000",   "0.625",    "0.625"},
                {"GC006",   "5.6875",   "6.1670",   "0.625",    "0.625"},
                {"GC062",   "5.6875",   "6.6250",   "0.625",    "0.625"},
                {"GC046",   "5.6875",   "7.3750",   "0.625",    "0.625"},
                {"GC009",   "5.6875",   "8.0625",   "0.625",    "0.625"},
                {"GC027",   "5.6875",   "3.5000",   "0.625",    "0.625"},
                {"GC038",   "6.0000",   "3.7813",   "0.625",    "0.625"},
                {"GC052",   "6.0000",   "4.7500",   "0.625",    "0.625"},
                {"GC017",   "6.0000",   "6.0000",   "0.625",    "0.625"},
                {"GC049",   "6.0000",   "6.2500",   "0.625",    "0.625"},
                {"GC051",   "6.0000",   "6.5625",   "0.625",    "0.625"},
                {"GC007",   "6.0000",   "6.8750",   "0.625",    "0.625"},
                {"GC070",   "6.0000",   "7.0000",   "0.625",    "0.625"},
                {"GC001",   "6.0000",   "7.6250",   "0.625",    "0.625"},
                {"GC034",   "6.0000",   "8.0625",   "0.625",    "0.625"},
                {"GC020",   "6.0000",   "8.5625",   "0.625",    "0.625"},
                {"GC044",   "6.1250",   "5.5833",   "0.625",    "0.625"},
                {"GC058",   "6.1420",   "5.9167",   "0.625",    "0.625"},
                {"GC036",   "6.2000",   "6.5625",   "0.625",    "0.625"},
                {"GC030",   "6.2500",   "4.5000",   "0.625",    "0.625"},
                {"GC012",   "6.2500",   "5.7500",   "0.625",    "0.625"},
                {"GC048",   "6.3750",   "3.8750",   "0.625",    "0.625"},
                {"GC028",   "6.5000",   "7.3125",   "0.625",    "0.625"},
                {"GC026",   "6.5630",   "6.4375",   "0.625",    "0.625"},
                {"GC014",   "6.6250",   "6.0000",   "0.625",    "0.625"},
                {"GC023",   "6.6250",   "8.2500",   "0.625",    "0.625"},
                {"GC059",   "6.7500",   "6.3125",   "0.625",    "0.625"},
                {"GC075",   "6.7500",   "7.7500",   "0.625",    "0.625"},
                {"GC054",   "7.0000",   "4.7500",   "0.625",    "0.625"},
                {"GC067",   "7.0000",   "5.2500",   "0.625",    "0.625"},
                {"GC033",   "7.0250",   "7.3750",   "0.625",    "0.625"},
                {"GC025",   "7.7500",   "6.5625",   "0.625",    "0.625"},
                {"GC039",   "8.0000",   "4.2500",   "0.625",    "0.625"},
                {"GC040",   "8.0000",   "4.5000",   "0.625",    "0.625"},
                {"GC002",   "8.0000",   "7.1250",   "0.625",    "0.625"},
                {"GC057",   "8.2500",   "4.7500",   "0.625",    "0.625"},
                {"GC068",   "9.8750",   "7.6875",   "0.625",    "0.625"},
                {"GC047",   "10.5000",  "5.7500",   "0.625",    "0.625"}
            };

            string toolNumber = "CUSTOM";

            int bound0 = dataCLD.GetUpperBound(0);
            int bound1 = dataCLD.GetUpperBound(1);
            // finds upper bounds for data array

            double xPlay = 0.125;
            double yPlay = 0.125;

            int[,] isMatch = new int[(bound0 + 1), bound1];
            // initializes found match array

            int bound2 = isMatch.GetUpperBound(0);
            int bound3 = isMatch.GetUpperBound(1);
            // finds upper bounds for match array

            for (int i = 0; i <= bound2; i++)
            {
                for (int j = 0; j <= bound3; j++)
                {
                    isMatch[i, j] = 0;
                }
            }
            // fills found match array

            for (int i = 0; i <= bound0; i++)
            {
                if ((double.Parse(dataCLD[i, 1]) >= (webActual - xPlay)) && (double.Parse(dataCLD[i, 1]) <= (webActual + xPlay)))
                {
                    isMatch[i, 0] = 1;
                }
                if ((double.Parse(dataCLD[i, 2]) >= (repeatActual - yPlay)) && (double.Parse(dataCLD[i, 2]) <= (repeatActual + yPlay)))
                {
                    isMatch[i, 1] = 1;
                }
            }
            // finds matches

            double difference = 5555555;
            double location = -1;
            double width;
            double repeat;
            double absval;

            for (int i = 0; i <= bound2; i++)
            {
                if (isMatch[i, 0] == 1 && isMatch[i, 1] == 1)
                {
                    width = double.Parse(dataCLD[i, 1]);
                    repeat = double.Parse(dataCLD[i, 2]);
                    absval = Math.Abs(width - webActual);
                    absval += Math.Abs(repeat - repeatActual);
                    if (absval < difference)
                    {
                        difference = absval;
                        location = i;
                    }
                }
            }
            if (location != -1)
            {
                toolNumber = dataCLD[(int)location, 0];
            }

            return toolNumber;
        }*/

        public string getCLDalt(double webActual, double repeatActual, double finsealActual, double crimpActual)
        {
            string toolNumber = "CUSTOM";
            toolList.Clear();

            int bound = dataCLDalt.GetUpperBound(0);
            //int bound1 = dataCLDalt.GetUpperBound(1);
            // finds upper bounds for data array

            double xPlay = 0.125;
            double yPlay = 0.125;

            int[,] isMatch = new int[(bound + 1), dataCLDalt.GetUpperBound(1)];
            // initializes found match array

            int bound2 = isMatch.GetUpperBound(0);
            int bound3 = isMatch.GetUpperBound(1);
            // finds upper bounds for match array

            for (int i = 0; i <= bound2; i++)
            {
                for (int j = 0; j <= bound3; j++)
                {
                    isMatch[i, j] = 0;
                }
            }
            // fills found match array

            for (int i = 0; i <= bound; i++)
            {
                if ((double.Parse(dataCLDalt[i, 1]) >= (webActual - xPlay)) && (double.Parse(dataCLDalt[i, 1]) <= (webActual + xPlay)))
                {
                    isMatch[i, 0] = 1;
                }
                if ((double.Parse(dataCLDalt[i, 2]) >= (repeatActual - yPlay)) && (double.Parse(dataCLDalt[i, 2]) <= (repeatActual + yPlay)))
                {
                    isMatch[i, 1] = 1;
                }
                if (double.Parse(dataCLDalt[i, 3]) == finsealActual)
                {
                    isMatch[i, 2] = 1;
                }
                if (double.Parse(dataCLDalt[i, 4]) == crimpActual)
                {
                    isMatch[i, 3] = 1;
                }
            }
            // finds matches

            double difference = 5555555;
            double location = -1;
            double width;
            double repeat;
            double absval;
            //listToolNumber.DataSource = null;

            for (int i = 0; i <= bound2; i++)
            {
                if (isMatch[i, 0] == 1 && isMatch[i, 1] == 1 && isMatch[i, 2] == 1 && isMatch[i, 3] == 1)
                {
                    toolList.Add(dataCLDalt[i, 0]);
                    //indexList.Add(i);
                    
                    //listToolNumber.Items.Add(new ListBoxItem(dataCLDalt(i, 0)), 0);
                    width = double.Parse(dataCLDalt[i, 1]);
                    repeat = double.Parse(dataCLDalt[i, 2]);
                    absval = (Math.Abs(width - webActual)) + (Math.Abs(repeat - repeatActual)); ;
                    if (absval < difference)
                    {
                        difference = absval;
                        location = i;
                    }
                }
            }

            if (location != -1)
            {
                toolNumber = dataCLDalt[(int)location, 0];
            }

            return toolNumber;
        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            //calculate button

            exportRepeat.Clear();
            exportWeb.Clear();
            exportFinseal.Clear();
            exportEndseal.Clear();
            //toolList.Clear();
            
            //comboBox1.SelectedItem = -1;
            
            //txtCLD.Text = "";
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
                    double widthActual = (double.Parse(txtWidth.Text))*2;
                    double HeightActual = (double.Parse(txtHeight.Text)) * 2;
                    double finsealActual = (double.Parse(txtFinseal.Text)) * 2;
                    double webActual = widthActual + HeightActual + finsealActual + 0.25;

                    double lengthActual = double.Parse(txtLength.Text);
                    double crimpActual = (double.Parse(txtCrimp.Text)) * 2;
                    double repeatActual = lengthActual + HeightActual + crimpActual + 0.25;

                    //txtWeb.Text = webActual.ToString();
                    //txtRepeat.Text = repeatActual.ToString();

                    getCLDalt(webActual, repeatActual, (finsealActual / 2), (crimpActual / 2));

                    btnExport.Enabled = true;
                    //comboBox1.Enabled = true;
                    listToolNumber.Enabled = true;
                }
                else
                {
                    //MessageBox.Show("There is a non-numeric character in one of the fields. Please resolve and try again.");
                }
            }

            toolList.Add("CUSTOM");
            //comboBox1.DataSource = null;
            //comboBox1.DataSource = toolList;
            //comboBox1.ResetText();
            //comboBox1.Items.Clear();
            //comboBox1.SelectedIndex = 0;
            listToolNumber.SelectedIndex = 0;
            updateTool();
            listToolNumber.Refresh();
            //comboBox1.Refresh();
            //listToolNumber.DataSource = toolList;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //start export
            /*var exportForm1 = new exportForm(txtCLD.Text, exportWeb.Text, exportRepeat.Text, exportFinseal.Text, txtHeight.Text, exportEndseal.Text);
            exportForm1.Show();
            this.Hide();*/
            double web = Double.Parse(exportWeb.Text);
            double finseal = Double.Parse(exportFinseal.Text);
            double height = Double.Parse(txtHeight.Text);
            double repeat = Double.Parse(exportRepeat.Text);
            double crimp = Double.Parse(exportEndseal.Text);
            //string CLDtoolnumber = "blah";
            string CLDtoolnumber = listToolNumber.SelectedItem.ToString();

            DateTime today = DateTime.Today;
            double front = (web - (finseal * 2) - (height * 2)) / 2;
            double back = front / 2;
            double ppiWeb;
            double ppiRepeat;
            string companyname = txtCompanyName.Text;
            string barname = txtBarName.Text;
            //string flavor = txtBarFlavor.Text;
            double sidemargin = 1;
            double topmargin = 0.75;
            if (companyname != "")
            {
                companyname = companyname + "_";
            }
            if (barname != "")
            {
                barname = barname + "_";
            }
            string filename = companyname + barname + DateTime.Now.ToString("yyyy-M-dd_HH-mm-ss") + ".pdf";
            string approver1 = ConfigurationManager.AppSettings.Get("Approver1");
            string approver2 = ConfigurationManager.AppSettings.Get("Approver2");
            double bleed = 0.125;

            double ppi = 72;
            double marginppi = 72;

            PdfDocument document = new PdfDocument();
            PdfPage page = document.AddPage();
            page.Size = PageSize.Letter;
            page.Orientation = PageOrientation.Landscape;
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XFont font = new XFont("Times New Roman", 8, XFontStyle.Regular);
            XFont font2 = new XFont("Times New Roman", 10, XFontStyle.Bold);
            XFont font3 = new XFont("Times New Roman", 9, XFontStyle.Italic);

            if ((web + (topmargin * 2)) > 8.5 || (repeat + (sidemargin * 2) + 3) > 11)
            {
                ppiWeb = ppi / ((web + (topmargin * 2)) / 7.75);
                ppiRepeat = ppi / ((repeat + (sidemargin * 2)) / 7.75);

                if (ppiWeb < ppiRepeat)
                {
                    ppi = ppiWeb;
                }
                else
                {
                    ppi = ppiRepeat;
                }
            }

            //************FILM LINES************
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), (topmargin * marginppi), ((repeat * ppi) + (sidemargin * marginppi)), (topmargin * marginppi));
            //line between top and finseal
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), ((topmargin * marginppi) + ((finseal) * ppi)), (((repeat) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal * ppi)));
            //line between finseal and back
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((finseal + back) * ppi)), (((repeat - crimp) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back) * ppi));
            //line between back and side
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((finseal + back + height) * ppi)), (((repeat - crimp) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + height) * ppi));
            //line between side and front
            gfx.DrawLine(XPens.Red, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((finseal + back + height + bleed) * ppi)), ((((repeat / 2) - 0.36) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + height + bleed) * ppi));
            //bleed line top left
            gfx.DrawLine(XPens.Red, ((sidemargin * marginppi) + ((repeat - crimp) * ppi)), ((topmargin * marginppi) + ((finseal + back + height + bleed) * ppi)), ((((repeat / 2) + 0.36) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + height + bleed) * ppi));
            //bleed line top right
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((finseal + back + height + front) * ppi)), (((repeat - crimp) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + height + front) * ppi));
            //line between front and side
            gfx.DrawLine(XPens.Red, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((finseal + back + height + front - bleed) * ppi)), ((((repeat / 2) - 0.36) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + height + front - bleed) * ppi));
            //bleed line bottom left
            gfx.DrawLine(XPens.Red, ((sidemargin * marginppi) + ((repeat - crimp) * ppi)), ((topmargin * marginppi) + ((finseal + back + height + front - bleed) * ppi)), ((((repeat / 2) + 0.36) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + height + front - bleed) * ppi));
            //bleed line bottom right
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((finseal + back + (height * 2) + front) * ppi)), (((repeat - crimp) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + back + (height * 2) + front) * ppi));
            //line between side and front
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + ((back + height) * 2) + front) * ppi)), (((repeat) * ppi) + (sidemargin * marginppi)), ((topmargin * marginppi) + (finseal + ((back + height) * 2) + front) * ppi));
            //line between finseal and side
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), ((topmargin * marginppi) + (web * ppi)), ((sidemargin * marginppi) + ((repeat) * ppi)), ((topmargin * marginppi) + (web * ppi)));
            //line between finseal and bottom
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), (topmargin * marginppi), (sidemargin * marginppi), ((topmargin * marginppi) + (web * ppi)));
            //left side line
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (repeat * ppi)), (topmargin * marginppi), ((sidemargin * marginppi) + (repeat * ppi)), ((topmargin * marginppi) + (web * ppi)));
            //right side line
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + (finseal * ppi)), ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin * marginppi) + ((web - finseal) * ppi)));
            //left crimp line
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + ((repeat - crimp) * ppi)), ((topmargin * marginppi) + (finseal * ppi)), ((sidemargin * marginppi) + ((repeat - crimp) * ppi)), ((topmargin * marginppi) + ((web - finseal) * ppi)));
            //right crimp line
            //***********************************


            //*********DIMENSION LINES***********
            gfx.DrawLine(XPens.Black, ((repeat * ppi) + ((sidemargin + 0.1) * marginppi)), (topmargin * marginppi), ((repeat * ppi) + ((sidemargin + 0.25) * marginppi)), (topmargin * marginppi));
            //vertical dimension top line
            gfx.DrawLine(XPens.Black, (((sidemargin + 0.175) * marginppi) + (repeat * ppi)), (topmargin * marginppi), (((sidemargin + 0.175) * marginppi) + (repeat * ppi)), (((topmargin - 0.175) * marginppi) + ((web / 2) * ppi)));
            //vertical dimension top center line
            gfx.DrawLine(XPens.Black, (((sidemargin + 0.175) * marginppi) + (repeat * ppi)), ((topmargin + 0.175) * marginppi) + ((web / 2) * ppi), (((sidemargin + 0.175) * marginppi) + (repeat * ppi)), ((topmargin * marginppi) + (web * ppi)));
            //vertical dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin + 0.1) * marginppi) + ((repeat) * ppi)), ((topmargin * marginppi) + (web * ppi)), (((sidemargin + 0.25) * marginppi) + ((repeat) * ppi)), ((topmargin * marginppi) + (web * ppi)));
            //vertical dimension bottom line

            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), (topmargin * marginppi), ((sidemargin - 0.1) * marginppi), (topmargin * marginppi));
            //finseal dimension top line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), (topmargin * marginppi), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + ((finseal / 2) * ppi)));
            //finseal dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + ((finseal / 2) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + (finseal * ppi)));
            //finseal dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + (finseal * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + (finseal * ppi)));
            //finseal dimension bottom line

            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + (finseal * ppi)), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + (((back / 2) + finseal) * ppi)));
            //back dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + (((back / 2) + finseal) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((back + finseal) * ppi)));
            //back dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + ((back + finseal) * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + ((back + finseal) * ppi)));
            //back dimension bottom line

            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((finseal + back) * ppi)), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + (((height / 2) + back + finseal) * ppi)));
            //side dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + (((height / 2) + back + finseal) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((height + back + finseal) * ppi)));
            //side dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + ((height + back + finseal) * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + ((height + back + finseal) * ppi)));
            //side dimension bottom line

            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((height + finseal + back) * ppi)), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + (((front / 2) + height + back + finseal) * ppi)));
            //front dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + (((front / 2) + height + back + finseal) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + height + back + finseal) * ppi)));
            //front dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + ((front + height + back + finseal) * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + ((front + height + back + finseal) * ppi)));
            //front dimension bottom line

            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + height + finseal + back) * ppi)), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + (((height / 2) + front + height + back + finseal) * ppi)));
            //side2 dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + (((height / 2) + front + height + back + finseal) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + (height * 2) + back + finseal) * ppi)));
            //side2 dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + ((front + (height * 2) + back + finseal) * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + ((front + (height * 2) + back + finseal) * ppi)));
            //side2 dimension bottom line

            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + (height * 2) + finseal + back) * ppi)), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + (((back / 2) + front + (height * 2) + back + finseal) * ppi)));
            //back2 dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + (((back / 2) + front + (height * 2) + back + finseal) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + ((height + back) * 2) + finseal) * ppi)));
            //back2 dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + ((front + ((height + back) * 2) + finseal) * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + ((front + ((height + back) * 2) + finseal) * ppi)));
            //back2 dimension bottom line

            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + ((height + back) * 2) + finseal) * ppi)), ((sidemargin - 0.175) * marginppi), (((topmargin - 0.125) * marginppi) + (((finseal / 2) + front + ((height + back) * 2) + finseal) * ppi)));
            //back2 dimension top center line
            gfx.DrawLine(XPens.Black, ((sidemargin - 0.175) * marginppi), ((topmargin + 0.125) * marginppi) + (((finseal / 2) + front + ((height + back) * 2) + finseal) * ppi), ((sidemargin - 0.175) * marginppi), ((topmargin * marginppi) + ((front + ((height + back + finseal) * 2)) * ppi)));
            //back2 dimension lower center line
            gfx.DrawLine(XPens.Black, (((sidemargin - 0.25) * marginppi)), ((topmargin * marginppi) + ((front + ((height + back + finseal) * 2)) * ppi)), ((sidemargin - 0.1) * marginppi), ((topmargin * marginppi) + ((front + ((height + back + finseal) * 2)) * ppi)));
            //back2 dimension bottom line

            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), (((topmargin + 0.175) * marginppi) + (web * ppi)), (((sidemargin - 0.35) * marginppi) + ((repeat / 2) * ppi)), (((topmargin + 0.175) * marginppi) + (web * ppi)));
            //horizontal dimension left center line
            gfx.DrawLine(XPens.Black, (((sidemargin + 0.35) * marginppi) + ((repeat / 2) * ppi)), (((topmargin + 0.175) * marginppi) + (web * ppi)), ((sidemargin * marginppi) + ((repeat) * ppi)), (((topmargin + 0.175) * marginppi) + (web * ppi)));
            //horizontal dimension right center line
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), (((topmargin + 0.1) * marginppi) + (web * ppi)), (sidemargin * marginppi), (((topmargin + 0.25) * marginppi) + (web * ppi)));
            //horizontal dimension left vertical line
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (repeat * ppi)), (((topmargin + 0.1) * marginppi) + (web * ppi)), ((sidemargin * marginppi) + (repeat * ppi)), (((topmargin + 0.25) * marginppi) + (web * ppi)));
            //horizontal dimension right vertical line

            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), ((topmargin - 0.175) * marginppi), (((sidemargin - 0.25) * marginppi) + ((crimp / 2) * ppi)), ((topmargin - 0.175) * marginppi));
            //crimp dimension left center line
            gfx.DrawLine(XPens.Black, (((sidemargin + 0.25) * marginppi) + ((crimp / 2) * ppi)), ((topmargin - 0.175) * marginppi), ((sidemargin * marginppi) + ((crimp) * ppi)), ((topmargin - 0.175) * marginppi));
            //crimp dimension right center line
            gfx.DrawLine(XPens.Black, (sidemargin * marginppi), ((topmargin - 0.1) * marginppi), (sidemargin * marginppi), ((topmargin - 0.25) * marginppi));
            //crimp dimension left vertical line
            gfx.DrawLine(XPens.Black, ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin - 0.1) * marginppi), ((sidemargin * marginppi) + (crimp * ppi)), ((topmargin - 0.25) * marginppi));
            //crimp dimension right vertical line
            //***********************************


            //************DATA LINES*************
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 0.25) * marginppi), (10.25 * marginppi), ((topmargin + 0.25) * marginppi));
            //company name line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 0.55) * marginppi), (10.25 * marginppi), ((topmargin + 0.55) * marginppi));
            //bar name line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 0.85) * marginppi), (10.25 * marginppi), ((topmargin + 0.85) * marginppi));
            //CL&D line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 2.2) * marginppi), (10.25 * marginppi), ((topmargin + 2.2) * marginppi));
            //length line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 2.5) * marginppi), (10.25 * marginppi), ((topmargin + 2.5) * marginppi));
            //width line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 2.8) * marginppi), (10.25 * marginppi), ((topmargin + 2.8) * marginppi));
            //height line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 3.1) * marginppi), (10.25 * marginppi), ((topmargin + 3.1) * marginppi));
            //finseal line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 3.4) * marginppi), (10.25 * marginppi), ((topmargin + 3.4) * marginppi));
            //endseal line

            //************************************


            //*************FILM LABELS************
            gfx.DrawString("Finseal", font, XBrushes.Black, new XRect((sidemargin * marginppi), (topmargin * marginppi), (repeat * ppi), (finseal * ppi)), XStringFormats.Center);
            //finseal label
            gfx.DrawString("Back", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + (finseal * ppi)), (repeat * ppi), ((back) * ppi)), XStringFormats.Center);
            //back label
            gfx.DrawString("Side", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + back) * ppi)), (repeat * ppi), ((height) * ppi)), XStringFormats.Center);
            //side label
            gfx.DrawString("Bleed: 0.125\"", font, XBrushes.Red, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + back + height) * ppi)), (repeat * ppi), (finseal + back + height + (bleed * 2) * ppi)), XStringFormats.Center);
            //bleed label
            gfx.DrawString("Front", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + back + height) * ppi)), (repeat * ppi), (front * ppi)), XStringFormats.Center);
            //front label
            gfx.DrawString("Bleed: 0.125\"", font, XBrushes.Red, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + back + height + front - (bleed * 2)) * ppi)), (repeat * ppi), ((bleed * 2) * ppi)), XStringFormats.Center);
            //bleed label
            gfx.DrawString("Side", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + back + height + front) * ppi)), (repeat * ppi), (height * ppi)), XStringFormats.Center);
            //side label
            gfx.DrawString("Back", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + back + (height * 2) + front) * ppi)), (repeat * ppi), (back * ppi)), XStringFormats.Center);
            //back label
            gfx.DrawString("Finseal", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin * marginppi) + ((finseal + ((back + height) * 2) + front) * ppi)), (repeat * ppi), (finseal * ppi)), XStringFormats.Center);
            //finseal label
            //*************************************


            //*********DIMENSION LABELS************
            gfx.DrawString("web:", font, XBrushes.Black, new XRect(((sidemargin * marginppi) + (repeat * ppi)), (((topmargin - 0.325) * marginppi) + ((web / 2) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //web label
            gfx.DrawString(web.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin * marginppi) + (repeat * ppi)), ((topmargin - 0.175) * marginppi) + ((web / 2) * ppi), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //vertical dimension label
            gfx.DrawString("repeat: " + repeat.ToString() + "\"", font, XBrushes.Black, new XRect((((sidemargin - 0.175) * marginppi) + ((repeat / 2) * ppi)), (((topmargin - 0.075) * marginppi) + (web * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //horizontal dimension label
            gfx.DrawString(finseal.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + ((finseal / 2) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //finseal dimension label
            gfx.DrawString(back.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + (((back / 2) + finseal) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //back dimension label
            gfx.DrawString(height.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + (((height / 2) + back + finseal) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //side dimension label
            gfx.DrawString(front.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + (((front / 2) + height + back + finseal) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //front dimension label
            gfx.DrawString(height.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + (((height / 2) + front + height + back + finseal) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //side2 dimension label
            gfx.DrawString(back.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + (((back / 2) + front + (height * 2) + back + finseal) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //back2 dimension label
            gfx.DrawString(finseal.ToString() + "\"", font, XBrushes.Black, new XRect(((sidemargin - 0.35) * marginppi), (((topmargin - 0.25) * marginppi) + (((finseal / 2) + front + ((height + back) * 2) + finseal) * ppi)), (0.35 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //finseal2 dimension label
            gfx.DrawString(crimp.ToString() + "\"", font, XBrushes.Black, new XRect((sidemargin * marginppi), ((topmargin - 0.425) * marginppi), (0.65 * marginppi), (0.5 * marginppi)), XStringFormats.Center);
            //crimp dimension label
            //*************************************


            //************DATA LABELS**************
            gfx.DrawString("Company Name:   " + companyname, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 0.1) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //company name label
            gfx.DrawString("Bar Name:   " + barname, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 0.4) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //bar name label
            gfx.DrawString("CL&D Tool #:   " + CLDtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 0.7) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //CL&D toolnumber label
            gfx.DrawString("Web:       " + web, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 2.05) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //bar name label
            gfx.DrawString("Repeat:   " + repeat, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 2.35) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //bar name label
            gfx.DrawString("Height:    " + height, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 2.65) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //bar name label
            gfx.DrawString("Finseal:   " + finseal, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 2.95) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //bar name label
            gfx.DrawString("Endseal:  " + crimp, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 3.25) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //bar name label
            //flavor label
            //gfx.DrawString("RELM#:   " + relmtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 1.0) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //relm toolnumber label
            //gfx.DrawString("CL&D#:   " + CLDtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 1.0) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //gfx.DrawString("CL&D#:   " + CLDtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 1.3) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //CL&D toolnumber label
           
            // Approver 1
            gfx.DrawString("Approver: " + approver1, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 4) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            gfx.DrawString("Date:    " + today.ToString("d"), font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 4.55) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 4.5) * marginppi), (10.25 * marginppi), ((topmargin + 4.5) * marginppi));

            // Approver 2
            if (!string.IsNullOrEmpty(approver2))
            {
                gfx.DrawString("Approver: " + approver2, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 5) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
                gfx.DrawString("Date:    " + today.ToString("d"), font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 5.55) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
                gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 5.5) * marginppi), (10.25 * marginppi), ((topmargin + 5.5) * marginppi));
            }

            gfx.DrawString("Note: The measurements listed are accurate.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 6.5) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 1
            gfx.DrawString("Drawing may be scaled down.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 6.75) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 2
            gfx.DrawString("Use the drawing only as a visual aid.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 7) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 3
            gfx.DrawString("Front artwork must not pass bleed lines.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 7.25) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 3
            //*************************************

            document.Save(filename);
            Process.Start(filename);
            //Application.Exit();
        }

        private void btnMasterCase_Click(object sender, EventArgs e)
        {
            masterCase masterCase1 = new masterCase();
            masterCase1.Show();
            this.Hide();
        }

        private void updateTool()
        {
            string toolSelected = listToolNumber.SelectedValue.ToString();
            int bound = dataCLDalt.GetUpperBound(0);
            if (toolSelected == "CUSTOM")
            {
                double widthActual = (double.Parse(txtWidth.Text)) * 2;
                double HeightActual = (double.Parse(txtHeight.Text)) * 2;
                double finsealActual = (double.Parse(txtFinseal.Text)) * 2;
                double webActual = widthActual + HeightActual + finsealActual + 0.25;

                double lengthActual = double.Parse(txtLength.Text);
                double crimpActual = (double.Parse(txtCrimp.Text)) * 2;
                double repeatActual = lengthActual + HeightActual + crimpActual + 0.25;

                exportWeb.Text = webActual.ToString();
                exportRepeat.Text = repeatActual.ToString();
                exportFinseal.Text = (finsealActual / 2).ToString();
                exportEndseal.Text = (crimpActual / 2).ToString();
            }
            else
            {
                for (int i = 0; i < bound; i++)
                {
                    if (dataCLDalt[i, 0].ToString() == toolSelected)
                    {
                        exportWeb.Text = dataCLDalt[i, 1].ToString();
                        exportRepeat.Text = dataCLDalt[i, 2].ToString();
                        exportFinseal.Text = dataCLDalt[i, 3].ToString();
                        exportEndseal.Text = dataCLDalt[i, 4].ToString();
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateTool();
        }

        private void listToolNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateTool();
        }
    }
}