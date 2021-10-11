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
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Charting;

namespace MingoPackaging
{
    public partial class exportForm : Form
    {
        public string CLDtoolnumber { get; set; }
        public string relmtoolnumber { get; set; }
        public double web { get; set; }
        public double repeat { get; set; }
        public double finseal { get; set; }
        public double height { get; set; }
        public double crimp { get; set; }

        public exportForm(string CLDtoolnumber1, string web1, string repeat1, string finseal1, string height1, string crimp1)
        {
            InitializeComponent();
            CLDtoolnumber = CLDtoolnumber1;
            //relmtoolnumber = relmtoolnumber1;
            web = Double.Parse(web1);
            repeat = Double.Parse(repeat1);
            finseal = Double.Parse(finseal1);
            height = Double.Parse(height1);
            crimp = Double.Parse(crimp1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            double front = (web - (finseal * 2) - (height * 2)) / 2;
            double back = front / 2;
            double ppiWeb;
            double ppiRepeat;
            string companyname = txtCompanyName.Text;
            string barname = txtBarName.Text;
            string flavor = txtBarFlavor.Text;
            double sidemargin = 1;
            double topmargin = 0.75;
            string filename = companyname + "_" + barname + "_" + flavor + "_" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ".pdf";
            string approver1 = "Fred Brayton";
            string approver2 = "";
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
            XFont font3 = new XFont("Times New Roman", 8, XFontStyle.Italic);

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
            //flavor line
            //gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 1.15) * marginppi), (10.25 * marginppi), ((topmargin + 1.15) * marginppi));
            //relm line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 1.15) * marginppi), (10.25 * marginppi), ((topmargin + 1.15) * marginppi));
            //gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 1.45) * marginppi), (10.25 * marginppi), ((topmargin + 1.45) * marginppi));
            //CL&D line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 2.5) * marginppi), (10.25 * marginppi), ((topmargin + 2.5) * marginppi));
            //approver 1 line
            gfx.DrawLine(XPens.Black, (8.25 * marginppi), ((topmargin + 4.5) * marginppi), (10.25 * marginppi), ((topmargin + 4.5) * marginppi));
            //approver2 line
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
            gfx.DrawString("Flavor:   " + flavor, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 0.7) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //flavor label
            //gfx.DrawString("RELM#:   " + relmtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 1.0) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //relm toolnumber label
            gfx.DrawString("CL&D#:   " + CLDtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 1.0) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //gfx.DrawString("CL&D#:   " + CLDtoolnumber, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 1.3) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //CL&D toolnumber label
            gfx.DrawString("Approver: " + approver1, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 2) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //approver 1 name label
            gfx.DrawString("Approver: " + approver2, font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 4) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //approver 2 name label
            gfx.DrawString("Date:    " + today.ToString("d"), font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 2.55) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //date 1 label
            gfx.DrawString("Date:    " + today.ToString("d"), font2, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 4.55) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //date 2 label
            gfx.DrawString("Note: The measurements listed are accurate.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 6.5) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 1
            gfx.DrawString("The drawing is scaled down, and is not accurate.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 6.75) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 2
            gfx.DrawString("Use the drawing only as a visual aid.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 7) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 3
            gfx.DrawString("Front artwork must not pass bleed lines.", font3, XBrushes.Black, new XRect((8.25 * marginppi), ((topmargin + 7.25) * marginppi), (2 * marginppi), (0.5 * marginppi)), XStringFormats.TopLeft);
            //warning label 3
            //*************************************

            document.Save(filename);
            Process.Start(filename);
            Application.Exit();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
