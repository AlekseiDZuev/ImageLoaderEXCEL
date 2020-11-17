using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ImageLoader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string linkRange = textBox1.Text;
            string imageColumn = textBox2.Text;
            int imageSize = Convert.ToInt32(textBox3.Text);
            string[] linkRangeArr  = linkRange.Split(new char[] { ':' });
            string firstVal = linkRangeArr[0];
            Match matches = Regex.Match(firstVal, @"\d+");
            int i = Convert.ToInt32(matches.Value);
            string secondVal = linkRangeArr[1];
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range linkCycleRange = activeWorksheet.get_Range(firstVal,secondVal);
            foreach (Excel.Range cellNum in linkCycleRange)
            {
                string picPath = @""+cellNum.Text;
                if (picPath == "")
                {
                    i = i + 1;
                }
                else
                {
                    Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)activeWorksheet.get_Range(@"" + imageColumn + i);
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    activeWorksheet.Shapes.AddPicture(picPath, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, imageSize, imageSize);
                    oRange.Columns.EntireColumn.AutoFit();
                    oRange.RowHeight = imageSize;
                    i = i + 1;
                }
            }
        }
    }
}
