using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection; 

namespace RWProcess
{
    public partial class frmControl : Form
    {
        string[] files;

        public frmControl()
        {
            InitializeComponent();
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            List<string> cleanedList = new List<string>();
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtFilename.Text = fbd.SelectedPath;
                    files = Directory.GetFiles(fbd.SelectedPath.ToString());
                    foreach(string filename in files)
                    {
                        if (Path.GetFileName(filename).StartsWith("Consignment") ||
                            Path.GetFileName(filename).StartsWith("Copy of Consignment"))
                        {
                            AddStatus("Currently evaluating - " + filename);
                            cleanedList.Add(filename);
                        }
                        else
                            AddStatus("Not considering - " + filename);
                    }
                    txtTotal.Text = cleanedList.Count.ToString();
                    files = cleanedList.ToArray();
                    btnProcess.Enabled = true;
                }
            }

            
/*
            // Displays an OpenFileDialog so the user can select a Cursor.  
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls";
            openFileDialog1.Title = "Select an Excel File";

            // Show the Dialog.  
            // If the user clicked OK in the dialog and  
            // a .CUR file was selected, open it.  
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Assign the cursor in the Stream to the Form's Cursor property.  
                //this.Cursor = new Cursor(openFileDialog1.OpenFile());
                txtFilename.Text = openFileDialog1.FileName;
            }   */
        }

        private void frmControl_Load(object sender, EventArgs e)
        {
            btnProcess.Enabled = false;
            this.Text = this.Text +  " " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        private void AddStatus(String status)
        {
            txtStatus.Text = "[" + DateTime.Now + "] " + status + "\r\n" + txtStatus.Text;
        }

        private void ProcessFile(string filename)
        {
            string dirprefix = @"C:\RWProcess\";
            string targetfilename = "test";
            bool GotData = true;
            int RunningNumber = 1;

            string re1 = ".*?"; // Non-greedy match on filler
            string re2 = "(\\d+)";  // Integer Number 1

            if (!Directory.Exists(dirprefix))
            {
                Directory.CreateDirectory(dirprefix);
            }

            Regex r = new Regex(re1 + re2, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(Path.GetFileName(filename));
            if (m.Success)
            {
                targetfilename = "MEL Equipment Condition Report - Batch " + m.Groups[1].ToString();
                txtConsignment.Text = m.Groups[1].ToString();
            }
            else
            {
                targetfilename = "test" + String.Format("{000}", RunningNumber++);
            }

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
            txtStatus.Text += "Opening spreadsheet...";
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            AddStatus("Retrieving range of used cells within the worksheet");
            Excel.Range xlRange = xlWorksheet.UsedRange;
            for (int i = 3; i <= 15; i++)
            {
                xlWorksheet.Rows[2].Delete();
            }
            AddStatus("Unwanted top rows deleted!");

            // Get bottom rows
            int CurRow = 4;
            while (GotData)
            {
                //MessageBox.Show("Current value is : " + xlRange[CurRow, 1].Value2);
                if (String.IsNullOrEmpty(Convert.ToString(xlRange[CurRow, 1].Value2)))
                {
                    GotData = !GotData;
                    AddStatus("Found Empty Row");
                    CurRow += 2;
                }
                else
                {

                    AddStatus("skip row...");
                    CurRow++;
                }
            }

            for (int i = 12; i <= 27; i++)
            {
                xlWorksheet.Rows[CurRow].Delete();
            }
            AddStatus("Unwanted bottom rows deleted!");

            Excel.Range UsedRange = xlWorksheet.Range["C1"];
            UsedRange.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            AddStatus("Add Column at C1");

            UsedRange = xlWorksheet.Range["C4"];
            UsedRange.Value = "Matching with RW Audit";
            UsedRange.Interior.Color = Excel.XlRgbColor.rgbYellow;

            UsedRange = xlWorksheet.Range["N4"];
            UsedRange.Value = "Quoted Price (MYR)";
            UsedRange.Interior.Color = Excel.XlRgbColor.rgbGray ;

            UsedRange = xlWorksheet.Range["O4"];
            UsedRange.Value = "Revised Price (MYR)";
            UsedRange.Interior.Color = Excel.XlRgbColor.rgbYellow;

            UsedRange = xlWorksheet.Range["N5", "O" + (CurRow - 3).ToString()];
            UsedRange.NumberFormat = "#,###.00";

            UsedRange = xlWorksheet.Range["O" + (CurRow - 2).ToString()];
            UsedRange.Formula = "=SUM(O5:" + "O" + (CurRow - 3).ToString();

            UsedRange = xlWorksheet.Range["M" + (CurRow-2).ToString()];
            UsedRange.Value = "Total";
            UsedRange.Font.Bold = true;

            UsedRange = xlWorksheet.Range["M" + (CurRow - 2).ToString(), "O" + (CurRow - 2).ToString()];
            foreach (Excel.Range cell in UsedRange.Cells)
            {
                cell.BorderAround2();
            }

            UsedRange = xlWorksheet.Range["N4", "Q" + (CurRow - 3).ToString()];
            foreach(Excel.Range cell in UsedRange.Cells)
            {
                cell.BorderAround2();
            }
            //UsedRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            //Excel.Borders border = UsedRange.Borders;
            //border


            UsedRange = xlWorksheet.Range["P4"];
            UsedRange.Value = "Remarks";
            UsedRange.Interior.Color = Excel.XlRgbColor.rgbYellow;

            UsedRange = xlWorksheet.Range["Q4"];
            UsedRange.Value = "Physical Serial No";
            UsedRange.Interior.Color = Excel.XlRgbColor.rgbYellow;
            // MessageBox.Show("Peekaboo!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Warning
            //     );
            AddStatus("Saving file: " + dirprefix + targetfilename);
            xlWorkbook.SaveAs(dirprefix + targetfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            AddStatus("Initiating garbage collection");
            GC.Collect();
            GC.WaitForPendingFinalizers();

            AddStatus("Release COM Objects");
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            AddStatus("Close and release Excel workbooks");
            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            AddStatus("Quit Excel and release Excel objects");
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            int FileCounter = 1; 

            foreach(string file in files)
            {
                txtProcessing.Text = FileCounter++.ToString();
                ProcessFile(file);               
            }
        }
    }
}
