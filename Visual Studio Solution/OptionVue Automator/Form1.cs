using Ookii.Dialogs.WinForms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OptionVue_Automator

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        String OutputPath;
        String OutDir;
        String SelectedFolder;
        string[] Folder;
        bool halt = false;
      

        Excel.Application oXL= new Microsoft.Office.Interop.Excel.Application();
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;

        public string fsymbol { get; private set; }
        public string fdate { get; private set; }
        public string ftime { get; private set; }


/////////////// MAIN PROG //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private void button1_Click(object sender, EventArgs e)
        {
            SelectFolder();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (halt) { return; }

                else {

                    if (startBtn.Text == "View Output")
                    {
                        Process.Start(SelectedFolder+"/Output");
                        return;
                    }


                    startBtn.Enabled = false;
                    label2.Visible = true;

                    // CREATE OUTPUT FILE///////////////////////////////////////////////////////
                    if (File.Exists(OutputPath))
                    {
                        LanguageUtils.IgnoreErrors(() => File.Delete(OutputPath));
                    }
                    else if (!Directory.Exists(OutDir))
                    {
                        Directory.CreateDirectory(OutDir);
                    }

                    // CHECK IF EXCEL IS INSTALLED/////////////////////////////////////////////////
                    if (oXL == null)
                    {
                        MessageBox.Show("Excel is not installed!!");
                        return;
                    }
                    // INSERT HEADER ROW///////////////////////////////////////////////////
                    object misValue = System.Reflection.Missing.Value;
                    oWB = oXL.Workbooks.Add(misValue);
                    oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);
                    oSheet.Cells[1, 1] = "SYMBOL";
                    oSheet.Cells[1, 2] = "DATE";
                    oSheet.Cells[1, 3] = "TIME";
                    oSheet.Cells[1, 4] = "AVERAGE IV";


                   // CALCULATE AVERAGE FROM FILES AND PUSH RESULTS TO OUTPUT FILE//////////////

                    foreach (string excelFile in Folder)
                    {
                        double avg = calculateAVG(getMIDIV(excelFile));
                        fsymbol = SplitFilename(excelFile, 1);
                        fdate = SplitFilename(excelFile, 2);
                        ftime = SplitFilename(excelFile, 3);

                        xlRange = oSheet.UsedRange;
                        int rowCount = xlRange.Rows.Count;

                        oSheet.Cells[rowCount + 1, 1] = fsymbol;
                        oSheet.Cells[rowCount + 1, 2] = fdate;
                        oSheet.Cells[rowCount + 1, 3] = ftime;
                        oSheet.Cells[rowCount + 1, 4] = avg;
                    }
                    LanguageUtils.IgnoreErrors(() => oWB.SaveAs(OutputPath));

                    oXL.Quit();

                    label3.Visible = true;
                    startBtn.Enabled = true;
                    startBtn.Text = "View Output";




                }
            
         

          
        }


 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        private void SelectFolder()
        {
            VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog();

            dlg.SelectedPath = Properties.Settings.Default.StoreFolder;

            dlg.ShowNewFolderButton = true;

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SelectedFolder = dlg.SelectedPath;
                OutputPath = SelectedFolder + "/Output/Output Average Implied Volatility.xlsx";
                OutDir = SelectedFolder + "/Output/";

                Folder = Directory.GetFiles(SelectedFolder, "*.xlsx");

                if(Folder.Length == 0)
                {
                    label4.Visible = true;
                    halt = true;
                    return;

                }else
                {
                    label4.Visible = false;

                    label1.Visible = true;
                    halt = false;
                }

                textBox1.Text = dlg.SelectedPath;
            }

        }

        private void createOutputFile()
        {
            if (File.Exists(OutputPath))
            {
                LanguageUtils.IgnoreErrors(() => File.Delete(OutputPath));
            }else if (!Directory.Exists(OutDir))
            {
                Directory.CreateDirectory(OutDir);
            }
            object misValue = System.Reflection.Missing.Value;

            if (oXL == null)
            {
                MessageBox.Show("Excel is not installed!!");
                return;
            }

            oWB = oXL.Workbooks.Add(misValue);
            oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);
            oSheet.Cells[1, 1] = "SYMBOL";
            oSheet.Cells[1, 2] = "DATE";
            oSheet.Cells[1, 3] = "TIME";
            oSheet.Cells[1, 4] = "AVERAGE IV";

            LanguageUtils.IgnoreErrors(() => oWB.SaveAs(OutputPath));

            oXL.Quit();


        }

       

        private List<double> getMIDIV(string fname)
        {
            List<double> MIDIVList = new List<double>();
            xlWorkbook = xlApp.Workbooks.Open(fname);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                if (xlWorksheet.Cells[i, 8].value != null)
                {
                    MIDIVList.Add(xlWorksheet.Cells[i, 8].value);
                }

            }
            xlWorkbook.Close();
            xlApp.Quit();

            return MIDIVList;
        }

        private double calculateAVG(List<double> list)
        {
            return list.Average();


        }

        private string SplitFilename(string fileinput, int choice)
        {
            string symbol = null;
            string date = null;
            string time = null;
            string avg = null;

            string fileName = Path.GetFileNameWithoutExtension(fileinput);
            string trim = Regex.Replace(fileName, @"s", "");

            
            time = fileName.Substring(fileName.Length - 4);
            int hh = Convert.ToInt32(time.Substring(0, 2));
            int mm = Convert.ToInt32(time.Substring(2, 2));
            if (hh > 12)
            {
                time = Convert.ToString((hh - 12)) + ":" + Convert.ToString((mm) + " PM");
            }
            else if (hh == 12)
            {
                time = Convert.ToString((hh)) + ":" + Convert.ToString((mm) + " PM");

            }
            else
            {
                time = Convert.ToString((hh)) + ":" + Convert.ToString((mm) + " AM");

            }

            trim =  trim.Remove(trim.Length - 5);

            date = trim.Substring(trim.Length - 8);
            date = date.Insert(2, "-");
            date = date.Insert(5, "-");


            trim = trim.Remove(trim.Length - 9);

             symbol = trim;


            if (choice == 1)
            {
                return symbol.ToUpper();
            }
            else if (choice == 2)
            {
                return date;
            }
            else if (choice == 3)
            {
                return time;
            }
            else
            {
                return null;
            }
   


            

        }

        public static class LanguageUtils
        {
            /// <summary>
            /// Runs an operation and ignores any Exceptions that occur.
            /// Returns true or falls depending on whether catch was
            /// triggered
            /// </summary>
            /// <param name="operation">lambda that performs an operation that might throw</param>
            /// <returns></returns>
            public static bool IgnoreErrors(Action operation)
            {
                if (operation == null)
                    return false;
                try
                {
                    operation.Invoke();
                }
                catch
                {
                    return false;
                }

                return true;
            }

            /// <summary>
            /// Runs an function that returns a value and ignores any Exceptions that occur.
            /// Returns true or falls depending on whether catch was
            /// triggered
            /// </summary>
            /// <param name="operation">parameterless lamda that returns a value of T</param>
            /// <param name="defaultValue">Default value returned if operation fails</param>
            public static T IgnoreErrors<T>(Func<T> operation, T defaultValue = default(T))
            {
                if (operation == null)
                    return defaultValue;

                T result;
                try
                {
                    result = operation.Invoke();
                }
                catch
                {
                    result = defaultValue;
                }

                return result;
            }
        }

       
    }
}
