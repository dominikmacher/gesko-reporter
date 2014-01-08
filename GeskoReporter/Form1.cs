using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;

namespace GeskoReporter
{
    public partial class Form1 : Form
    {
        object missing = Type.Missing;
        Microsoft.Office.Interop.Excel.Application exe = null;

        Collection<CallRecord> callRecords;

        readonly int START_ROW = 6;
        readonly int END_ROW_OFFSET = 3;
        readonly string RELEVANT_VERBINDUNG = "Von";
        readonly Column COL_VERBINDUNG;
        readonly Column COL_TEILNEHMER;
        readonly Column COL_NAME;
        readonly Column COL_RUFNUMMER;
        readonly Column COL_DATUM;
        readonly Column COL_UHRZEIT;
        readonly Column COL_DAUER;
        readonly Column COL_EINHEITEN;
        readonly Column COL_BETRAG;

        public Form1()
        {
            InitializeComponent();
            this.callRecords = new Collection<CallRecord>();
            COL_VERBINDUNG = new Column(1, "Verbindung");
            COL_TEILNEHMER = new Column(2, "Teilnehmer");
            COL_NAME = new Column(4, "Name");
            COL_RUFNUMMER = new Column(5, "Rufnummer");
            COL_DATUM = new Column(9, "Datum");
            COL_UHRZEIT = new Column(10, "Uhrzeit");
            COL_DAUER = new Column(11, "Dauer");
            COL_EINHEITEN = new Column(12, "Einheiten");
            COL_BETRAG = new Column(15, "Betrag"); 
        }

        private void btnFileDialog_Click(object sender, EventArgs e)
        {
            this.openFileDialog.ShowDialog();
        }

        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            txtFileName.Text = this.openFileDialog.FileName;
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFileName.Text))
            {
                MessageBox.Show("Keine Datei ausgewählt!");
                return;
            }

            // Parse existing excel
            exe = new Microsoft.Office.Interop.Excel.Application();
            _Workbook workbook = null;
            workbook = exe.Workbooks.Open(txtFileName.Text);
            workbook.Activate();

            _Worksheet worksheet = (_Worksheet)workbook.Worksheets[1];
            try
            {
                Range xlRange = worksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                
                lblProcess.Visible = true;
                lblProcess.Text = "";
                for (int i = START_ROW; i <= rowCount-END_ROW_OFFSET; i++)
                {
                    Range tmpRange = worksheet.Range[worksheet.Cells[i, 1], worksheet.Cells[i, 1]];
                    if (String.IsNullOrEmpty(tmpRange.Value2.ToString()))
                        break;
                    else if (String.Compare(tmpRange.Value2.ToString(), RELEVANT_VERBINDUNG) == 0)
                    {
                        lblProcess.Text = "Verarbeite Zeile " + i.ToString() + " / " + (rowCount - END_ROW_OFFSET).ToString();
                        CallRecord record = new CallRecord();
                        for (int j = 1; j <= colCount; j++)
                        {
                            tmpRange = worksheet.Range[worksheet.Cells[i, j], worksheet.Cells[i, j]];
                            if (j==COL_TEILNEHMER.index)
                                record.phoneId = tmpRange.Value2.ToString();
                            else if (j==COL_NAME.index)
                                record.phoneName = tmpRange.Value2.ToString();
                            else if (j == COL_RUFNUMMER.index)
                                record.phoneNumber = tmpRange.Value2.ToString();
                            else if (j == COL_DATUM.index)
                                record.date = tmpRange.Value2.ToString();
                            else if (j == COL_UHRZEIT.index)
                                record.time = tmpRange.Value2.ToString();
                            else if (j == COL_DAUER.index)
                                record.duration = tmpRange.Value2.ToString();
                            else if (j == COL_BETRAG.index)
                                record.cost = tmpRange.Value2.ToString();
                        }
                        this.callRecords.Add(record);
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
            finally
            {
                lblProcess.Visible = false;
                workbook.Close();

                //cleanup
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook) != 0) ;
                //while (System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks) != 0) ;

                // Cleanup 
                GC.Collect();
                GC.WaitForPendingFinalizers();

                exe.Quit();
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(exe) != 0) ;
            }

            decimal sum = 0, sumFF = 0, sumRK = 0;
            for (int i = 0; i < this.callRecords.Count; i++)
            {
                if (this.callRecords[i].phoneName.IndexOf("FF") > 0)
                {
                    sumFF += decimal.Parse(this.callRecords[i].cost);
                }
                else if (this.callRecords[i].phoneName.IndexOf("RK") > 0)
                {
                    sumRK += decimal.Parse(this.callRecords[i].cost);
                }
                else
                {
                    MessageBox.Show("Fehler: Weder FF- noch RK-Telefonat (" + this.callRecords[i].phoneName + ")");
                }
                sum += decimal.Parse(this.callRecords[i].cost);
            }
            txtSumFF.Text = sumFF.ToString() + " €";
            txtSumRK.Text = sumRK.ToString() + " €";
            txtSum.Text = sum.ToString() + " €";
        }


        private void createExcel()
        {
            Microsoft.Office.Interop.Excel.Application exe = null;
            try
            {
                exe = new Microsoft.Office.Interop.Excel.Application();
                Workbooks workbooks = exe.Workbooks;
                _Workbook workbook = (_Workbook)(workbooks.Add(XlWBATemplate.xlWBATWorksheet));
                
                /*worksheet.Name = "Deckblatt";
                ((Range)worksheet.Cells[1, 1]).EntireColumn.ColumnWidth = 5;
                ((Range)worksheet.Cells[1, 2]).EntireColumn.ColumnWidth = 20;
                ((Range)worksheet.Cells[1, 3]).EntireColumn.ColumnWidth = 45;
                ((Range)worksheet.Cells[1, 4]).EntireColumn.ColumnWidth = 20;

                _CreateHeader(worksheet, 80, "");
                _CreateSubHeader(worksheet);

                Range range;
                int row = subHeaderStartRow + 4;

                //Report für:
                range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 1]];
                range.Value2 = "Report für:";
                range = worksheet.Range[worksheet.Cells[row, 2], worksheet.Cells[row, 2]];
                range.Font.Bold = true;
                range.Value2 = reportFuer;
                */

                workbook.SaveAs("D:\\test.xls");
                workbook.Close();

                //cleanup
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook) != 0) ;
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks) != 0) ;

            }
            catch (Exception exc)
            {

            }
            finally
            {
                // Cleanup 
                GC.Collect();
                GC.WaitForPendingFinalizers();

                exe.Quit();
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(exe) != 0) ;
            }
        }
    }
}
