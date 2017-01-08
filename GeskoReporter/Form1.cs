using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Reflection;

namespace GeskoReporter
{
    public partial class Form1 : Form
    {
        object missing = Type.Missing;
        Microsoft.Office.Interop.Excel.Application exe = null;

        Collection<CallRecord> callRecords;
        string firstDate, lastDate;
        decimal sum = 0, sumFF = 0, sumRK = 0;
        int einheiten = 0, einheitenFF = 0, einheitenRK = 0;
        SortedDictionary<string, decimal> sumRKperMonth = new SortedDictionary<string, decimal>();
        SortedDictionary<string, int> einheitenRKperMonth = new SortedDictionary<string, int>();

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

            try
            {
                _Worksheet worksheet = (_Worksheet)workbook.Worksheets[1];
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
                            else if (j == COL_EINHEITEN.index)
                                record.phoneUnits = tmpRange.Value2.ToString();
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

            for (int i = 0; i < this.callRecords.Count; i++)
            {
                if (this.callRecords[i].phoneName.IndexOf("FF") > 0)
                {
                    einheitenFF += int.Parse(this.callRecords[i].phoneUnits);
                    sumFF += decimal.Parse(this.callRecords[i].cost);
                }
                else if (this.callRecords[i].phoneName.IndexOf("RK") > 0)
                {
                    einheitenRK += int.Parse(this.callRecords[i].phoneUnits);
                    sumRK += decimal.Parse(this.callRecords[i].cost);
                    if (this.callRecords[i].date.IndexOf('.') < 0)
                    {
                        //MessageBox.Show("Fehler: Datum von Eintrag #" + (i+1).ToString() + " nicht erkannt (" + this.callRecords[i].date + ")");
                        double tmp1 = double.Parse(this.callRecords[i].date);
                        DateTime tmp2 = DateTime.FromOADate(tmp1);
                        this.callRecords[i].date = tmp2.ToString("dd.MM.yyyy");
                        //return;
                    }
                    string[] date = this.callRecords[i].date.Split('.');
                    string tmpKey = date[2] + " / " + date[1];
                    if (!sumRKperMonth.ContainsKey(tmpKey))
                    {
                        sumRKperMonth.Add(tmpKey, 0);
                        einheitenRKperMonth.Add(tmpKey, 0);
                    }
                    sumRKperMonth[tmpKey] += decimal.Parse(this.callRecords[i].cost);
                    einheitenRKperMonth[tmpKey] += int.Parse(this.callRecords[i].phoneUnits);
                }
                else
                {
                    MessageBox.Show("Fehler: Weder FF- noch RK-Telefonat (" + this.callRecords[i].phoneName + ")");
                }
                einheiten += int.Parse(this.callRecords[i].phoneUnits);
                sum += decimal.Parse(this.callRecords[i].cost);

                if (i == 0)
                {
                    lastDate = this.callRecords[i].date;
                }
                firstDate = this.callRecords[i].date;
            }
            txtSumFF.Text = sumFF.ToString() + " €";
            txtSumRK.Text = sumRK.ToString() + " €";
            txtSum.Text = sum.ToString() + " €";

            createExcel();
        }


        private void createExcel()
        {
            lblProcess.Visible = true;
            lblProcess.Text = "Erzeuge Excel-Abrechnung...";
            
            Microsoft.Office.Interop.Excel.Application exe = null;
            string excelFilePath = txtFileName.Text.Substring(0, txtFileName.Text.LastIndexOf("\\")) + "\\RK-Abrechnung_" + firstDate + "-" + lastDate + ".xlsx";
            try
            {
                exe = new Microsoft.Office.Interop.Excel.Application();
                exe.DisplayAlerts = false;
                Workbooks workbooks = exe.Workbooks;
                _Workbook workbook = (_Workbook)(workbooks.Add(XlWBATemplate.xlWBATWorksheet));

                lblProcess.Text = "Erzeuge Excel-Abrechnung... (Übersicht)";
                _Worksheet worksheet = (Worksheet)workbook.ActiveSheet;
                worksheet.Name = "Übersicht";
                ((Range)worksheet.Cells[1, 1]).EntireColumn.ColumnWidth = 20;
                ((Range)worksheet.Cells[1, 2]).EntireColumn.ColumnWidth = 20;
                
                Range range;
                int row = 1;
                range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 3]];
                range.Merge();
                range.EntireRow.Font.Bold = true;
                range.Value2 = "Abrechnungsübersicht GESKO Telefonanlage";
                row++;
                range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 3]];
                range.Merge();
                range.Value2 = "Zeitraum: " + firstDate + " bis " + lastDate;
                range.Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black.ToArgb();

                row += 3;
                range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 1]];
                range.Value2 = "Monat";
                range = worksheet.Range[worksheet.Cells[row, 2], worksheet.Cells[row, 2]];
                range.Value2 = "Einheiten";
                range = worksheet.Range[worksheet.Cells[row, 3], worksheet.Cells[row, 3]];
                range.Value2 = "Betrag in €";
                row++;
                int tmpRow = row;
                foreach (KeyValuePair<string, decimal> kvp in this.sumRKperMonth)
                {
                    range = worksheet.Range[worksheet.Cells[tmpRow, 1], worksheet.Cells[tmpRow, 1]];
                    range.Value2 = kvp.Key;
                    range = worksheet.Range[worksheet.Cells[tmpRow, 3], worksheet.Cells[tmpRow, 3]];
                    range.Value2 = kvp.Value.ToString();
                    tmpRow++;
                }
                tmpRow = row;
                foreach (KeyValuePair<string, int> kvp in this.einheitenRKperMonth)
                {
                    range = worksheet.Range[worksheet.Cells[tmpRow, 2], worksheet.Cells[tmpRow, 2]];
                    range.Value2 = kvp.Value.ToString();
                    tmpRow++;
                }
                row = tmpRow;
                range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 2]];
                range.Merge();
                range.Font.Bold = true;
                range.Value2 = "Summe in €: ";
                range = worksheet.Range[worksheet.Cells[row, 3], worksheet.Cells[row, 3]];
                range.Font.Bold = true;
                range.Value2 = this.sumRK.ToString();

                lblProcess.Text = "Erzeuge Excel-Abrechnung... (Einzelverbindungsnachweis)";
                workbook.Worksheets.Add(Missing.Value, workbook.Worksheets[1]);
                worksheet = (Worksheet)workbook.Worksheets[2];
                worksheet.Name = "Einzelverbindungsnachweis";
                
                row = 1;
                range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 1]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_TEILNEHMER.name;
                range = worksheet.Range[worksheet.Cells[row, 2], worksheet.Cells[row, 2]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_NAME.name;
                range = worksheet.Range[worksheet.Cells[row, 3], worksheet.Cells[row, 3]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_RUFNUMMER.name;
                range = worksheet.Range[worksheet.Cells[row, 4], worksheet.Cells[row, 4]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_DATUM.name;
                range = worksheet.Range[worksheet.Cells[row, 5], worksheet.Cells[row, 5]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_UHRZEIT.name;
                range = worksheet.Range[worksheet.Cells[row, 6], worksheet.Cells[row, 6]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_DAUER.name;
                range = worksheet.Range[worksheet.Cells[row, 7], worksheet.Cells[row, 7]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_EINHEITEN.name;
                range = worksheet.Range[worksheet.Cells[row, 8], worksheet.Cells[row, 8]];
                range.EntireRow.Font.Bold = true;
                range.Value2 = COL_BETRAG.name;

                for (int i = 0; i < this.callRecords.Count; i++)
                {
                    row++;
                    range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 1]];
                    range.Value2 = this.callRecords[i].phoneId;
                    range = worksheet.Range[worksheet.Cells[row, 2], worksheet.Cells[row, 2]];
                    range.Value2 = this.callRecords[i].phoneName;
                    range = worksheet.Range[worksheet.Cells[row, 3], worksheet.Cells[row, 3]];
                    range.Value2 = this.callRecords[i].phoneNumber;
                    range = worksheet.Range[worksheet.Cells[row, 4], worksheet.Cells[row, 4]];
                    range.Value2 = this.callRecords[i].date;
                    range = worksheet.Range[worksheet.Cells[row, 5], worksheet.Cells[row, 5]];
                    range.Value2 = this.callRecords[i].time;
                    range = worksheet.Range[worksheet.Cells[row, 6], worksheet.Cells[row, 6]];
                    range.Value2 = this.callRecords[i].duration;
                    range = worksheet.Range[worksheet.Cells[row, 7], worksheet.Cells[row, 7]];
                    range.Value2 = this.callRecords[i].phoneUnits;
                    range = worksheet.Range[worksheet.Cells[row, 8], worksheet.Cells[row, 8]];
                    range.Value2 = this.callRecords[i].cost;
                }

                worksheet = (Worksheet)workbook.Worksheets[1];
                worksheet.Activate();
                workbook.SaveAs(excelFilePath);
                workbook.Close();

                //cleanup
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook) != 0) ;
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks) != 0) ;
            }
            catch (Exception exc)
            { }
            finally
            {
                lblProcess.Visible = false;

                // Cleanup 
                GC.Collect();
                GC.WaitForPendingFinalizers();

                exe.Quit();
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(exe) != 0) ;
            }

            MessageBox.Show("RK-Abrechnung erfolgreich unter '" + excelFilePath + "' erstellt.");
        }
    }
}
