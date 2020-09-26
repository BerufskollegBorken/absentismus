using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Absentismus
{
    /// <summary>
    /// Offene und unentschuldigte Abwesenheiten an Schultagen.
    /// </summary>
    public class Abwesenheiten : List<Abwesenheit>
    {
        public Abwesenheiten()
        {
            var dateien = Directory.GetFiles(@"C:\Users\bm\AppData\Local\Temp\").OrderBy(f => f);

            var datei = (from d in dateien where d.Contains("AbsencePerStudent") select d).LastOrDefault();

            using (var reader = new StreamReader(Global.InputAbwesenheitenCsv))
            {
                reader.ReadLine();

                while (!reader.EndOfStream)
                {
                    Abwesenheit abwesenheit = new Abwesenheit();
                    var line = reader.ReadLine();
                    var values = line.Split('\t');
                    abwesenheit.Name = values[0];                    
                    abwesenheit.StudentId = Convert.ToInt32(Convert.ToString(values[1]));
                    abwesenheit.Klasse = (values[3] != null) ? Convert.ToString(values[3]) : "";
                    abwesenheit.Datum = DateTime.ParseExact(values[4], "dd.MM.yy", CultureInfo.InvariantCulture);
                    abwesenheit.Fehlstunden = (values[6] != null && values[6] != null) ? Convert.ToInt32(values[6]) : 0;
                    abwesenheit.Fehlminuten = (values[7] != null && values[7] != null) ? Convert.ToInt32(values[7]) : 0;
                    abwesenheit.GanzerFehlTag = (values[15] != null && values[15] != null) ? Convert.ToInt32(values[15]) : 0;
                    abwesenheit.Grund = (values[8] != null && values[8] != null) ? Convert.ToString(values[8]) : "";
                    abwesenheit.Text = (values[9] != null && values[9] != null) ? Convert.ToString(values[9]) : "";
                    abwesenheit.Status = (values[14] != null && values[14] != null) ? Convert.ToString(values[14]) : "";

                    if (
                            abwesenheit.Status == "offen" ||
                            abwesenheit.Status == "nicht entsch.")
                    {
                        this.Add(abwesenheit);
                    }
                }
            }


            ////Create COM Objects. Create a COM object for everything that is referenced
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Global.InputAbwesenheitenCsv);
            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //Excel.Range xlRange = xlWorksheet.UsedRange;
            
            //for (int i = 2; true ; i++)
            //{
            //    Abwesenheit abwesenheit = new Abwesenheit();
            //    abwesenheit.Name = (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null) ? Convert.ToString(xlRange.Cells[i, 1].Value2) : "";
                
            //    if (abwesenheit.Name == "")
            //    {
            //        break;
            //    }

            //    abwesenheit.StudentId = Convert.ToInt32((xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null) ? Convert.ToString(xlRange.Cells[i, 2].Value2) : 0);
            //    abwesenheit.Klasse = (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null) ? Convert.ToString(xlRange.Cells[i, 4].Value2) : "";                
            //    abwesenheit.Datum = DateTime.FromOADate(double.Parse(Convert.ToString(xlRange.Cells[i, 5].Value2))); 
            //    abwesenheit.Fehlstunden = (xlRange.Cells[i, 7] != null && xlRange.Cells[i, 7].Value2 != null) ? Convert.ToInt32(xlRange.Cells[i, 7].Value2) : 0;
            //    abwesenheit.Fehlminuten = (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null) ? Convert.ToInt32(xlRange.Cells[i, 8].Value2) : 0;
            //    abwesenheit.GanzerFehlTag = (xlRange.Cells[i, 16] != null && xlRange.Cells[i, 16].Value2 != null) ? Convert.ToInt32(xlRange.Cells[i, 16].Value2) : 0;
            //    abwesenheit.Grund = (xlRange.Cells[i, 9] != null && xlRange.Cells[i, 9].Value2 != null) ? Convert.ToString(xlRange.Cells[i, 9].Value2) : "";
            //    abwesenheit.Text = (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null) ? Convert.ToString(xlRange.Cells[i, 10].Value2) : "";
            //    abwesenheit.Status = (xlRange.Cells[i, 15] != null && xlRange.Cells[i, 15].Value2 != null) ? Convert.ToString(xlRange.Cells[i, 15].Value2) : "";

            //    if (
            //            abwesenheit.Status == "offen" ||
            //            abwesenheit.Status == "nicht entsch.")
            //    {
            //        this.Add(abwesenheit);
            //    }
            //}

            ////cleanup
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            ////close and release
            //xlWorkbook.Close();
            //Marshal.ReleaseComObject(xlWorkbook);

            ////quit and release
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);

            Console.WriteLine(("Abwesenheiten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
        }
        
        internal void Get20StundenIn30Tage()
        {
            throw new NotImplementedException();
        }
    }
}