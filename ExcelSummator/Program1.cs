using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace combine
{
    class Program
    {
        static void Main(string[] args)
        {
            var outDir = Path.GetFullPath("out");
            try
            {
                if (!Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("не смог создать папку куда записывать полчившийся файл " + ex);
                finish();
                return;
            }
            initLog(outDir);
            log("программа стартовала ", Priority.High);

            try
            {
                run(outDir);
            }
            catch (Exception ex)
            {
                log("программа завершила работу с ошибкой " + ex, Priority.High);
                finish();
            }
        }

        private static void run(string outDir)
        {
            // search template
            var inFile = Path.GetFullPath("Template.xls");
            if (!File.Exists(inFile))
            {
                log("не найден файл шаблона " + inFile, Priority.High);
                finish();
                return;
            }
            report("шаблон " + inFile);

            // search data files
            var inDir = Path.GetFullPath("in");
            if (!Directory.Exists(inDir))
            {
                Directory.CreateDirectory(inDir);
            }
            var files = Directory.GetFiles(inDir).Select(Path.GetFullPath).ToArray();
            report("количество файлов в папке " + inDir + " равно " + files.Length);
            if (files.Length == 0)
            {
                finish();
                return;
            }

            _excel = new Application();
            _excel.EnableEvents = false;
            _template = _excel.Workbooks.Open(inFile);
            _filledWorkbooks = new List<Workbook>(files.Length);
            _filledWorkbooks.AddRange(files.Select(file =>
                                                      {
                                                          try
                                                          {
                                                              Workbook workBook = _excel.Workbooks.Open(file);
                                                              // validate in file
                                                              if (workBook.Sheets.Count != _template.Sheets.Count )
                                                              {
                                                                  log("количесво листов в документе " + file + " равно " + workBook.Sheets.Count, Priority.High);
                                                                  return null;
                                                              }
                                                              return workBook;
                                                          }
                                                          catch (Exception ex)
                                                          {
                                                              logError(file, ex, "Не смог прочитать файл", Priority.High);
                                                              return null;
                                                          }
                                                          ;
                                                      }
                                         ).Where(x => x != null).ToArray());
            report("количество соответсвующих шаблону документов Excel равно " + _filledWorkbooks.Count);
            if (_filledWorkbooks.Count == 0)
            {

                finish();
                return;
            }
            int errorCounter = 0;
            int emptyCounter = 0;
            foreach (Worksheet sheet in _template.Sheets)
            {
                report(sheet.Name);
                var lastCell = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                var r = lastCell.Row;
                var c = sheet.Columns.Count;

                for (int i = 1; i <= r; i++)
                {
                    for (int j = 1; j <= c; j++)
                    {
                        Range cell = sheet.Cells[i, j];
                        int colorNumber = System.Convert.ToInt32(cell.Interior.Color);
                        Color color = System.Drawing.ColorTranslator.FromOle(colorNumber);
                        if (color != Color.White) //cell is gray
                        {
                            decimal summ = 0;
                            foreach (var filled in _filledWorkbooks)
                            {
                                var filledSheet = filled.Sheets[sheet.Name] as Worksheet;
                                var filledCell = filledSheet.Cells[i, j] as Range;
                                var value = filledCell.Value2;
                                string message = null;
                                decimal add = 0;
                                Priority p = Priority.Low;

                                if (value == null)
                                {
                                    emptyCounter++;
                                }
                                else
                                {
                                    string repr = string.Format("{0}", value);
                                    if (String.IsNullOrWhiteSpace(repr))
                                    {
                                        message = "Ячейка пробел";
                                        errorCounter++;
                                        p = Priority.High;
                                    }
                                    else
                                    {
                                        if (!Decimal.TryParse(repr, out add))
                                        {
                                            message = "Нечисленное значение";
                                            errorCounter++;
                                            p = Priority.High;
                                        }
                                    }
                                }
                                if (message != null)
                                {
                                    log(filled.FullName, filledSheet.Name, filledCell.Row, filledCell.Column, message, p, Severity.Error);
                                }
                                summ += add;
                            }
                            cell.Value2 = summ;
                        }
                    }
                }
            }
            report("ошибок " + errorCounter);
            report("пустых цветных ячеек" + emptyCounter);
            if (!Directory.Exists(inDir))
            {
                Directory.CreateDirectory(inDir);
            }
            var outFile = Path.Combine(outDir, "all.xls");
            _template.SaveAs(outFile);
            _template.Close(false);
            report("сделано " + outFile);
            report("нажмите любую клавишу для завершения");
            finish();
        }

        private static void finish()
        {
            finishInternal();
            Console.ReadKey();
        }

        private static void finishInternal()
        {
            if (_template != null)
            {
                Marshal.FinalReleaseComObject(_template);
                _template = null;
            }
            if (_filledWorkbooks != null)
            {
                foreach (Workbook filledWorkbook in _filledWorkbooks)
                {
                    filledWorkbook.Close(false);
                    Marshal.FinalReleaseComObject(filledWorkbook);
                }
                _filledWorkbooks = null;
            }

            if (_excel != null)
            {
                _excel.Quit();
                Marshal.FinalReleaseComObject(_excel);
                _excel = null;
            }
            _writer.Flush();
        }

        private static void initLog(string outDir)
        {
            def = Priority.Mid;
            _writer = new StreamWriter(Path.Combine(outDir, "log.txt"), false);
            AppDomain.CurrentDomain.DomainUnload += new EventHandler(CurrentDomain_DomainUnload);
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_DomainUnload);
        }

        static void CurrentDomain_DomainUnload(object sender, EventArgs e)
        {
            finishInternal();
        }

        private static void report(string message)
        {
            Console.WriteLine(message);
            log(message, Priority.Mid);
        }
        
        private static void log(string message, Priority p,Severity s = Severity.Info)
        {
            if (p >= def)
            {
                _writer.WriteLine(DateTime.Now + "------------------------------------------------------");
                _writer.WriteLine(message);
            }
            if (p == Priority.High)
            {
                var fore = Console.ForegroundColor;
                var back = Console.BackgroundColor;
                if (s == Severity.Error)
                {
                    if (back == ConsoleColor.Black) Console.ForegroundColor = ConsoleColor.Red;
                }
               Console.WriteLine(message);
               Console.ForegroundColor = fore;
                
                
            }
        }

        private static void logError(string file, Exception exception, string message, Priority p)
        {
            var formatted = string.Format("файл {0}; ошибка {1}; сообщение {2} ", file, exception.Message, message);
            log(formatted, p);
        }

        private static Priority def;
        private static StreamWriter _writer;
        private static Application _excel;
        private static Workbook _template;
        private static List<Workbook> _filledWorkbooks;

        private static void log(string fullName, string name, int row, int column, string message, Priority p, Severity s = Severity.Info)
        {
            string columnName = GetExcelColumnName(column);
            var formatted = string.Format("файл {0}; лист {1}; ячейка {2},{3} - " + message, fullName, name, row, columnName);
            log(formatted, p,s);
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
