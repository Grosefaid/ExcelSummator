using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelReports
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            radioButton1.Checked = true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            textBox1.Text = openFileDialog1.FileName;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() != DialogResult.OK) return;
            textBox2.Text = openFileDialog2.FileName;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog3.ShowDialog() != DialogResult.OK) return;
            textBox3.Text = openFileDialog3.FileName;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog4.ShowDialog() != DialogResult.OK) return;
            textBox4.Text = openFileDialog4.FileName;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog5.ShowDialog() != DialogResult.OK) return;
            textBox5.Text = openFileDialog5.FileName;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog6.ShowDialog() != DialogResult.OK) return;
            textBox6.Text = openFileDialog6.FileName;
        }

        private void calculate_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();

            Hashtable final_arrays = get_settings();
            bool error_flag = false;

            if (openFileDialog1.FileName != "")
                get_values(openFileDialog1.FileName, xlApp, final_arrays, ref  error_flag);

            if (openFileDialog2.FileName != "")
                get_values(openFileDialog2.FileName, xlApp, final_arrays, ref error_flag);

            if (openFileDialog3.FileName != "")
                get_values(openFileDialog3.FileName, xlApp, final_arrays, ref error_flag);

            if (openFileDialog4.FileName != "")
                get_values(openFileDialog4.FileName, xlApp, final_arrays, ref error_flag);

            if (openFileDialog5.FileName != "")
                get_values(openFileDialog5.FileName, xlApp, final_arrays, ref error_flag);

            if (openFileDialog6.FileName != "")
                set_values(openFileDialog6.FileName, xlApp, final_arrays);

            xlApp.Quit();
            if (!error_flag)
                MessageBox.Show("Готово!", ":)", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void get_values(string filename, Excel.Application xlApp, Hashtable final_arrays, ref bool error_flag)
        {
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filename);
            List<string> keys = new List<string>();
            foreach (System.Collections.DictionaryEntry de in final_arrays)
                keys.Add(de.Key.ToString());

            foreach (string key in keys) {
                string[] parameters = key.Split('-');
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(Int32.Parse(parameters[0]));
                double[,] values = GetValues(xlWorkSheet.get_Range(parameters[1]).Cells.Value2, filename, key, ref error_flag);
                sum_arrays(values, final_arrays, key);
            }

            xlWorkBook.Close(true);
        }

        double[,] GetValues(Object rangeValues, string filename, string key, ref bool error_flag)
        {
            double[,] values = null;

            Array array = rangeValues as Array;
            if (null != array)
            {
                int rank = array.Rank;
                if (rank > 1)
                {
                    int rowCount = array.GetLength(0);
                    int columnCount = array.GetUpperBound(1);

                    values = new double[rowCount, columnCount];

                    for (int index = 0; index < rowCount; index++)
                    {
                        for (int index2 = 0; index2 < columnCount; index2++)
                        {
                            Object obj = array.GetValue(index + 1, index2 + 1);
                            if (null != obj)
                            {
                                try
                                {
                                    values[index, index2] = (double)obj;
                                }
                                catch
                                {
                                    string[] parameters = key.Split('-');
                                    string[] start_range = parameters[1].Split(':');
                                    string start_column = Regex.Match(start_range[0], @"^[A-Za-z]+", RegexOptions.None).Groups[0].Value.ToUpperInvariant();
                                    int sum = 0;

                                    for (int i = 0; i < start_column.Length; i++)
                                    {
                                        sum *= 26;
                                        sum += (start_column[i] - 'A' + 1);
                                    }

                                    int error_column = sum + index2;
                                    string colLetter = String.Empty;
                                    int mod = 0;

                                    while (error_column > 0)
                                    {
                                        mod = (error_column - 1) % 26;
                                        colLetter = (char)(65 + mod) + colLetter;
                                        error_column = (int)((error_column - mod) / 26);
                                    }

                                    MessageBox.Show("Файл: " + filename + " \nНомер листа: " + parameters[0] + " Ячейка: " + colLetter + (int.Parse(start_range[0].Substring(start_column.Length)) + index), 
                                        "Ошибка значения ячейки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    error_flag = true;
                                }
                            }
                        }
                    }
                }
            }

            return values;
        }

        Hashtable get_settings()
        {
            Hashtable settings = new Hashtable();
            string file_name = "";
            if (radioButton1.Checked)
            {
                file_name = "settings/ugolovniy.txt";
            }
            if (radioButton2.Checked)
            {
                file_name = "settings/administrativniy.txt";
            }
            if (radioButton3.Checked)
            {
                file_name = "settings/grazhdanskiy.txt";
            }
            if (radioButton4.Checked)
            {
                file_name = "settings/operativniy.txt";
            }
            if (radioButton5.Checked)
            {
                file_name = "settings/uscherb.txt";
            }
            if (radioButton6.Checked)
            {
                file_name = "settings/uscherb_pilogenie.txt";
            }
            
            string[] fileLines = File.ReadAllLines(file_name);
            foreach (string line in fileLines)
            {
                settings.Add(line, null);
            }
            
            return settings;
        }

        void sum_arrays(double[,] values, Hashtable final_array, string key)
        {
            if (final_array[key] == null)
                final_array[key] = values;
            else
            {
                for (int i = 0; i < values.GetLength(0); i++)
                {
                    for (int j = 0; j < values.GetLength(1); j++)
                    {
                        ((double[,])final_array[key])[i, j] += values[i,j];
                    }
                }
            }
        }

        void set_values(string filename, Excel.Application xlApp, Hashtable final_arrays)
        {
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filename);
            List<string> keys = new List<string>();
            foreach (System.Collections.DictionaryEntry de in final_arrays)
                keys.Add(de.Key.ToString());

            foreach (string key in keys)
            {
                string[] parameters = key.Split('-');
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(Int32.Parse(parameters[0]));
                xlWorkSheet.get_Range(parameters[1]).Formula = final_arrays[key];
            }

            xlWorkBook.Close(true);
        }
    }
}
