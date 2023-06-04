using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IPP_Lb1_Brigade6
{
    public partial class MainWindow : Form
    {
        private void StartExcel()
        {
            try
            {
                var excel1Process = Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE");

                while (true)
                {
                    if (excel1Process.HasExited)
                    {
                        dataGridViewExcel.Rows.Clear();
                        dataGridViewExcel.Rows.Add("Процесс завершенно!");
                        break;
                    }
                    else
                    {
                        var totalProcessorTime = DateTime.Now - excel1Process.StartTime;
                        var userProcessorTime = totalProcessorTime - excel1Process.UserProcessorTime;
                        var privilegedProcessorTime = totalProcessorTime - excel1Process.PrivilegedProcessorTime;
                        // Виведення відомостей про час виконання потоку
                        dataGridViewExcel.Rows.Clear();
                        dataGridViewExcel.Rows.Add("ID процесу", excel1Process.Id);
                        dataGridViewExcel.Rows.Add("Загальний час виконання", totalProcessorTime);
                        dataGridViewExcel.Rows.Add("Час виконання частини користувача", userProcessorTime);
                        dataGridViewExcel.Rows.Add("Час виконання частини ядра", privilegedProcessorTime);
                        Thread.Sleep(5000);
                    }
                }
            }
            catch
            {
                dataGridViewExcel.Rows.Clear();
                dataGridViewExcel.Rows.Add("Не можу запустити Excel!");
            }
        }

        private void StartWord()
        {
            try
            {
                var word2Process = Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE");

                while (true)
                {
                    if (word2Process.HasExited)
                    {
                        dataGridViewWord.Rows.Clear();
                        dataGridViewWord.Rows.Add("Процесс завершенно!");
                        break;
                    }
                    else
                    {
                        var totalProcessorTime = DateTime.Now - word2Process.StartTime;
                        var userProcessorTime = totalProcessorTime - word2Process.UserProcessorTime;
                        var privilegedProcessorTime = totalProcessorTime - word2Process.PrivilegedProcessorTime;
                        // Виведення відомостей про час виконання потоку
                        dataGridViewWord.Rows.Clear();
                        dataGridViewWord.Rows.Add("ID процесу", word2Process.Id);
                        dataGridViewWord.Rows.Add("Загальний час виконання", totalProcessorTime);
                        dataGridViewWord.Rows.Add("Час виконання частини користувача", userProcessorTime);
                        dataGridViewWord.Rows.Add("Час виконання частини ядра", privilegedProcessorTime);
                        Thread.Sleep(5000);
                    }
                }
            }
            catch
            {
                dataGridViewWord.Rows.Clear();
                dataGridViewWord.Rows.Add("Не можу запустити Word!");
            }
        }
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            var threadExcel = new Thread(StartExcel);
            var threadWord = new Thread(StartWord);
            threadExcel.Start();
            threadWord.Start();
        }
    }
}