using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDataReader;

namespace AnalyticalCalculator
{
    public partial class Form1 : Form
    {
        private string mainFolder = @"..\..\..\Company";

        private DataTableCollection tableCollection = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            string path = mainFolder;
            if (Directory.Exists(path))
            {
                // Получаем названия из папки Company
                var directories = Directory.GetDirectories(path).Select(Path.GetFileName).ToArray();

                // Очищаем toolStripComboBox1 и записываем туда названия всех папок в папке Company
                toolStripComboBox1.Items.Clear();
                toolStripComboBox1.Items.AddRange(directories);
            }
            else
            {
                MessageBox.Show($"Папка не найдена по пути: {path}");
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedFolder = toolStripComboBox1.SelectedItem.ToString();
            string path = Path.Combine(mainFolder, selectedFolder);

            if (Directory.Exists(path))
            {
                // Очищаем toolStripComboBox2 перед добавлением новых элементов
                toolStripComboBox2.Items.Clear();

                // Получаем все файлы Excel в выбранной папке
                var excelFiles = Directory.GetFiles(path, "*.xls*").Select(Path.GetFileName).ToArray();

                // Добавляем файлы в toolStripComboBox2
                toolStripComboBox2.Items.AddRange(excelFiles);
            }
            else
            {
                MessageBox.Show($"Папка не найдена по пути: {path}");
            }
        }

        private void toolStripComboBox2_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedFolder = toolStripComboBox1.SelectedItem.ToString();
            string selectedFile = toolStripComboBox2.SelectedItem.ToString();
            string path = Path.Combine(mainFolder, selectedFolder, selectedFile);

            if (File.Exists(path))
            {
                //в этом потоке мы открываем нужный нам файл
                FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                //нам нужно привести то значение которое нам вернет CreateReader к интерфейсу IExcelDataReader
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                //создаем базу данных в конструкции которой мы настраиваем как она будет считана
                DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                    {
                        //UseHeaderRow = true
                    }
                });

                //отображаем содержимое таблицы

                tableCollection = db.Tables;

                DataTable table = tableCollection[0];

                dataGridView1.DataSource = table;
            }
            else
            {
                MessageBox.Show($"Папка не найдена по пути: {path}");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {                           
            string word = textBox1.Text;
            int count = 0;

            for (int i = 0; i < dataGridView1.RowCount; i++)                 //покраска в белый
                for (int j = 0; j < dataGridView1.ColumnCount; j++)                                   
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.White;
            

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {    
                    
                    if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(word) && word != "")
                    {                        
                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.GreenYellow;
                        count++;
                    }                    
                }
            }

            textBox2.Text = count.ToString();
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            label1.Visible = true;
            textBox2.Visible = true;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            label1.Visible = false;
            textBox2.Visible = false;
        }
        
    }
}
