﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDataReader;
using OfficeOpenXml;

namespace AnalyticalCalculator
{
    public partial class Form1 : Form
    {
        private string mainFolder = @"..\..\..\Company";
        private string allFolder = @"..\..\..\";

        private DataTableCollection tableCollection = null;

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Установка контекста лицензии
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            string path = mainFolder;
            if (Directory.Exists(path))
            {
                var directories = Directory.GetDirectories(path).Select(Path.GetFileName).ToArray();
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
                toolStripComboBox2.Items.Clear();
                var excelFiles = Directory.GetFiles(path, "*.xls*").Select(Path.GetFileName).ToArray();
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
                FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                    {
                        //UseHeaderRow = true
                    }
                });

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

            for (int i = 0; i < dataGridView1.RowCount; i++)
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

        private void LoadExcelFile(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    tableCollection = result.Tables;
                    DataTable table = tableCollection[0];
                    dataGridView1.DataSource = table;
                }
            }
        }

        private Dictionary<string, List<double>> ExtractIndicators(string companyFolder)
        {
            var indicators = new Dictionary<string, List<double>>();

            try
            {
                string incomeStatementPath = Path.Combine(companyFolder, "Income Statement_Annual_As Originally Reported.xls");
                string balanceSheetPath = Path.Combine(companyFolder, "Balance Sheet_Annual_As Originally Reported.xls");
                string cashFlowPath = Path.Combine(companyFolder, "Cash Flow_Annual_As Originally Reported.xls");

                if (!File.Exists(incomeStatementPath) || !File.Exists(balanceSheetPath) || !File.Exists(cashFlowPath))
                {
                    MessageBox.Show($"Для компании {Path.GetFileName(companyFolder)} не хватает одного или нескольких файлов отчетности.");
                    return null;
                }

                indicators["Total Revenue"] = GetValuesFromExcel(incomeStatementPath, "Total Revenue");
                indicators["Net Income"] = GetValuesFromExcel(incomeStatementPath, "Diluted Net Income Available to Common Stockholders");
                indicators["Gross Profit"] = GetValuesFromExcel(incomeStatementPath, "Gross Profit");
                indicators["Operating Expenses"] = GetValuesFromExcel(incomeStatementPath, "Operating Income/Expenses");
                indicators["EPS"] = GetValuesFromExcel(incomeStatementPath, "Diluted EPS");

                indicators["Total Assets"] = GetValuesFromExcel(balanceSheetPath, "Total Assets");
                indicators["Total Liabilities"] = GetValuesFromExcel(balanceSheetPath, "Total Liabilities");
                indicators["Long-term Debt"] = GetValuesFromExcel(balanceSheetPath, "Financial Liabilities, Non-Current");
                indicators["Short-term Debt"] = GetValuesFromExcel(balanceSheetPath, "Financial Liabilities, Current");
                indicators["Total Equity"] = GetValuesFromExcel(balanceSheetPath, "Total Equity");
                indicators["Accounts Receivable"] = GetValuesFromExcel(balanceSheetPath, "Trade/Accounts Receivable, Current");
                indicators["Total Current Liabilities"] = GetValuesFromExcel(balanceSheetPath, "Total Current Liabilities");
                indicators["Total Current Assets"] = GetValuesFromExcel(balanceSheetPath, "Total Current Assets");
                indicators["Inventories"] = GetValuesFromExcel(balanceSheetPath, "Inventories");
                indicators["Cash, Cash Equivalents and Short Term Investments"] = GetValuesFromExcel(balanceSheetPath, "Cash, Cash Equivalents and Short Term Investments");

                indicators["Depreciation, Amortization and Depletion"] = GetValuesFromExcel(cashFlowPath, "Depreciation, Amortization and Depletion, Non-Cash Adjustment");
                indicators["Operating Cash Flow"] = GetValuesFromExcel(cashFlowPath, "Cash Generated from Operating Activities");
                indicators["Capital Expenditures, CapEx"] = GetValuesFromExcel(cashFlowPath, "Purchase/Sale and Disposal of Property, Plant and Equipment, Net");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при извлечении показателей для компании {Path.GetFileName(companyFolder)}: {ex.Message}");
                return null;
            }

            return indicators;
        }

        private List<double> GetValuesFromExcel(string filePath, string indicatorName)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    var table = result.Tables[0];

                    foreach (DataRow row in table.Rows)
                    {
                        if (row[0].ToString().Contains(indicatorName))
                        {
                            List<double> values = new List<double>();
                            for (int i = 1; i <= 5; i++)
                            {
                                if (row[i] == DBNull.Value || string.IsNullOrEmpty(row[i].ToString()))
                                {
                                    values.Add(0); // Если ячейка пустая, записать 0
                                }
                                else
                                {
                                    values.Add(Convert.ToDouble(row[i]));
                                }
                            }
                            return values;
                        }
                    }
                }
            }

            // throw new Exception($"Показатель {indicatorName} не найден в файле {filePath}");
            // Если показатель не найден, вернуть список из 5 нулей
            return new List<double> { 0, 0, 0, 0, 0 };
        }


        private void CreateSummaryFiles()
        {
            string summaryFolder = Path.Combine(allFolder, "SummaryReports");
            if (!Directory.Exists(summaryFolder))
            {
                Directory.CreateDirectory(summaryFolder);
            }

            foreach (string company in toolStripComboBox1.Items)
            {
                string companyFolder = Path.Combine(mainFolder, company);
                var indicators = ExtractIndicators(companyFolder);

                if (indicators != null)
                {
                    string summaryFilePath = Path.Combine(summaryFolder, $"{company}.xlsx");
                    SaveToExcel(summaryFilePath, indicators);
                }
            }
        }

        private void SaveToExcel(string filePath, Dictionary<string, List<double>> indicators)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Summary");

                worksheet.Cells[1, 1].Value = "Indicator";
                worksheet.Cells[1, 2].Value = "2019";
                worksheet.Cells[1, 3].Value = "2020";
                worksheet.Cells[1, 4].Value = "2021";
                worksheet.Cells[1, 5].Value = "2022";
                worksheet.Cells[1, 6].Value = "2023";

                int row = 2;
                foreach (var indicator in indicators)
                {
                    worksheet.Cells[row, 1].Value = indicator.Key;
                    for (int i = 0; i < indicator.Value.Count; i++)
                    {
                        worksheet.Cells[row, i + 2].Value = indicator.Value[i];
                    }
                    row++;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateSummaryFiles();
        }
    }
}
