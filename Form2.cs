using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace AnalyticalCalculator
{
    public partial class Form2 : Form
    {
        //public static Form2 SelfRef { get; set; }
        private Dictionary<string, Dictionary<string, List<double>>> allData;

        public Form2(Dictionary<string, Dictionary<string, List<double>>> data)
        {
            //SelfRef = this;
            allData = data;
            InitializeComponent();
        }

        // Метод для построения графика
        public void PlotCharts(string metric)
        {
            chart1.Series.Clear();
            chart1.ChartAreas.Clear();

            var chartArea = new ChartArea();
            chart1.ChartAreas.Add(chartArea);

            foreach (var companyData in allData)
            {
                var companyName = companyData.Key;
                if (!companyData.Value.ContainsKey(metric))
                {
                    continue;
                }

                var metricValues = companyData.Value[metric];

                var series = new Series(companyName)
                {
                    ChartType = SeriesChartType.Line,
                    BorderWidth = 2
                };

                for (int i = 0; i < metricValues.Count; i++)
                {
                    series.Points.AddXY(i + 1, metricValues[i]); // X value can be adjusted as needed (e.g., year)
                }

                chart1.Series.Add(series);
            }

            chart1.Invalidate();
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            toolStripComboBox1.Items.Clear();
            toolStripComboBox1.Items.AddRange(new string[]
            {
                "Total Revenue", "Net Income", "EBITDA", "EPS", "Free Cash Flow",
                "ROA", "ROE", "ROS", "Current Ratio", "Quick Ratio",
                "Debt-to-Equity Ratio", "P/E", "EV/EBITDA", "P/S", "P/BV", "EV/S", "Net Debt"
            });
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedMetric = toolStripComboBox1.SelectedItem.ToString();
            PlotCharts(selectedMetric);
        }
    }
}
