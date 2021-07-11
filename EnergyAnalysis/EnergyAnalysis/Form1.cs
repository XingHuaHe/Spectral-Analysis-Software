using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

using Excel = Microsoft.Office.Interop.Excel;
using EnergyAnalysis.Models;
using EnergyAnalysis.Services;

namespace EnergyAnalysis
{
    public partial class Form1 : Form
    {
        //保存纯元素字典
        Dictionary<string, Element> pureElement = new Dictionary<string, Element>();
        //保存待分析元素字典
        Dictionary<string, AnalysisElement> analysisElement = new Dictionary<string, AnalysisElement>();
        //保存已绘制能谱字典
        Dictionary<string, AnalysisElement> chartElements = new Dictionary<string, AnalysisElement>();
        //保存标定元素字典
        Dictionary<string, CalibrationElement> calibrationElement = new Dictionary<string, CalibrationElement>();
        //a为斜率,b为截距
        private double a = 0, b = 0;
        //保存已经绘制的标签曲线名称
        private List<string> SerierName = new List<string>();
        //通道转换为能量的标准数据
        private double Slope = 0.033929229;
        private double intercept = -0.10118409;
        //初始化完成标志（Flase为未完成，True为完成）
        private bool FinishInitFlag = false;

        public Form1()
        {
            FinishInitFlag = false;
            InitializeComponent();

            InitChart();
            InitCalibration();//初始化元素标定窗口，因为需要在后端打开excel读取能量数据表，故因电脑按照office不同，可能导致报错
            InitListView2();
            //InitListView3();
            this.chart1.GetToolTipText += Chart1_GetToolTipText;
            FinishInitFlag = true;
        }


        private void Chart1_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            if(e.HitTestResult.ChartElementType == ChartElementType.DataPoint)
            {
                int i = e.HitTestResult.PointIndex;
                DataPoint dp = e.HitTestResult.Series.Points[i];
                e.Text = string.Format("{0:F3}  {1:F3}", dp.XValue, dp.YValues[0]);
            }
        }

        /// <summary>
        /// 初始化ListView3
        /// </summary>
        public void InitListView3()
        {
            string[] ways = { "一阶指数滤波", "二阶指数滤波", "三阶指数滤波", "平滑滤波" };
            comboBox2.Items.Clear();
            foreach (var way in ways)
            {
                comboBox2.Items.Add(way);
            }
            comboBox2.Text = "一阶指数滤波";
            comboBox3.Items.Clear();
            foreach (var way in this.chart1.Series)
            {
                comboBox3.Items.Add(way);
            }
            textBox8.Text = "5";
        }

        /// <summary>
        /// 初始化ListView2
        /// </summary>
        public void InitListView2()
        {
            listView2.MultiSelect = true;
            string[] items = { "通道分析", "能量分析" };
            foreach (var item in items)
            {
                comboBox1.Items.Add(item);
            }
            comboBox1.Text = "通道分析";
        }

        /// <summary>
        /// 初始化绘制框
        /// </summary>
        public void InitChart()
        {
            
            string label = "Energy";
            this.chart1.ChartAreas.Clear();
            this.chart1.Series.Clear();

            ChartArea chartarea = new ChartArea(label);
            chartarea.AxisX.Minimum = 0;
            chartarea.AxisX.Maximum = 2048;
            chartarea.AxisY.Minimum = 0;
            chartarea.AxisY.Maximum = 2048;
            chartarea.AxisX.MinorGrid.Enabled = true;
            chartarea.AxisY.MajorGrid.Enabled = true;
            this.chart1.ChartAreas.Add(chartarea);

            Series series = new Series("S1");
            series.ChartType = SeriesChartType.Spline;
            series.ChartArea = label;
            series.Points.AddXY(0, 0);
            series.Points.AddXY(2048, 2048);
            this.chart1.Series.Add(series);
        }

        /// <summary>
        /// 初始化元素标定(ListView1)
        /// </summary>
        public void InitCalibration()
        {
            string file_path = string.Format(@"{0}\..\..\Labels.xlsx", Application.StartupPath);//获取当前启动文件的路径，并读取上二层的文件。其中..\表示返回上一层，注意要在当前路径下\，即{0}\
            //string file_path = "D:\\Program Files\\EnergyAnalysis\\EnergyAnalysis\\Labels.xlsx";
            //listView1.MultiSelect = true;

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = null;
            excel.Visible = false;
            wb = excel.Workbooks.Open(file_path);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
            int rowCount = 0;
            try
            {
                rowCount = ws.UsedRange.Rows.Count;//有效行
                for(int i = 2; i <= rowCount; i++)
                {
                    CalibrationElement calElement = new CalibrationElement();
                    if (ws.Rows[i] != null)
                    {
                        //若该行数据不为空，则读取每列数据
                        calElement.id = ws.Cells[i, 1].Value2.ToString();
                        calElement.symbol = ws.Cells[i, 2].Value2.ToString();
                        calElement.name = ws.Cells[i, 3].Value2.ToString();
                        calElement.Ka = ws.Cells[i, 4].Value2.ToString();
                        calElement.Kb = ws.Cells[i, 5].Value2.ToString();
                        calElement.La = ws.Cells[i, 6].Value2.ToString();
                        calElement.Lb = ws.Cells[i, 7].Value2.ToString();
                        calElement.Lg = ws.Cells[i, 8].Value2.ToString();
                        calElement.Ll = ws.Cells[i, 9].Value2.ToString();
                    }
                    //将数据写到ListView3
                    ListViewItem item = new ListViewItem();
                    item.Text = (i - 1).ToString();
                    item.SubItems.Add(calElement.name);
                    item.SubItems.Add(calElement.id);
                    listView1.Items.Add(item);
                    calibrationElement.Add(calElement.id, calElement);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //关闭进程
                Process[] localByNameApp = Process.GetProcessesByName(file_path);//获取程序名的所有进程
                if (localByNameApp.Length > 0)
                {
                    foreach (var app in localByNameApp)
                    {
                        if (!app.HasExited)
                        {
                            app.Kill();
                        }
                    }
                }
                if (wb != null)
                {
                    wb.Close(true, Type.Missing, Type.Missing);
                }
                excel.Quit();
                System.GC.GetGeneration(excel);
            }
        }

        /// <summary>
        /// ListView1的确认按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //删除原有的标定垂直曲线
            foreach (string name in SerierName)
            {
                var a = this.chart1.Series.IndexOf(name);
                this.chart1.Series.RemoveAt(a);
            }
            SerierName.Clear();
            //添加曲线
            if(a != 0 && b != 0)
            {
                if (comboBox1.Text == "能量分析")
                {
                    foreach (ListViewItem item in listView1.CheckedItems)
                    {
                        CalibrationElement element = new CalibrationElement();
                        AnalysisElement analyElement = new AnalysisElement();
                        List<float> energy = new List<float>();
                        string Ka = string.Empty;
                        string key = item.SubItems[2].Text;//获取元素序号
                        calibrationElement.TryGetValue(key, out element);//取出标定元素数据
                        Ka = element.Ka;

                        foreach (ListViewItem it in listView3.SelectedItems)
                        {
                            string id = it.SubItems[4].Text;
                            analysisElement.TryGetValue(id, out analyElement);
                            energy = analyElement.energy;
                            string label = item.SubItems[1].Text;
                            Series series = new Series(label);
                            series.IsVisibleInLegend = false;
                            series.ChartType = SeriesChartType.Spline;

                            double x = Convert.ToDouble(Ka);//能量分析坐标
                            int xx = (int)((x - b) / a);//通道分析坐标
                                                        //int xx = (int)((x - intercept) / Slope);
                            float y_peak = energy[xx];
                            series.Points.AddXY(x, 0);
                            series.Points.AddXY(x, y_peak);
                            this.chart1.Series.Add(series);
                            SerierName.Add(label);
                        }
                    }
                }
                else if (comboBox1.Text == "通道分析")
                {
                    foreach (ListViewItem item in listView1.CheckedItems)
                    {
                        CalibrationElement element = new CalibrationElement();
                        AnalysisElement analyElement = new AnalysisElement();
                        List<float> energy = new List<float>();
                        string Ka = string.Empty;
                        string key = item.SubItems[2].Text;
                        calibrationElement.TryGetValue(key, out element);//取出标定元素数据
                        Ka = element.Ka;

                        foreach (ListViewItem it in listView3.SelectedItems)
                        {
                            string id = it.SubItems[4].Text;
                            analysisElement.TryGetValue(id, out analyElement);
                            energy = analyElement.energy;
                            string label = item.SubItems[1].Text;//曲线集 利用元素名称作为标识
                            Series series = new Series(label);
                            series.IsVisibleInLegend = false;
                            series.ChartType = SeriesChartType.Spline;

                            double x = Convert.ToDouble(Ka);//能量分析坐标
                            int xx = (int)((x - b) / a);//通道分析坐标
                                                        //int xx = (int)((x - intercept) / Slope) + 1;
                            float y_peak = energy[xx];
                            series.Points.AddXY(xx, 0);
                            series.Points.AddXY(xx, y_peak);
                            this.chart1.Series.Add(series);
                            SerierName.Add(label);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("请先通过纯元素计算斜率、截距等参数！", "提示");
            }
        }

        /// <summary>
        /// ListView1的清除按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            foreach(ListViewItem item in listView1.CheckedItems)
            {
                item.Checked = false;
                var a = this.chart1.Series.IndexOf(item.SubItems[1].Text);
                this.chart1.Series.RemoveAt(a);
            }
            SerierName.Clear();
        }

        /// <summary>
        /// ListView2的打开文件按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            string[] files = openFileDialog.FileNames;
            int c = 1;
            foreach(string file in files)
            {
                FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read);
                BinaryReader binaryReader = new BinaryReader(fileStream);
                byte[] energyData = binaryReader.ReadBytes(6166);
                binaryReader.Close();
                fileStream.Close();

                int i = 5;
                float sum = 0;
                List<float> energy = new List<float>();
                Element element = new Element();
                for(int j = 1; j <=2048 ; j++)
                {
                    sum = Convert.ToInt32(energyData[i + 2]) + Convert.ToInt32(energyData[i + 1]) * 256 + Convert.ToInt32(energyData[i]) * 256 * 256;
                    energy.Add(sum);
                    sum = 0;
                    i += 3;
                }
                element.energy = energy;
                element.id = (energyData[3] * 256 + energyData[4]).ToString();
                element.dieTime = ((energyData[6160] + 256 * energyData[6161]) / 10).ToString();
                element.probeTemperature = ((energyData[6158] * 256 + energyData[6159]) * 0.1 - 273.15).ToString();
                element.batteryVoltage = (energyData[6162] + 256 * energyData[6163]).ToString();
                element.collectionTime = ((energyData[6150] + energyData[6151] * 256 + energyData[6152] * 256 * 256 + energyData[6153] * 256 * 256 * 256) * 0.001).ToString();

                try
                {
                    List<string> ls = new List<string>();
                    ls = pureElement.Keys.ToList();
                    bool flag = ls.Contains(element.id);
                    if (!flag)
                    {
                        pureElement.Add(element.id, element);//将数据保存在字典中

                        //将数据显示到窗体中
                        try
                        {
                            ListViewItem item = new ListViewItem();
                            CalibrationElement CElement = new CalibrationElement();
                            calibrationElement.TryGetValue(element.id, out CElement);
                            item.Text = c.ToString();
                            item.SubItems.Add(CElement.name);
                            item.SubItems.Add(CElement.id);
                            item.SubItems.Add(CElement.Ka);
                            //item.SubItems.Add((energy.IndexOf(energy.Max()) + 1).ToString());
                            item.SubItems.Add((energy.IndexOf(energy.Max())).ToString());
                            listView2.Items.Add(item);
                            c++;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("警告", "导入的文件不属于存元素文件！");
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("错误", "数据导入出错！！");
                }
            }
        }

        /// <summary>
        /// ListView2清除按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            listView2.Items.Clear();
            pureElement.Clear();
        }

        /// <summary>
        /// ListView2删除按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                listView2.Items.Remove(item);
                pureElement.Remove(item.SubItems[2].Text);
            }
        }

        /// <summary>
        /// ListView2计算按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button13_Click(object sender, EventArgs e)
        {
            List<Element> element_dict = new List<Element>();
            string id = string.Empty;
            List<double> x = new List<double>();
            List<double> y = new List<double>();
            if (listView2.SelectedItems.Count == 2)
            {
                //只有两个纯元素
                foreach (ListViewItem item in listView2.SelectedItems)
                {
                    Element element = new Element();
                    CalibrationElement cal_element = new CalibrationElement();
                    List<float> energy = new List<float>();

                    id = item.SubItems[2].Text;
                    pureElement.TryGetValue(id, out element);
                    calibrationElement.TryGetValue(id, out cal_element);

                    energy = element.energy;
                    x.Add(energy.IndexOf(energy.Max()));
                    y.Add(Convert.ToDouble(cal_element.Ka));
                }
                a = (y[0] - y[1]) / (x[0] - x[1]);
                b = y[0] - a * x[0];
                textBox1.Text = a.ToString();
                textBox2.Text = b.ToString();
            }
        }

        /// <summary>
        /// ListView2绘制按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                Element element = new Element();
                List<float> energy = new List<float>();
                string id = item.SubItems[2].Text;
                pureElement.TryGetValue(id, out element);
                energy = element.energy;

                //绘制图形
                this.chart1.Series.Clear();
                Series series = new Series(item.SubItems[1].Text);
                series.ChartType = SeriesChartType.Spline;
                series.ChartArea = "Energy";
                this.chart1.ChartAreas[0].AxisY.Maximum = energy.Max() + 10000;
                for (int i = 0; i < 2048; i++)
                {
                    series.Points.AddXY(i + 1, energy[i]);
                }
                this.chart1.Series.Add(series);
            }
        }

        /// <summary>
        /// ListView2分析方式选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "通道分析" && FinishInitFlag == true)
            {
                SerierName.Clear();//将先前标定的曲线记录清楚

                this.chart1.Series.Clear();
                var keys = chartElements.Keys.ToArray<string>();
                for (int i = 0; i < chartElements.Count; i++)
                {
                    AnalysisElement energy = new AnalysisElement();
                    chartElements.TryGetValue(keys[i], out energy);
                    List<float> Y = new List<float>();
                    Y = energy.energy;
                    this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max();
                    if (this.chart1.ChartAreas[0].AxisY.Maximum < Y.Max())
                    {
                        this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max();
                    }
                    this.chart1.ChartAreas[0].AxisY.Maximum = this.chart1.ChartAreas[0].AxisY.Maximum + 10000;
                    Series series = new Series(energy.id);
                    series.ChartType = SeriesChartType.Spline;
                    series.ChartArea = "Energy";
                    //series.Label = "#VAL";
                    //series.ToolTip = "#VALX:#VAL";
                    double x = 0;
                    for (int j = 1; j <= 2048; j++)
                    {
                        series.Points.AddXY(x, Y[j - 1]);
                        x++;
                    }
                    this.chart1.ChartAreas[0].AxisX.Maximum = x;
                    this.chart1.Series.Add(series);
                }
                InitListView3();
            }
            else if (comboBox1.Text == "能量分析" && FinishInitFlag == true)
            {
                SerierName.Clear();//将先前标定的曲线记录清楚

                this.chart1.Series.Clear();
                var keys = chartElements.Keys.ToArray<string>();
                for (int i = 0; i < chartElements.Count; i++)
                {
                    AnalysisElement energy = new AnalysisElement();
                    chartElements.TryGetValue(keys[i], out energy);
                    List<float> Y = new List<float>();
                    Y = energy.energy;
                    this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max();
                    if (this.chart1.ChartAreas[0].AxisY.Maximum < Y.Max())
                    {
                        this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max();
                    }
                    this.chart1.ChartAreas[0].AxisY.Maximum = this.chart1.ChartAreas[0].AxisY.Maximum + 10000;
                    Series series = new Series(energy.id);
                    series.ChartType = SeriesChartType.Spline;
                    series.ChartArea = "Energy";
                    //series.Label = "#VAL";
                    //series.ToolTip = "#VALX:#VAL";
                    double x = 0;
                    for (int j = 1; j <= 2048; j++)
                    {
                        x = a * j + b;
                        series.Points.AddXY(x, Y[j - 1]);
                    }
                    this.chart1.ChartAreas[0].AxisX.Maximum = x;
                    this.chart1.Series.Add(series);
                }
                InitListView3();
            }
        }

        /// <summary>
        /// ListView2修改按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            string energy = textBox6.Text;
            string peak = textBox7.Text;

        }

        /// <summary>
        /// ListView3打开文件按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            string[] files = openFileDialog.FileNames;
            int count = 1;
            foreach(string file in files)
            {
                FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read);
                BinaryReader binaryReader = new BinaryReader(fileStream);
                byte[] energyData = binaryReader.ReadBytes(6166);
                binaryReader.Close();
                fileStream.Close();

                int i = 5;
                float sum = 0;
                List<float> energy = new List<float>();
                AnalysisElement element = new AnalysisElement();
                for (int j = 1; j <= 2048; j++)
                {
                    sum = Convert.ToInt32(energyData[i + 2]) + Convert.ToInt32(energyData[i + 1]) * 256 + Convert.ToInt32(energyData[i]) * 256 * 256;
                    energy.Add(sum);
                    sum = 0;
                    i += 3;
                }
                element.energy = energy;
                string[] names = file.Split('\\');
                element.fileName = names[names.Length - 1];
                element.id = (energyData[3] * 256 + energyData[4]).ToString();
                element.dieTime = ((energyData[6160] + 256 * energyData[6161]) / 10).ToString();
                element.probeTemperature = ((energyData[6158] * 256 + energyData[6159]) * 0.1 - 273.15).ToString();
                element.batteryVoltage = (energyData[6162] + 256 * energyData[6163]).ToString();
                element.collectionTime = ((energyData[6150] + energyData[6151] * 256 + energyData[6152] * 256 * 256 + energyData[6153] * 256 * 256 * 256) * 0.001).ToString();
                analysisElement.Add(element.id, element);//将数据保存在字典中

                //将数据显示到窗体中
                ListViewItem item = new ListViewItem();
                item.Text = count.ToString();
                item.SubItems.Add(element.id);
                item.SubItems.Add(count.ToString());
                item.SubItems.Add(element.fileName);
                item.SubItems.Add(element.id);
                item.SubItems.Add("0");
                item.SubItems.Add("2048");
                item.SubItems.Add((energy.IndexOf(energy.Max()) + 1).ToString());
                listView3.Items.Add(item);
            }
        }

        /// <summary>
        /// ListView3删除按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            foreach(ListViewItem item in listView3.SelectedItems)
            {
                listView3.Items.Remove(item);
                analysisElement.Remove(item.SubItems[1].Text);

                var index = this.chart1.Series.IndexOf(item.SubItems[1].Text);
                if (index >= 0)
                {
                    //如果没有进行绘图，则Series的数量为0，此时a=-1
                    chartElements.Remove(item.SubItems[1].Text);
                    this.chart1.Series.RemoveAt(index);
                }
            }
            if (chartElements.Count == 0)
            {
                InitChart();
            }
            //重置滤波曲线选择
            comboBox3.Text = "";
            InitListView3();
        }

        /// <summary>
        /// ListView3清除按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            listView3.Items.Clear();
            analysisElement.Clear();
            chartElements.Clear();
            comboBox3.Items.Clear();
            comboBox3.Text = "";
            this.chart1.ChartAreas.Clear();
            this.chart1.Series.Clear();
            InitChart();
        }

        /// <summary>
        /// ListView3绘制图形按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            if (listView3.SelectedItems.Count != 0)
            {
                this.chart1.ChartAreas.Clear();
                chartElements.Clear();
                this.chart1.Series.Clear();

                //绘图区域
                ChartArea chartarea = new ChartArea("Energy");
                chartarea.AxisX.Minimum = 0;
                chartarea.AxisX.Maximum = 2048;
                chartarea.AxisY.Minimum = 0;
                //chartarea.AxisY.Maximum = energy.Max();
                chartarea.AxisX2.MinorGrid.Enabled = true;
                chartarea.AxisY2.MajorGrid.Enabled = true;
                chartarea.AxisX2.ScaleView.Zoomable = true;
                chartarea.AxisY2.ScaleView.Zoomable = true;
                this.chart1.ChartAreas.Add(chartarea);

                //this.chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
                //this.chart1.ChartAreas[0].CursorY.IsUserEnabled = true;
                this.chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                this.chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
                this.chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                this.chart1.ChartAreas[0].AxisY.ScaleView.Zoomable = true;

                bool flag = false;
                foreach (ListViewItem item in listView3.SelectedItems)
                {
                    AnalysisElement energy = new AnalysisElement();
                    analysisElement.TryGetValue(item.SubItems[1].Text, out energy);
                    chartElements.Add(item.SubItems[1].Text, energy);
                    List<float> Y = new List<float>();
                    Y = energy.energy;
                    if (!flag)
                    {
                        this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max() + 10000;
                        flag = true;
                    }
                    else
                    {
                        if (this.chart1.ChartAreas[0].AxisY.Maximum < (Y.Max() + 10000))
                        {
                            this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max() + 10000;
                        }
                    }
                    Series series = new Series(energy.id);
                    series.ChartType = SeriesChartType.Spline;
                    series.ChartArea = "Energy";
                    
                    //series.Label = "#VAL";
                    //series.ToolTip = "#VALX:#VAL";
                    for (int i = 0; i < 2048; i++)
                    {
                        series.Points.AddXY(i+1, Y[i]);
                    }
                    this.chart1.Series.Add(series);
                }
                InitListView3();
            }
            else if (listView3.CheckedItems.Count != 0)
            {
                this.chart1.ChartAreas.Clear();
                this.chart1.Series.Clear();

                //绘图区域
                ChartArea chartarea = new ChartArea("Energy");
                chartarea.AxisX.Minimum = 0;
                chartarea.AxisX.Maximum = 2048;
                chartarea.AxisY.Minimum = 0;
                //chartarea.AxisY.Maximum = energy.Max();
                chartarea.AxisX2.MinorGrid.Enabled = true;
                chartarea.AxisY2.MajorGrid.Enabled = true;
                chartarea.AxisX2.ScaleView.Zoomable = true;
                chartarea.AxisY2.ScaleView.Zoomable = true;
                this.chart1.ChartAreas.Add(chartarea);

                bool flag = false;
                foreach (ListViewItem item in listView3.CheckedItems)
                {
                    AnalysisElement energy = new AnalysisElement();
                    analysisElement.TryGetValue(item.SubItems[1].Text, out energy);
                    List<float> Y = new List<float>();
                    Y = energy.energy;
                    if (!flag)
                    {
                        this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max() + 10000;
                        flag = true;
                    }
                    else
                    {
                        if (this.chart1.ChartAreas[0].AxisY.Maximum < (Y.Max() + 10000))
                        {
                            this.chart1.ChartAreas[0].AxisY.Maximum = Y.Max() + 10000;
                        }
                    }
                    Series series = new Series(energy.id);
                    series.ChartType = SeriesChartType.Spline;
                    series.ChartArea = "Energy";
                    //series.Label = "#VAL";
                    //series.ToolTip = "#VALX:#VAL";
                    for (int i = 0; i < 2048; i++)
                    {
                        series.Points.AddXY(i+1, Y[i]);
                    }
                    this.chart1.Series.Add(series);
                }
                InitListView3();
            }
        }

        /// <summary>
        /// ListView3应用按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            //利用已经显示在Chart1上的曲线数据进行重新绘制
            string AlgType = comboBox2.Text;
            string series = comboBox3.Text;
            int n = Convert.ToInt32(textBox8.Text);
            if(comboBox1.Text == "通道分析")
            {
                if(series != "")
                {
                    if(AlgType == "一阶指数滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        chartElements.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.primary_Exponential(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        //double x = 0;
                        for (int i = 0; i < 2048; i++)
                        {
                            //x = a * i + b;
                            new_series.Points.AddXY(i, new_datas[i]);
                        }
                        //this.chart1.ChartAreas[0].AxisX.Maximum = x;
                        this.chart1.Series.Add(new_series);
                    }
                    else if(AlgType == "二阶指数滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        analysisElement.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.quadratic_Exponential(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        for (int i = 0; i < 2048; i++)
                        {
                            new_series.Points.AddXY(i, new_datas[i]);
                        }
                        this.chart1.Series.Add(new_series);
                    }
                    else if(AlgType == "三阶指数滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        analysisElement.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.cubic_Exponential(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        for (int i = 0; i < 2048; i++)
                        {
                            new_series.Points.AddXY(i, new_datas[i]);
                        }
                        this.chart1.Series.Add(new_series);
                    }
                    else if(AlgType == "平滑滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        analysisElement.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.mobile_Smoothing(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        for (int i = 0; i < 2048; i++)
                        {
                            new_series.Points.AddXY(i, new_datas[i]);
                        }
                        this.chart1.Series.Add(new_series);
                    }
                }
            }
            else if(comboBox1.Text == "能量分析")
            {
                if (series != "")
                {
                    if (AlgType == "一阶指数滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        chartElements.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.primary_Exponential(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        double x = 0;
                        for (int i = 0; i < 2048; i++)
                        {
                            x = a * i + b;
                            new_series.Points.AddXY(x, new_datas[i]);
                        }
                        this.chart1.ChartAreas[0].AxisX.Maximum = x;
                        this.chart1.Series.Add(new_series);
                    }
                    else if (AlgType == "二阶指数滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        chartElements.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.quadratic_Exponential(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        double x = 0;
                        for (int i = 0; i < 2048; i++)
                        {
                            x = a * i + b;
                            new_series.Points.AddXY(x, new_datas[i]);
                        }
                        this.chart1.ChartAreas[0].AxisX.Maximum = x;
                        this.chart1.Series.Add(new_series);
                    }
                    else if (AlgType == "三阶指数滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        chartElements.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.cubic_Exponential(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        double x = 0;
                        for (int i = 0; i < 2048; i++)
                        {
                            x = a * i + b;
                            new_series.Points.AddXY(x, new_datas[i]);
                        }
                        this.chart1.ChartAreas[0].AxisX.Maximum = x;
                        this.chart1.Series.Add(new_series);
                    }
                    else if (AlgType == "平滑滤波")
                    {
                        string key = series.Split('-')[1];
                        AnalysisElement element = new AnalysisElement();
                        List<float> datas = new List<float>();
                        List<float> new_datas = new List<float>();
                        Smooth smooth = new Smooth();
                        chartElements.TryGetValue(key, out element);
                        datas = element.energy;
                        new_datas = smooth.mobile_Smoothing(datas, datas.Count, n);

                        var index = this.chart1.Series.IndexOf(key);
                        this.chart1.Series.RemoveAt(index);
                        Series new_series = new Series(key);
                        new_series.ChartType = SeriesChartType.Spline;
                        new_series.ChartArea = "Energy";
                        //series.Label = "#VAL";
                        //series.ToolTip = "#VALX:#VAL";
                        double x = 0;
                        for (int i = 0; i < 2048; i++)
                        {
                            x = a * i + b;
                            new_series.Points.AddXY(x, new_datas[i]);
                        }
                        this.chart1.ChartAreas[0].AxisX.Maximum = x;
                        this.chart1.Series.Add(new_series);
                    }
                }
            }
        }

        /// <summary>
        /// TextBox8输入是否为整形检测（滤波算法参数设置）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox8_TextChanged(object sender, EventArgs e)
        {
            string nn = textBox8.Text;
            if (nn != "")
            {
                double n = Convert.ToDouble(nn);
                if (n != (int)n)
                {
                    MessageBox.Show("输入不为整形！", "提示");
                    textBox8.Text = "";
                }
            }
        }

        /// <summary>
        /// ListView3修改按键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            string CWStart = textBox3.Text;//卡窗起点
            string CWEnd = textBox4.Text;//卡窗终点
            string CWPeak = textBox5.Text;//卡窗峰位

            this.chart1.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(CWStart);
            this.chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(CWEnd);
        }

        /// <summary>
        /// 菜单栏（Figure Save）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string file_path = string.Format(@"{0}\..\..\..\Save Image", Application.StartupPath);
                if (!System.IO.Directory.Exists(file_path))
                {
                    //若文件夹不存在，则创建文件夹
                    System.IO.Directory.CreateDirectory(file_path);
                }
                //若文件夹存在，则保存图片
                string file_name = System.DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                string path = file_path + "\\" + file_name + ".png";
                this.chart1.SaveImage(path, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
            }
            catch (Exception)
            {
                MessageBox.Show("保存错误，请重新保存！", "警告");
            }
        }

        /// <summary>
        /// 菜单栏（Figure Save As）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string file_path = string.Format(@"{0}\..\..\..\Save Image", Application.StartupPath);
                if (!System.IO.Directory.Exists(file_path))
                {
                    //若文件夹不存在，则创建文件夹
                    System.IO.Directory.CreateDirectory(file_path);
                }
                //若文件夹存在，则保存图片
                string file_name = System.DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                //string save_path = file_path + "\\" + file_name + ".png";

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "图片文件（*.png）|*.png|图片文件（*.jpg）|*.jpg|图片文件（*.jpeg）|*.jpeg";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = file_name;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = saveFileDialog.FileName;
                    try
                    {
                        this.chart1.SaveImage(fileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("保存错误，请重新保存！", "警告");
                    }
                    finally
                    {
                        saveFileDialog.Dispose();
                    }

                }
            }
            catch (Exception)
            {
                MessageBox.Show("保存错误，请重新保存！", "警告");
            }
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                textBox6.Text = item.SubItems[3].Text;
                textBox7.Text = item.SubItems[4].Text;
            }
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView3_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView3.SelectedItems)
            {
                textBox3.Text = item.SubItems[5].Text;
                textBox4.Text = item.SubItems[6].Text;
                textBox5.Text = item.SubItems[7].Text;
            }
        }
    }
}
