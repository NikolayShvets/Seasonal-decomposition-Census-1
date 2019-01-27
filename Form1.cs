using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection; // указ ссылка на использование типов в пространстве имен System.Reflection, при этом уточнение типа не требуется 
using ExcelObj = Microsoft.Office.Interop.Excel; // создание псевдонима простарнства имен
using MathNet.Numerics.Statistics;


namespace GettingDataAndPlot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void OutButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new
            OpenFileDialog();
            //Задаем расширение имени файла по умолчанию
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, определяющая варианты, доступные в 
            //поле "файлы типа" диалогового окна
            ofd.Filter = " Excel 2010(*.xlsx)|*.xlsx|Excel 2010(*.xls)|*.xls";
            //задаем заголовок диалогового окна
            ofd.Title = "Выберите документ для выгрузки";
            ExcelObj.Application app = new
            ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value);
                //Устанавливаем номер листа из которого будут извлекаться данные, нумерация с 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(new
                        DataColumn((ShtRange.Cells[1, Cnum]
                        as
                        ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();
                string[] columnNames = new
                string[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }
                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                dataGridView1.DataSource = dt;
                app.Quit();
            }
            else
                Application.Exit();
        }

        private void Plotting_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 1;
            List<double> DataList = new List<double>();
            List<double> XList = new List<double>();
            chart1.Series[0].BorderWidth = 1;
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Maximum = dataGridView1.Rows.Count - 1;
            chart1.ChartAreas[0].AxisX.MajorGrid.Interval = 12;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                DataList.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells["AvTemp"].Value));
                XList.Add(i);

            }
            chart1.Series[0].Points.DataBindXY(XList, DataList);
        }

        private void AutoCorr_Click(object sender, EventArgs e)
        {
            ///<summary>
            ///Вычисление автокорреляции.
            ///Делаем две копии ряда,
            ///затем копии выравниваем и перемножаем одинаковые во времени значения ряда,
            ///суммируем результат перемножения --> коэф. автокорр. для смещения 0 (Corr(0)),
            ///далее сдвигаем копии на 1 шаг во времени --> коэф. автокорр. для смещения 1 (Corr(1)).
            ///Повторяем вычисления для всей длины ряда, увеличивая смещение на 1 шаг на каждой итерции.
            ///</summary>
            //пересчиать на листке, разобраться с выборочными средними и подправить, а так все ок.

            this.tabControl1.SelectedIndex = 2;


            List<Nullable<double>> StatCopy = new List<Nullable<double>>();
            List<Nullable<double>> DinamCopy = new List<Nullable<double>>();
            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; ++i)
            {
                for (int j = 1; j < this.dataGridView1.Columns.Count; ++j)
                {
                    StatCopy.Add(Convert.ToDouble(dataGridView1[j, i].Value));
                    DinamCopy.Add(Convert.ToDouble(dataGridView1[j, i].Value));
                }

            }

            if (StatCopy.Count != DinamCopy.Count)
            {
                MessageBox.Show("Series must be the same length");
                this.Close();
            }
            for (int lag = 1; lag < StatCopy.Count; lag++)
            {
                double sumOfstat = 0;
                double sumOfstat1 = 0;
                for (int i = lag; i < StatCopy.Count; i++)
                {
                    sumOfstat += Convert.ToDouble(StatCopy[i]);
                    sumOfstat1 += Convert.ToDouble(StatCopy[i - lag]);
                }
                var avg = sumOfstat / (StatCopy.Count - lag);
                var avg1 = sumOfstat1 / (StatCopy.Count - lag);
                double chisl = 0;
                double znamMn1 = 0;
                double znamMn2 = 0;
                double r = 0;
                //считаем коэффициент автокорреляции уровней ряда порядка lag
                //сумма(lag --> n) (yt - y1)(yt-1 - y2)/sqrt(сумма(lag --> n) sqr(yt-y1) * сумма(lag --> n) sqr(yt-1 - y2))
                for (int i = lag; i < StatCopy.Count; i++)
                {
                    chisl += (Convert.ToDouble(StatCopy[i]) - avg) * (Convert.ToDouble(StatCopy[i - lag]) - avg1);
                    znamMn1 += Math.Pow((Convert.ToDouble(StatCopy[i]) - avg), 2);
                    znamMn2 += Math.Pow((Convert.ToDouble(StatCopy[i - lag]) - avg1), 2);
                }
                r = chisl / Math.Sqrt(znamMn1 * znamMn2);
                //MessageBox.Show(Convert.ToString(r),"Коэфф. автокорреляции уровней ряда порядка lag");
                //MessageBox.Show(Convert.ToString(chisl), "Числитель");
                CorrData.Rows.Add(lag, r);
            }
            //Cчитатаем сезонность как количество лагов между максимальными показателями коэф-та автокорр.
            double maxCorr = Convert.ToDouble(CorrData[1, 0].Value);
            int Seasonality = 0;
            for (int i = 0; i < this.CorrData.Rows.Count - 1; i++)
            {
                if ((Convert.ToDouble(CorrData[1, i].Value) <= Convert.ToDouble(CorrData[1, i + 1].Value)) && (Convert.ToDouble(CorrData[1, i + 1].Value) > Convert.ToDouble(CorrData[1, i + 2].Value)))
                {
                    Seasonality += (i + 1) + 1;
                    maxCorr = Convert.ToDouble(CorrData[1, Seasonality - 1].Value);

                    break;
                }

                else
                {
                    continue;
                }

            }
            LagText.Text = Convert.ToString(Seasonality);
            CorrText.Text = Convert.ToString(maxCorr);
        }

        private void ShowThePlot_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 1;
        }

        private void SeasonalButton_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 3;
            bool flag = false;
            //Высчитываем скользящие средние, центрированные скользящие средние и сезонные отклонения
            int window = Convert.ToInt32(LagText.Text);
            int k = 0;
            double sum = 0;
            double moveAvg;
            //List<double> CensusCopy = new List<double> {6.0, 4.4, 5.0, 9.0, 7.2, 4.8, 6, 10.0, 8.0, 5.6, 6.4, 11.0, 9.0, 6.6, 7.0, 10.8 };
            //List<double> CensusCopy = new List<double> { 10, 12, 13, 16, 19, 23, 26, 30, 28, 18, 16, 14 };
            //List<double> CensusCopy = new List<double> { 72, 100, 90, 64, 70, 98, 80, 58, 62, 80, 68, 48, 52, 60, 50, 30 };
            List<double> CensusCopy = new List<double>();
            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; ++i)
            {
                for (int j = 1; j < this.dataGridView1.Columns.Count; ++j)
                {
                    CensusCopy.Add(Convert.ToDouble(dataGridView1[j, i].Value));
                }

            }
            SeasDecomp.Columns.Add("Count", "Count");
            SeasDecomp.Columns.Add("Data", "Data");
            SeasDecomp.Columns.Add("MovingAvg", "MovingAvg");
            for (int i = 0; i < CensusCopy.Count; i++)
            {
                SeasDecomp.Rows.Add(Convert.ToString(i + 1));
                SeasDecomp.Rows[i].Cells["Data"].Value = CensusCopy[i];
            }



            for (int i = 0; i < CensusCopy.Count - window + 1; i++)
            {
                for (int j = i; j < window + i; j++)
                {
                    sum += CensusCopy[j];
                }
                moveAvg = (sum / window);
                SeasDecomp.Rows[i + (window / 2)].Cells["MovingAvg"].Value = moveAvg;
                sum = 0;
            }
            if (Convert.ToInt32(this.LagText.Text) % 2 == 0)
            {
                flag = true;
                List<double> CentrCopy = new List<double>();
                double CentrMoveAvg;
                for (int i = 0; i < this.SeasDecomp.Rows.Count - 1; ++i)
                {
                    if (SeasDecomp.Rows[i].Cells["MovingAvg"].Value == null)
                    {
                        continue;
                    }
                    else
                    {
                        CentrCopy.Add(Convert.ToDouble(SeasDecomp.Rows[i].Cells["MovingAvg"].Value));
                    }
                }

                SeasDecomp.Columns.Add("CentrMovingAvg", "CentrMovingAvg");
                for (int i = 0; i < CentrCopy.Count - 1; i++)
                {
                    for (int j = i; j < 2 + i; j++)
                    {
                        sum += CentrCopy[j];
                    }
                    CentrMoveAvg = (sum / 2);
                    SeasDecomp.Rows[i + (window / 2)].Cells["CentrMovingAvg"].Value = CentrMoveAvg;
                    sum = 0;
                }
            }
            //Считаем сезонные отклонения
            SeasDecomp.Columns.Add("Diffrncs", "Diffrncs");
            if (radioAdd.Checked == true)
            {
                if (flag == true)
                {
                    for (int i = 0/*(window / 2)*/ ; i < SeasDecomp.Rows.Count - 1; i++)
                    {
                        if (SeasDecomp.Rows[i].Cells["CentrMovingAvg"].Value == null)
                        {
                            continue;
                        }
                        SeasDecomp.Rows[i].Cells["Diffrncs"].Value = CensusCopy[i] - Convert.ToDouble(SeasDecomp.Rows[i].Cells["CentrMovingAvg"].Value);
                    }
                }
                else
                {
                    for (int i = 0/*(window / 2)*/; i < SeasDecomp.Rows.Count - 1; i++)
                    {
                        if (SeasDecomp.Rows[i].Cells["MovingAvg"].Value == null)
                        {
                            continue;
                        }
                        SeasDecomp.Rows[i].Cells["Diffrncs"].Value = CensusCopy[i] - Convert.ToDouble(SeasDecomp.Rows[i].Cells["MovingAvg"].Value);
                    }
                }
            }
            else if (radioMult.Checked == true)
            {
                for (int i = 0/*(window / 2)*/ ; i < SeasDecomp.Rows.Count - 1; i++)
                {
                    if (SeasDecomp.Rows[i].Cells["CentrMovingAvg"].Value == null)
                    {
                        continue;
                    }
                    SeasDecomp.Rows[i].Cells["Diffrncs"].Value = CensusCopy[i] / Convert.ToDouble(SeasDecomp.Rows[i].Cells["CentrMovingAvg"].Value);
                }
            }
            //Определяем сезонную составляющую
            SeasDecomp.Columns.Add("SeasonalFactors", "SeasonalFactors");
            List<double> SeasonCopy = new List<double>();
            for (int i = 0; i < this.SeasDecomp.Rows.Count - 1; i++)
            {
                if (SeasDecomp.Rows[i].Cells["Diffrncs"].Value == null)
                {
                    continue;
                }
                else
                {
                    SeasonCopy.Add(Convert.ToDouble(SeasDecomp.Rows[i].Cells["Diffrncs"].Value));
                }
            }
            List<double> SeasAvg = new List<double>();
            double SeasSum = 0;
            double SumSeasAvg = 0;
            double CorrectKoef;
            int p = 0;
            int count = (SeasDecomp.Rows.Count - window) / window;
            for (int i = 0; i < /*SeasDecomp.Rows.Count -*/ window; i++)
            {
                for (int j = i; j < SeasDecomp.Rows.Count - window - 1; j++)
                {

                    if (p % window == 0)
                    {
                        SeasSum += SeasonCopy[j];
                        p += 1;
                    }
                    else
                    {
                        p += 1;
                        continue;

                    }

                }
                //MessageBox.Show(Convert.ToString(count));
                //MessageBox.Show(Convert.ToString(SeasSum));
                SeasAvg.Add(SeasSum / count);
                //MessageBox.Show(Convert.ToString(SeasAvg[i]));
                SumSeasAvg += SeasAvg[i];
                SeasSum = 0;
                p = 0;
            }
            //MessageBox.Show(Convert.ToString(SumSeasAvg));
            k = 0;
            CorrectKoef = SumSeasAvg / window;
            //MessageBox.Show(Convert.ToString(CorrectKoef));
            for (int i = 0; i < SeasDecomp.Rows.Count - 1/*SeasAvg.Count*/; i++)
            {
                if (k == window)
                {
                    k = 0;
                }
                if (i % (window) == 0)
                {
                    k = window / 2;
                }

                SeasDecomp.Rows[i].Cells["SeasonalFactors"].Value = SeasAvg[k] - CorrectKoef;
                //MessageBox.Show(Convert.ToString(SeasDecomp.Rows[i].Cells["SeasonalFactors"].Value));
                k += 1;
            }
            //AdjustedSeries (вычитаем сезонные составляющие из исходного ряда)
            SeasDecomp.Columns.Add("AdjustedSeries", "AdjustedSeries");
            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                SeasDecomp.Rows[i].Cells["AdjustedSeries"].Value = Convert.ToDouble(SeasDecomp.Rows[i].Cells["Data"].Value) - Convert.ToDouble(SeasDecomp.Rows[i].Cells["SeasonalFactors"].Value);
            }
            //Тренд
            double CountAvg = 0;
            double DataAvg = 0;
            double SummCount = 0;
            double SummData = 0;
            double SummCountData = 0;
            double SummSqrCount = 0;
            double b = 0;
            double a = 0;
            string TrendEquation = "";
            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                SummCount += Convert.ToDouble(SeasDecomp.Rows[i].Cells["Count"].Value);
            }
            CountAvg = SummCount / (SeasDecomp.Rows.Count - 1);
            //MessageBox.Show(Convert.ToString(CountAvg));
            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                SummData += Convert.ToDouble(SeasDecomp.Rows[i].Cells["Data"].Value);
            }
            DataAvg = SummData / (SeasDecomp.Rows.Count - 1);
            //MessageBox.Show(Convert.ToString(DataAvg));
            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                SummCountData += Convert.ToDouble(SeasDecomp.Rows[i].Cells["Count"].Value) * Convert.ToDouble(SeasDecomp.Rows[i].Cells["Data"].Value);
            }
            for (int i = 0; i < SeasDecomp.Rows.Count; i++)
            {
                SummSqrCount += Math.Pow(Convert.ToDouble(SeasDecomp.Rows[i].Cells["Count"].Value), 2.0);
            }
            //MessageBox.Show(Convert.ToString(SummSqrCount));
            b = (SummCountData - (SeasDecomp.Rows.Count/* - 1*/) * CountAvg * DataAvg) / (SummSqrCount - (SeasDecomp.Rows.Count - 1) * Math.Pow(CountAvg, 2.0));
            //MessageBox.Show(Convert.ToString(b));
            a = DataAvg - b * CountAvg;
            //MessageBox.Show(Convert.ToString(a));
            TrendEquation = ("Yt = " + Convert.ToString(a) + " + " + Convert.ToString(b) + " * Xt");
            MessageBox.Show(TrendEquation);


            //График сезонной составляющей

            List<double> DataList = new List<double>();
            List<double> XList = new List<double>();
            chart2.Series[0].BorderWidth = 1;
            chart2.ChartAreas[0].AxisX.Minimum = 0;
            chart2.ChartAreas[0].AxisX.Maximum = SeasDecomp.Rows.Count - 1;
            chart2.ChartAreas[0].AxisX.MajorGrid.Interval = window;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                DataList.Add(Convert.ToDouble(SeasDecomp.Rows[i].Cells["SeasonalFactors"].Value));
                XList.Add(i);

            }
            chart2.Series[0].Points.DataBindXY(XList, DataList);
            //График тренда

            double[] x = new double[SeasDecomp.Rows.Count - 1];
            double[] y = new double[SeasDecomp.Rows.Count - 1];
            chart3.Series[0].BorderWidth = 1;
            chart3.ChartAreas[0].AxisX.Minimum = 0;
            chart3.ChartAreas[0].AxisY.Minimum = a - 1;
            chart3.ChartAreas[0].AxisX.Maximum = SeasDecomp.Rows.Count - 1;
            chart3.ChartAreas[0].AxisX.MajorGrid.Interval = window;
            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                y[i] = a + b * (Convert.ToDouble(SeasDecomp.Rows[i].Cells["Count"].Value));
                x[i] = i;
            }
            chart3.Series[0].Points.DataBindXY(x, y);
            // График без сезонности
            List<double> AdjSeries = new List<double>();
            List<double> XAdjSeries = new List<double>();
            chart4.Series[0].BorderWidth = 1;
            chart4.ChartAreas[0].AxisX.Minimum = 0;
            chart4.ChartAreas[0].AxisX.Maximum = SeasDecomp.Rows.Count - 1;
            chart4.ChartAreas[0].AxisX.MajorGrid.Interval = window;
            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                AdjSeries.Add(Convert.ToDouble(SeasDecomp.Rows[i].Cells["AdjustedSeries"].Value));
                XAdjSeries.Add(i);
            }
            chart4.Series[0].Points.DataBindXY(XAdjSeries, AdjSeries);

            //сглаженный тренд цикл
            List<double> STSeries = new List<double>();
            List<double> XSTSeries = new List<double>();
            chart5.Series[0].BorderWidth = 1;
            chart5.ChartAreas[0].AxisX.Minimum = 0;
            chart5.ChartAreas[0].AxisX.Maximum = SeasDecomp.Rows.Count - 1;
            chart5.ChartAreas[0].AxisX.MajorGrid.Interval = window;

            for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                STSeries.Add(Convert.ToDouble(SeasDecomp.Rows[i].Cells["Data"].Value) - Convert.ToDouble(SeasDecomp.Rows[i].Cells["SeasonalFactors"].Value));
                XSTSeries.Add(i);
            }
            chart5.Series[0].Points.DataBindXY(XSTSeries, STSeries);

            //ряд разностей
            List<double> DiffSeries = new List<double>();
            List<double> XDiffSeries = new List<double>();
            chart6.Series[0].BorderWidth = 1;
            chart6.ChartAreas[0].AxisX.Minimum = 0;
            chart6.ChartAreas[0].AxisX.Maximum = SeasDecomp.Rows.Count - 1;
            chart6.ChartAreas[0].AxisX.MajorGrid.Interval = window;

           /* for (int i = 0; i < SeasDecomp.Rows.Count - 1; i++)
            {
                DiffSeries.Add(Convert.ToDouble(SeasDecomp.Rows[i].Cells["Diffrncs"].Value) - Convert.ToDouble(SeasDecomp.Rows[i].Cells["SeasonalFactors"].Value));
                XDiffSeries.Add(i);
            }
            chart6.Series[0].Points.DataBindXY(XDiffSeries, DiffSeries);*/
        }

        private void OutButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new
            OpenFileDialog();
            //Задаем расширение имени файла по умолчанию
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, определяющая варианты, доступные в 
            //поле "файлы типа" диалогового окна
            ofd.Filter = " Excel 2010(*.xlsx)|*.xlsx|Excel 2010(*.xls)|*.xls";
            //задаем заголовок диалогового окна
            ofd.Title = "Выберите документ для выгрузки";
            ExcelObj.Application app = new
            ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value);
                //Устанавливаем номер листа из которого будут извлекаться данные, нумерация с 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(new
                        DataColumn((ShtRange.Cells[1, Cnum]
                        as
                        ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();
                string[] columnNames = new
                string[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }
                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                dataGridView2.DataSource = dt;
                app.Quit();
            }
            else
                Application.Exit();
        }

        private void Separate_Click(object sender, EventArgs e)
        {
            List<double>[] IntervalMassive = new List<double>[21];
            List<double> TempDiap = new List<double>();
            List<double> Cases = new List<double>();
            List<double> Tvg = new List<double>();
            List<double> First = new List<double>();
            List<double> FirstCases = new List<double>();
            List<double> FirstTvg = new List<double>();
            List<double> Second = new List<double>();
            List<double> SecondCases = new List<double>();
            List<double> SecondTvg = new List<double>();
            List<double> Third = new List<double>();
            List<double> ThridCases = new List<double>();
            List<double> ThridTvg = new List<double>();
            List<double> Fourth = new List<double>();
            List<double> FourthCases = new List<double>();
            List<double> FourthTvg = new List<double>();
            List<double> Fifth = new List<double>();
            List<double> FifthCases = new List<double>();
            List<double> FifthTvg = new List<double>();
            List<double> Sixth = new List<double>();
            List<double> SixthCases = new List<double>();
            List<double> SixthTvg = new List<double>();
            double DiapWin = 5;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++ )
            {
                if (Convert.ToDouble(dataGridView2.Rows[i].Cells["Значение параметра AVGТпол_вх_расч"].Value) > 1.53 && Convert.ToDouble(dataGridView2.Rows[i].Cells["Значение параметра AVGТпол_вх_расч"].Value) < 28.17)
                {
                    TempDiap.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells["Значение параметра AVGТпол_вх_расч"].Value));
                    Tvg.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells["Значение параметра AVGТвг_дв1_расч"].Value));
                    //dataGridView3.Rows[i].Cells["Cases"].Value = dataGridView2.Rows[i].Cells["Значение параметра AVGТпол_вх_расч"].Value;
                    Cases.Add(i);
                    //dataGridView3.Rows[i].Cells["Case"].Value = i;
                }
                else
                {
                    continue;
                }
            }
            for (int i = 0; i < TempDiap.Count - 1; i++)
            {
                if (TempDiap[i] >= 1.53 && TempDiap[i] < 1.53 + DiapWin)
                {
                    First.Add(TempDiap[i]);
                    FirstCases.Add(Cases[i]);
                    FirstTvg.Add(Tvg[i]);
                }
                if (TempDiap[i] >= 1.53 + DiapWin && TempDiap[i] < 1.53 + 2 * DiapWin)
                {
                    Second.Add(TempDiap[i]);
                    SecondCases.Add(Cases[i]);
                    SecondTvg.Add(Tvg[i]);
                }
                if (TempDiap[i] >= 1.53 + 2 * DiapWin && TempDiap[i] < 1.53 + 3 * DiapWin)
                {
                    Third.Add(TempDiap[i]);
                    ThridCases.Add(Cases[i]);
                    ThridTvg.Add(Tvg[i]);
                }
                if (TempDiap[i] >= 1.53 + 3 * DiapWin && TempDiap[i] < 1.53 + 4 * DiapWin)
                {
                    Fourth.Add(TempDiap[i]);
                    FourthCases.Add(Cases[i]);
                    FourthTvg.Add(Tvg[i]);
                }
                if (TempDiap[i] >= 1.53 + 4 * DiapWin && TempDiap[i] < 1.53 + 5 * DiapWin)
                {
                    Fifth.Add(TempDiap[i]);
                    FifthCases.Add(Cases[i]);
                    FifthTvg.Add(Tvg[i]);
                }
                if (TempDiap[i] >= 1.53 + 5 * DiapWin && TempDiap[i] < 1.53 + 6 * DiapWin)
                {
                    Sixth.Add(TempDiap[i]);
                    SixthCases.Add(Cases[i]);
                    SixthTvg.Add(Tvg[i]);
                }
            }
            for (int i = 0; i < FifthCases.Count - 1; i++)
            {
                dataGridView3.Rows.AddCopies(i, 1);
                dataGridView3.Rows[i].Cells["Case"].Value = Convert.ToString(FifthCases[i]);
                dataGridView3.Rows[i].Cells["Tvg"].Value = Convert.ToString(FifthTvg[i]);
            }
        }

        private void Next_Click(object sender, EventArgs e)
        {
           
        }
    }
    
}
