using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using Excel = Microsoft.Office.Interop.Excel;

namespace baes
{
    public partial class Form1 : Form
    {
        System.Data.DataTable dt;
        public double[,] massiv1;
        public double[,] massiv2;
        public double[,] massiv3;
        public List<double> massiv11;
        public List<double> massiv21;
        public List<double> massiv31;
        public List<double> massiv12;
        public List<double> massiv22;
        public List<double> massiv32;
        public List<double> massiv13;
        public List<double> massiv23;
        public List<double> massiv33;
        public double[,] shans1;
        public double[,] shans2;
        public double[,] shans3;
        public List<double> box= new List<double>();
        public List<double> box1 = new List<double>();
        public ZedGraphControl zedGraphControl;
        PictureBox pb;
        Bitmap bmp1;
        Bitmap bmp2;
        PointPairList coordinats1;
        PointPairList coordinats2;
        private OleDbCommand selectCommand;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExportExcel();

            sorting();
            textBox1.KeyPress += textBox1_KeyPress;
            textBox2.KeyPress += textBox2_KeyPress;
            textBox3.KeyPress += textBox1_KeyPress;
        }
        public void sorting()
        {
            massiv11 = new List<double>();
            massiv21 = new List<double>();
            massiv31= new List<double>();
            massiv12 = new List<double>();
            massiv22 = new List<double>();
            massiv32 = new List<double>();
            massiv13 = new List<double>();
            massiv23 = new List<double>();
            massiv33 = new List<double>();
            massiv1 = new double[dataGridView1.RowCount-2, 2];
            massiv2 = new double[dataGridView1.RowCount-2, 2];
            massiv3 = new double[dataGridView1.RowCount-2, 2];
            for (int i = 1; i < dataGridView1.RowCount-2; i++)
            {
                massiv1[i-1, 0] = Convert.ToDouble(dataGridView1[0, i].Value);
                massiv1[i-1, 1] = Convert.ToDouble(dataGridView1[3, i].Value);
                massiv2[i - 1, 0] = Convert.ToDouble(dataGridView1[1, i].Value);
                massiv2[i - 1, 1] = Convert.ToDouble(dataGridView1[3, i].Value);
                massiv3[i - 1, 0] = Convert.ToDouble(dataGridView1[2, i].Value);
                massiv3[i - 1, 1] = Convert.ToDouble(dataGridView1[3, i].Value);
            }
            for(int i = 0; i < massiv1.Length/2; i++)
            {
                if (massiv1[i, 1] == 1)
                {
                    massiv11.Add(massiv1[i, 0]);
                    massiv21.Add(massiv2[i, 0]);
                    massiv31.Add(massiv3[i, 0]);
                }
            }
            
            
            ShellSort(massiv11);
            ShellSort(massiv21);
            ShellSort(massiv31);
            for (int i = 29; i < 100; i++)
            {
                massiv12.Add(massiv11.Count(j => j == i));

            }
            for (int i = 0; i < 2; i++)
            {
                massiv22.Add(massiv21.Count(j => j == i));

            }
            for (int i = 94; i < 181; i++)
            {
                massiv32.Add(massiv31.Count(j => j == i));

            }
            massiv12.RemoveAll(j => j == 0);
            massiv22.RemoveAll(j => j == 0);
            massiv32.RemoveAll(j => j == 0);
            shans1 = new double[massiv12.Count,2];
            shans2 = new double[massiv22.Count, 2];
            shans3 = new double[massiv32.Count, 2];
            massiv13= massiv11.Distinct().ToList();
            massiv23 = massiv21.Distinct().ToList();
            massiv33 = massiv31.Distinct().ToList();
            int count = 0;
            for(int i = 0; i < massiv12.Count; i++)
            {
                
                    shans1[i, 0] = massiv13[i];
                shans1[i, 1] = massiv12[i] / massiv11.Count;
                
            }
            for (int i = 0; i < massiv22.Count; i++)
            {

                shans2[i, 0] = massiv23[i];
                shans2[i, 1] = massiv22[i] / massiv21.Count;

            }
            for (int i = 0; i < massiv32.Count; i++)
            {

                shans3[i, 0] = massiv33[i];
                shans3[i, 1] = massiv32[i] / massiv31.Count;

            }


            //grapf(shans1);
            DrawGraph();
            (double det,double det1) = ret();
        }
        public (double ser, double ser2) ret()
        {
            double ser=0;
            double ser2=1;
            return (ser, ser2);
        }

        private static void ShellSort(List<double> Array)
        {
            int step = Array.Count / 2;
            while (step > 0)
            {
                int i, j;
                for (i = step; i < Array.Count; i++)
                {
                    double value = Array[i];
                    for (j = i - step; (j >= 0) && (Array[j] > value); j -= step)
                        Array[j + step] = Array[j];
                    Array[j + step] = value;
                }
                step /= 2;
            }
        }
        public double regres(int sizer,double[,]example1,double x)
        {
            double y = 0;
            double x1 = x_(sizer, example1);
            double y1 = y_(sizer, example1);
            double dx = Dx(sizer, x1, example1);
            double dy = Dy(sizer, y1, example1);
            double quadrox = Math.Sqrt(dx);
            double quadroy = Math.Sqrt(dy);
            double sumofxy = sumxy(example1);
            double kaef = koefofkor(sumofxy, sizer, x1, y1, quadrox, quadroy);
            double k = K(kaef, quadrox, quadroy);
            double b = B(k, x1, y1);
            y = Y(k, b, x);
            return y;
        }
        public double x_(int size, double[,] massiv)
        {
            double x = 0;
            x = sum(0, size, massiv) / size;
            return x;
        }
        public double y_(int size, double[,] massiv)
        {
            double y = 0;
            y = sum(1, size, massiv) / size;
            return y;
        }
        public double Dx(int size, double number, double[,] massiv)
        {
            double[,] dif = new double[size, 1];
            for (int i = 0; i < size; i++)
            {
                dif[i, 0] = Math.Pow((massiv[i, 0] - number), 2);
            }
            double dispers = 0;
            dispers = sum(0, size, dif) / size;
            return dispers;
        }
        public double Dy(int size, double number, double[,] massiv)
        {
            double[,] dif = new double[size, 1];
            for (int i = 0; i < size; i++)
            {
                dif[i, 0] = Math.Pow((massiv[i, 1] - number), 2);
            }
            double dispers = 0;
            dispers = sum(0, size, dif) / size;
            return dispers;
        }
        public double sum(int numberofrows, int countofcolumns, double[,] massiv)
        {
            double summa = 0;
            for (int i = 0; i < countofcolumns; i++)
            {
                summa += massiv[i, numberofrows];
            }
            return summa;
        }
        public double sumxy(double[,] massiv)
        {
            double sum = 0;
            for (int i = 0; i < massiv.Length / 2; i++)
            {
                sum += massiv[i, 0] * massiv[i, 1];
            }
            return sum;
        }
        public double koefofkor(double xy, int size, double x, double y, double quadx, double quady)
        {
            double kaef = 0;
            kaef = (xy - size * x * y) / (size * quadx * quady);
            return kaef;
        }
        public double K(double kaef, double quadx, double quady)
        {
            double k = 0;
            k = kaef * quady / quadx;
            return k;
        }
        public double B(double k, double x, double y)
        {
            double b = 0;
            b = (k * -x) + y;
            return b;
        }
        public double Y(double k, double b, double x)
        {
            double y = 0;
            y = k * x + b;
            return y;
        }
        private void ExportExcel()
        {

            string file = Application.StartupPath.ToString()+ @"\heart.xlsx"; //variable for the Excel File Location
            dt = new System.Data.DataTable(); //container for our excel data
            DataRow row;
            //DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            
                //file = openFileDialog1.FileName; //get the filename with the location of the file
                try
                {
                    //Create Object for Microsoft.Office.Interop.Excel that will be use to read excel file

                    Excel.Application excelApp = new Excel.Application();

                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);

                    Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                    Excel.Range excelRange = excelWorksheet.UsedRange;

                    int rowCount = excelRange.Rows.Count; //get row count of excel data

                    int colCount = excelRange.Columns.Count; // get column count of excel data

                    //Get the first Column of excel file which is the Column Name

                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                        }
                        break;
                    }

                    //Get Row Data of Excel

                    int rowCounter; //This variable is used for row index number
                    for (int i = 1; i <= rowCount; i++) //Loop for available row of excel data
                    {
                        row = dt.NewRow(); //assign new row to DataTable
                        rowCounter = 0;
                        for (int j = 1; j <= colCount; j++) //Loop for available column of excel data
                        {
                            //check if cell is empty
                            
                                row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                            
                            rowCounter++;
                        }
                        dt.Rows.Add(row); //add row to DataTable
                    }





                    //close and clean excel process
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(excelRange);
                    Marshal.ReleaseComObject(excelWorksheet);
                    //quit apps
                    excelWorkbook.Close();
                    Marshal.ReleaseComObject(excelWorkbook);
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                DataSet1 dataSet1 = new DataSet1();
                dataSet1.Tables.Add(dt);
                DataTable DataTable1 = new DataTable();
                DataTable1=dataSet1.Tables[1];

            }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                    return;
                }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            box.Clear();
            double s1 = 0;
            double s2 = 0;
            double s3 = 0;
            box.Add(Convert.ToDouble(textBox1.Text));
            box.Add(Convert.ToDouble(textBox2.Text));
            box.Add(Convert.ToDouble(textBox3.Text));
            for(int i = 0; i < shans1.Length/2; i++)
            {
                if(shans1[i,0] == box[0])
                {
                    s1 = shans1[i,1];
                }
            }
            for (int i = 0; i < shans2.Length/2; i++)
            {
                if (shans2[i, 0] == box[1])
                {
                    s2 = shans2[i, 1];
                }
            }
            for (int i = 0; i < shans3.Length / 2; i++)
            {
                if (shans3[i, 0] == box[2])
                {
                    s3 = shans3[i, 1] ;
                }
                
            }
            if (s1 == 0)
            {
                s1 = regres(shans1.Length / 2, shans1, box[0]);
            }
            if (s3 == 0)
            {
                s3 = regres(shans3.Length / 2, shans3, box[2]);
            }
            double chekc = s1 * s2 * s3;
            if ((1-s1)*(1-s2)*(1-s3)>0.5)
            {
                label1.Text = "болен";
            }
            else
            {
                label1.Text = "здоров";
            }
            if (textBox4.Text == "")
            {
                textBox4.Text = "30";
            }
            label2.Text = Convert.ToString(test(Convert.ToDouble(textBox4.Text)));
        }
        public double test(double longer)
        {
            double srednee = 0;
            double sum = 0;
            double testanswer = 0;
            int testcount = 0;
            for(int i = 1; i < longer; i++)
            {
                double s1 = 0;
                double s2 = 0;
                double s3 = 0;
                box1.Clear();
                box1.Add(Convert.ToDouble(dataGridView1[0, i].Value));
                box1.Add(Convert.ToDouble(dataGridView1[1, i].Value));
                box1.Add(Convert.ToDouble(dataGridView1[2, i].Value));
                for (int j = 1; j < shans1.Length / 2; j++)
                {
                    if (shans1[j, 0] == box1[0])
                    {
                        s1 =  shans1[j, 1];
                    }
                    
                        
                    
                }
                
                for (int j = 0; j < shans2.Length / 2; j++)
                {
                    if (shans2[j, 0] == box1[1])
                    {
                        s2 = shans2[j, 1];
                    }
                }
                for (int j = 1; j < shans3.Length / 2; j++)
                {
                    if (shans3[j, 0] == box1[2])
                    {
                        s3 =  shans3[j, 1]; ;
                    }
                    
                }
                if (s1 == 0)
                {
                    s1 = regres(shans1.Length / 2, shans1, box[0]);
                }
                if (s3 == 0)
                {
                    s3 = regres(shans3.Length / 2, shans3, box[2]);
                }
                double sravn = 0;
                
                if ( (1 - s1) * (1 - s2) * (1 - s3)>0.5)
                {
                    sravn = 1;
                    if (sravn == Convert.ToDouble(dataGridView1[3, i].Value))
                    {
                        testcount++;
                    }
                    
                }
                else
                {
                    if (sravn == Convert.ToDouble(dataGridView1[3, i].Value))
                    {
                        testcount++;
                    }
                    
                }
                sum += s1 * s2 * s3;
            }
            srednee = sum / longer;
            testanswer = testcount / longer;
            return testanswer;
        }
        public void grapf(double[,] massiv)
        {
            try
            {
                coordinats1 = new PointPairList();
                for (int i = 0; i < massiv.Length / 2; i++)
                {
                    coordinats1.Add(massiv[i, 0], massiv[i, 1]);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Введены неправельные величины");
            }

            GraphPane myPane1 = new GraphPane();
            zedGraphControl.GraphPane = myPane1;

            myPane1.Title.Text = "График распределения";//подпись графика
            myPane1.Fill = new Fill(Color.White, Color.White, 45.0f);//фон графика
            myPane1.Chart.Fill.Type = FillType.Brush;
            myPane1.Legend.Position = LegendPos.Float;
            myPane1.Legend.IsHStack = false;
            LineItem myCurve1 = myPane1.AddCurve("График распределения", coordinats1, Color.MidnightBlue, SymbolType.Diamond);
            myCurve1.Symbol.Fill = new Fill(Color.White);
            //myCurve1.Line.IsVisible = false;
            myCurve1.Symbol.Fill.Color = Color.Blue;
            LineItem myCurve2 = myPane1.AddCurve("График распределения", coordinats2, Color.Black, SymbolType.Diamond);
            myCurve2.Symbol.Fill = new Fill(Color.White);
            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();
            zedGraphControl1.Visible = true;
            zedGraphControl1.IsEnableWheelZoom = false;

        }
        
        private void DrawGraph()
        {
            // Удалим существующую панель с графиком
            zedGraphControl1.MasterPane.PaneList.Clear();
            zedGraphControl2.MasterPane.PaneList.Clear();
            zedGraphControl3.MasterPane.PaneList.Clear();
            // Создадим две панели для графика, где будут отображаться
            // одинаковые данные, но с разными значениями BarType
            GraphPane pane1 = new GraphPane();
            GraphPane pane2 = new GraphPane();
            GraphPane pane3 = new GraphPane();
            // Количество столбцов

            

            // Сгенерируем данные для высот столбцов
            double[] YValues1 = GenerateData(massiv12.Count,massiv12);
            double[] YValues2 = GenerateData(massiv22.Count, massiv22);
            double[] YValues3 = GenerateData(massiv32.Count, massiv32);

            double[] XValues1 = new double[massiv12.Count];
            double[] XValues2 = new double[massiv22.Count];
            double[] XValues3 = new double[massiv32.Count];
            // Заполним данные
            for (int i = 0; i < massiv12.Count; i++)
            {
                XValues1[i] = i+29;
            }
            for (int i = 0; i < massiv22.Count; i++)
            {
                XValues2[i] = i;
            }
            for (int i = 0; i < massiv32.Count; i++)
            {
                XValues3[i] = i+94;
            }
            // По одинаковым данным построим две гистограммы
            CreateBars(pane1, XValues1, YValues1);
            CreateBars(pane2, XValues2, YValues2);
            CreateBars(pane3, XValues3, YValues3);
            // !!! У первого графика столбцы накладываются один на другой
            // всегда в одинаковой последовательности:
            // впереди синий, затем красный, затем желтый
            pane1.BarSettings.Type = BarType.Overlay;
            pane1.Title.Text = "Возраст";
            

            //!!!У второго графика порядок наложения столбцов такой, чтобы все они были видны
            pane2.BarSettings.Type = BarType.Overlay;
            pane2.Title.Text = "Пол";
            pane3.BarSettings.Type = BarType.Overlay;
            pane3.Title.Text = "Давление";
            // Добавим созданные панели в MasterPane
            zedGraphControl1.MasterPane.Add(pane1);
            zedGraphControl2.MasterPane.Add(pane2);
            zedGraphControl3.MasterPane.Add(pane3);
            //zedGraphControl1.MasterPane.Add(pane2);

            // Зададим расположение графиков
            using (Graphics g1 = CreateGraphics())
            {
                // Графики будут размещены в один столбец друг под другом
                zedGraphControl1.MasterPane.SetLayout(g1, PaneLayout.ForceSquare);
            }
            using (Graphics g2 = CreateGraphics())
            {
                // Графики будут размещены в один столбец друг под другом
                zedGraphControl2.MasterPane.SetLayout(g2, PaneLayout.ForceSquare);
            }
            using (Graphics g3 = CreateGraphics())
            {
                // Графики будут размещены в один столбец друг под другом
                zedGraphControl3.MasterPane.SetLayout(g3, PaneLayout.ForceSquare);
            }

            // Обновим данные об осях
            zedGraphControl1.AxisChange();

            // Обновляем график
            zedGraphControl1.Invalidate();
            zedGraphControl2.AxisChange();

            // Обновляем график
            zedGraphControl2.Invalidate();
            zedGraphControl3.AxisChange();

            // Обновляем график
            zedGraphControl3.Invalidate();
        }
        private static void CreateBars(GraphPane pane,
        double[] XValues,
        double[] YValues1)
        {
            pane.CurveList.Clear();

            // Создадим три гистограммы
            pane.AddBar("", XValues, YValues1, Color.Blue);
            
        }
        private double[] GenerateData(int count,List<double>massiv)
        {
            double[] values = new double[massiv.Count];

            //Заполним данные
            for (int i = 0; i < massiv.Count; i++)
            {
                values[i] = massiv[i];
            }

            return values;
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ( e.KeyChar != Convert.ToChar(48)&& e.KeyChar != Convert.ToChar(49) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }
        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
