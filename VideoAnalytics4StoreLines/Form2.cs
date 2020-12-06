using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;


using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;
using System.Data.SqlServerCe;
using System.Windows.Forms.DataVisualization.Charting;


namespace VideoAnalytics4StoreLines
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        string path = @"Data Source = D:\12112015\company\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";

        private void button1_Click(object sender, EventArgs e)
        {
            //Image im = new Bitmap(@"pic4.jpg");
            //Image im = new Bitmap("myGray.png");
            Image im = Image.FromFile(@"myGray.png");
            Image im2 = im;
            pictureBox1.Image = im2;
            pictureBox1.Refresh();

            float[] px = new float[]{
                                  41,
                                  72,
                                  73,
                                  73,
                                  73,
                                  73,
                                  100,
                                  120,
                                  130,
                                  130,
                                  131,
                                  132,
                                  73
                                  
                              };
            float[] py = new float[]{
                                  41,
                                  72,
                                  73,
                                  73,
                                  73,
                                  73,
                                  100,
                                  110,
                                  112,
                                  113,
                                  114,
                                  115,
                                  73
                                  
                              };


            Image canvas = HeatMap.NET.HeatMap.GenerateHeatMap(im, px, py);
            pictureBox2.Image = canvas;

            //canvas.Save("Jenna.Heated.Jpg", ImageFormat.Jpeg);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                double alpha = trackBar1.Value;
                alpha = alpha / 10.0;
                //string tID = "5";
                //string tID = textBox1.Text;
                //string tID = strSelect;

                string tID = comboBox1.SelectedItem.ToString();
                flowLayoutPanel1.Controls.Clear();
                PictureBox ob = new PictureBox();
                ob.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
                ob.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

                ob.BackgroundImage = Image.FromFile(tID.ToString() + ".bmp");

                ob.Size = new System.Drawing.Size(75, 100);


                flowLayoutPanel1.Controls.Add(ob);
                this.toolTip1.SetToolTip(ob, "TrackID=" + tID);

                strCaller = "Track="+tID.ToString();

                ArrayList ar1 = new ArrayList();
                ArrayList ar2 = new ArrayList();


                NewMethod(tID, ar1, ar2);

                Image im = Image.FromFile(@"mks1.png");
                //Image im = Image.FromFile(@"myGray.png");
                Image im2 = im;
                //pictureBox1.Image = im2;
                //pictureBox1.Refresh();
                float[] px = new float[ar1.Count];
                float[] py = new float[ar2.Count];

                ConvertToFloat(ar1, px);
                ConvertToFloat(ar2, py);



                int xx = 0;

                //Image canvas = HeatMap.NET.HeatMap.GenerateHeatMap(im, px, py);
                //pictureBox2.Image = canvas;

                Image canvas = Class1.GenerateHeatMap(im, px, py, alpha, strCaller);
                pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
                pictureBox2.BackgroundImage = canvas;
            }
            catch 
            {
                MessageBox.Show("Select track id first");
            }


          
        }
       

        private static void ConvertToFloat(ArrayList ar1, float[] px)
        {
            for (int i = 0; i < ar1.Count; i++)
            {
                float val1 = 0.0F;
                string val = ar1[i].ToString();
                double val2 = Convert.ToDouble(val);
                px[i] = (float)val2;
            }
        }

        private void NewMethod(string tID, ArrayList ar1, ArrayList ar2)
        {
            SqlCeConnection conn = null;

            DataTable dt = new DataTable();

            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf; Persist Security Info=False");
                string path = @"Data Source = D:\12112015\company\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);
                conn.Open();

                //string strCommand = "SELECT Gate1  FROM Table1 where TrackID=" + tID;
                string strCommand = "SELECT Px, Py" + " FROM Table1 where TrackID=" + tID;
                //string strCommand = "SELECT Px, Py" + " FROM Table1";
                SqlCeCommand cmd = new SqlCeCommand(strCommand, conn);

                SqlCeDataReader myReader = cmd.ExecuteReader();

                ////DataTable dt = new DataTable();
                //SqlCeDataAdapter da = new SqlCeDataAdapter(strCommand, conn);
                //da.Fill(dt);

                //dataGridView1.DataSource = dt;


                while (myReader.Read())
                {
                    string myReaderData1 = (myReader[0].ToString());
                    ar1.Add(myReaderData1);
                    string myReaderData2 = (myReader[1].ToString());
                    ar2.Add(myReaderData2);
                }

                ///


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {
                //MessageBox.Show("success");

                conn.Close();
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            ArrayList ar1 = new ArrayList();
            string strCommand = "SELECT DISTINCT TrackID FROM Table1";
            requestQuery(strCommand, ar1);
            for (int i = 0; i < ar1.Count; i++)
            {
                //comboBox1.
                comboBox1.Items.Add(ar1[i].ToString());
            }
            chart1.Dock = DockStyle.None;
            chart1.Location = new System.Drawing.Point(5000, 5000);
            pictureBox2.Dock = DockStyle.Fill;

            button7.Enabled = false;
            button8.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;
            LoadFromExcellFile();
            helpProvider1.SetHelpString(button3, "display gate wise detail");
        }

        string strSelect = "";
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            strSelect = comboBox1.SelectedItem.ToString();
            //MessageBox.Show(st);
        }

        string pathDB = @"Data Source = D:\12112015\company\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
        string strSelectGate = "";
        string strCaller = "";
        private void button3_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();

            double alpha = trackBar1.Value;
            alpha = alpha / 10.0;
            //MessageBox.Show(test);
            ArrayList ar1 = new ArrayList();
           
            strSelectGate=comboBox2.SelectedItem.ToString();

            //pictureBox3.BackgroundImage = Image.FromFile(strSelectGate+".bmp");
            //string GateToInvistigate = "4";
            string GateToInvistigate = strSelectGate;
            
            string myGate = "Gate" + GateToInvistigate;
                string strCommand = "SELECT TrackID FROM Table2 WHERE (" + myGate + " = - 1)";
                requestQuery(strCommand, ar1);

                strCaller = myGate;

                
                ArrayList ar2 = new ArrayList();
                ArrayList ar3 = new ArrayList();

                for (int i = 0; i < ar1.Count; i++)
                {
                    string tID = ar1[i].ToString();
                    NewMethod(tID, ar2, ar3);
                    //string PBname = "pictureBox" + (2 + i).ToString();
                    //((PictureBox)PBname).Name
                    PictureBox ob = new PictureBox();
                    ob.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
                    ob.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                    //ob.Location = new System.Drawing.Point(6, 7);
                    ob.BackgroundImage = Image.FromFile(tID.ToString()+".bmp");
                    //ob.Size = new System.Drawing.Size(112, 100);  
                    ob.Size = new System.Drawing.Size(75, 100);
                    this.toolTip1.SetToolTip(ob, "TrackID=" + tID);
                    ob.ContextMenuStrip = contextMenuStrip2;
                    ob.Click += new System.EventHandler(handleTrackDetail);
                  
                    flowLayoutPanel1.Controls.Add(ob);
                }

                Image im = Image.FromFile(@"mks1.png");
                //Image im = Image.FromFile(@"myGray.png");
                Image im2 = im;
                //pictureBox1.Image = im2;
                //pictureBox1.Refresh();
                float[] px = new float[ar2.Count];
                float[] py = new float[ar3.Count];

                ConvertToFloat(ar2, px);
                ConvertToFloat(ar3, py);



                int xx = 0;

                //int indexOfSelectedMarker = comboBox3.SelectedIndex;
                //Image marker = imageList2.Images[indexOfSelectedMarker];

                //Image canvas = Class1.GenerateHeatMap(im, px, py, alpha, marker);
                Image canvas = Class1.GenerateHeatMap(im, px, py, alpha, strCaller);
                pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
                pictureBox2.BackgroundImage = canvas;
                //pictureBox2.Image = canvas;
                //pictureBox2.Refresh();
               
          
        }

        private void requestQuery(string strCommand, ArrayList ar1)
        {
            SqlCeConnection conn = null;

            DataTable dt = new DataTable();

            try
            {

                conn = new SqlCeConnection(pathDB);
                conn.Open();

                SqlCeCommand cmd = new SqlCeCommand(strCommand, conn);

                SqlCeDataReader myReader = cmd.ExecuteReader();


                //SqlCeDataAdapter da = new SqlCeDataAdapter(strCommand, conn);
                //da.Fill(dt);

                //dataGridView1.DataSource = dt;


                while (myReader.Read())
                {
                    string myReaderData1 = (myReader[0].ToString());
                    ar1.Add(myReaderData1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {

                conn.Close();
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            double alpha = trackBar1.Value;
            alpha = alpha / 10.0;
            //MessageBox.Show(test);
            ArrayList ar1 = new ArrayList();

            strSelectGate = comboBox2.SelectedItem.ToString();

            //pictureBox3.BackgroundImage = Image.FromFile(strSelectGate + ".bmp");
            //string GateToInvistigate = "4";
            string GateToInvistigate = strSelectGate;

            string myGate = "Gate" + GateToInvistigate;
            string strCommand = "SELECT TrackID FROM Table2 WHERE (" + myGate + " = - 1)";
            requestQuery(strCommand, ar1);



            ArrayList ar2 = new ArrayList();
            ArrayList ar3 = new ArrayList();

            for (int i = 0; i < ar1.Count; i++)
            {
                string tID = ar1[i].ToString();
                NewMethod(tID, ar2, ar3);
            }

            Image im = Image.FromFile(@"mks1.png");
            //Image im = Image.FromFile(@"myGray.png");
            Image im2 = im;
            //pictureBox1.Image = im2;
            //pictureBox1.Refresh();
            float[] px = new float[ar2.Count];
            float[] py = new float[ar3.Count];

            ConvertToFloat(ar2, px);
            ConvertToFloat(ar3, py);



            int xx = 0;

            //Image canvas = HeatMap.NET.HeatMap.GenerateHeatMap(im, px, py);

            //int indexOfSelectedMarker = comboBox3.SelectedIndex;
            //Image marker = imageList2.Images[indexOfSelectedMarker];
 
            //Image canvas = Class1.GenerateHeatMap(im, px, py, alpha,marker);

            int indexOfSelectedMarker = comboBox3.SelectedIndex;
            Image marker = imageList2.Images[indexOfSelectedMarker];

            Image canvas = Class1.GenerateHeatMap(im, px, py, alpha, strCaller);


            pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
            pictureBox2.BackgroundImage = canvas;
            //pictureBox2.Image = canvas;
            //pictureBox2.Refresh();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            double alpha = trackBar1.Value;
            alpha = alpha / 10.0;
            //MessageBox.Show(test);
            ArrayList ar1 = new ArrayList();

            //strSelectGate = comboBox2.SelectedItem.ToString();

            //pictureBox3.BackgroundImage = Image.FromFile(strSelectGate + ".bmp");
            //string GateToInvistigate = "4";
            //string GateToInvistigate = strSelectGate;

            //string myGate = "Gate" + GateToInvistigate;

            //double alpha = 0.5;
            string strCommand = "SELECT TrackID FROM Table2";
            requestQuery(strCommand, ar1);

            strCaller = "All track";

            ArrayList ar2 = new ArrayList();
            ArrayList ar3 = new ArrayList();

            for (int i = 0; i < ar1.Count; i++)
            {
                string tID = ar1[i].ToString();
                NewMethod(tID, ar2, ar3);
            }

            Image im = Image.FromFile(@"mks1.png");
            //Image im = Image.FromFile(@"myGray.png");
            Image im2 = im;
            //pictureBox1.Image = im2;
            //pictureBox1.Refresh();
            float[] px = new float[ar2.Count];
            float[] py = new float[ar3.Count];

            ConvertToFloat(ar2, px);
            ConvertToFloat(ar3, py);



            int xx = 0;

            //Image canvas = HeatMap.NET.HeatMap.GenerateHeatMap(im, px, py);

            //int indexOfSelectedMarker = comboBox3.SelectedIndex;
            //Image marker = imageList2.Images[indexOfSelectedMarker];

            //Image canvas = Class1.GenerateHeatMap(im, px, py, alpha,marker);

            //int indexOfSelectedMarker = comboBox3.SelectedIndex;
            //Image marker = imageList2.Images[indexOfSelectedMarker];

            Image canvas = Class1.GenerateHeatMap(im, px, py, alpha, strCaller);


            pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
            pictureBox2.BackgroundImage = canvas;
            //pictureBox2.Image = canvas;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            //string fname = DateTime.Now.Ticks.ToString() + ".png";
            //pictureBox2.BackgroundImage.Save(fname);
            //MessageBox.Show("Done");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
            //pictureBox2.Dock = DockStyle.None;
            //pictureBox2.Location = new System.Drawing.Point(5000, 5000);
            //chart1.Dock = DockStyle.Fill;

            ArrayList ar1 = new ArrayList();

            strSelectGate = comboBox2.SelectedItem.ToString();

            //pictureBox3.BackgroundImage = Image.FromFile(strSelectGate+".bmp");
            //string GateToInvistigate = "4";
            string GateToInvistigate = strSelectGate;

            string myGate = "Gate" + GateToInvistigate;
            string strCommand = "SELECT TrackID FROM Table2 WHERE (" + myGate + " = - 1)";
            requestQuery(strCommand, ar1);

            int len = chart1.Series.Count;
            if (len > 0)
            {
                chart1.Series.Clear();
                //chart1.Series.RemoveAt(0);
                chart1.Series.Add(myGate);

            }
            else
            {
                chart1.Series.Add(myGate);
            }

          

            //for (int i = 0; i < ar1.Count; i++)
            int ind1=comboBox1.Items.Count;
            int val1= Convert.ToInt32( comboBox1.Items[ind1-1]);
            int[] seq1 = new int[val1];
            int[] seq2 = new int[val1];
            for (int i = 0; i < val1; i++)
            {
                //ar.Add(i+1);
                seq1[i] = i ;
              
            }
            for (int i = 0; i < ar1.Count; i++)
            {
                int val2=Convert.ToInt32(ar1[i]);
                seq2[val2] = val2;
 
            }



            for (int i = 0; i < val1; i++)
            {
                //seq1[i] = i + 1;
                //chart1.Series[myGate].Points.AddY(ar1[i]);
                //chart1.Series[myGate].Points.AddXY(ar1[i], seq1[i]);
                chart1.Series[myGate].Points.AddXY(seq1[i], seq2[i]);
            }

            
            chart1.Series[myGate].ChartType = SeriesChartType.Column;
            chart1.Series[myGate].Color = Color.Red;


            chart1.Series[myGate].IsValueShownAsLabel = true;

        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {

            if (checkBox1.Checked == true)
            {
                pictureBox2.Dock = DockStyle.None;
                pictureBox2.Location = new System.Drawing.Point(5000, 5000);
                chart1.Dock = DockStyle.Fill;
                button7.Enabled = true;
                button8.Enabled = true;
                button9.Enabled = true;
                button10.Enabled = true;

            }
            else
            {
               
               
                chart1.Dock = DockStyle.None;
                chart1.Location = new System.Drawing.Point(5000, 5000);
                pictureBox2.Dock = DockStyle.Fill;

                button7.Enabled = false;
                button8.Enabled = false;
                button9.Enabled = false;
                button10.Enabled = false;
 
            }
        }

        //private void TrackCrossingGate(string tID, ArrayList ar, string myGate)
        //{
        //    SqlCeConnection conn = null;

        //    DataTable dt = new DataTable();

        //    try
        //    {
        //        //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf; Persist Security Info=False");
        //        //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
        //        conn = new SqlCeConnection(path);
        //        conn.Open();

        //        //string strCommand = "SELECT Gate1  FROM Table1 where TrackID=" + tID;
        //        string strCommand = "SELECT " + myGate + " FROM Table1 where TrackID=" + tID;

        //        SqlCeCommand cmd = new SqlCeCommand(strCommand, conn);

        //        SqlCeDataReader myReader = cmd.ExecuteReader();

        //        //DataTable dt = new DataTable();
        //        //SqlCeDataAdapter da = new SqlCeDataAdapter(strCommand, conn);
        //        //da.Fill(dt);
        //        ////return dt;
        //        //dataGridView1.DataSource = dt;


        //        while (myReader.Read())
        //        {
        //            string myReaderData = (myReader[0].ToString());
        //            ar.Add(myReaderData);
        //        }

        //        ///


        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error");
        //    }
        //    finally
        //    {
        //        //MessageBox.Show("success");

        //        conn.Close();
        //    }
        //}

        private void button8_Click(object sender, EventArgs e)
        {
            ArrayList ar = new ArrayList();
            string GateToInvistigate = "";
            int numGate = comboBox2.Items.Count;

            for (int i = 1; i <= numGate; i++)
            {
                GateToInvistigate = i.ToString();


                ExtractGateCrossingInfo(ar, GateToInvistigate);
            }

            int len = chart1.Series.Count;

            string myGate = "Frequency of gate used";
            if (len > 0)
            {
                chart1.Series.Clear();
                //chart1.Series.RemoveAt(0);
                chart1.Series.Add(myGate);

            }
            else
            {
                chart1.Series.Add(myGate);
            }
            for (int i = 0; i < ar.Count; i++)
            {

                chart1.Series[myGate].Points.AddY(ar[i]);

            }


            chart1.Series[myGate].ChartType = SeriesChartType.Column;
            chart1.Series[myGate].Color = Color.Blue;



            chart1.Series[myGate].IsValueShownAsLabel = true;
        }
        private void ExtractGateCrossingInfo(ArrayList ar, string GateToInvistigate)
        {
            string myGate = "Gate" + GateToInvistigate;
            SqlCeConnection conn = null;

            DataTable dt = new DataTable();

            try
            {

                conn = new SqlCeConnection(path);
                conn.Open();


                //string strCommand = "SELECT " + myGate + " FROM Table1 where TrackID=" + tID;
                string strCommand = "SELECT COUNT(" + myGate + ") FROM Table2 WHERE (" + myGate + " = - 1)";

                SqlCeCommand cmd = new SqlCeCommand(strCommand, conn);

                SqlCeDataReader myReader = cmd.ExecuteReader();


                SqlCeDataAdapter da = new SqlCeDataAdapter(strCommand, conn);
                da.Fill(dt);

                dataGridView1.DataSource = dt;


                while (myReader.Read())
                {
                    string myReaderData = (myReader[0].ToString());
                    ar.Add(myReaderData);
                }

                ///


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {
                //MessageBox.Show("success");

                conn.Close();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ArrayList ar1 = new ArrayList();

            strSelectGate = comboBox2.SelectedItem.ToString();

            //pictureBox3.BackgroundImage = Image.FromFile(strSelectGate+".bmp");
            //string GateToInvistigate = "4";
            string GateToInvistigate = strSelectGate;

            string myGate = "Gate" + GateToInvistigate;

            string chartText = "Trajectory length";
            //string strCommand = "SELECT TrackID FROM Table2 WHERE (" + myGate + " = - 1)";
            string strCommand = " SELECT COUNT(TrackID) FROM Table1 GROUP BY TrackID";
           
            requestQuery(strCommand, ar1);

            int len = chart1.Series.Count;
            if (len > 0)
            {
                chart1.Series.Clear();
                //chart1.Series.RemoveAt(0);
                chart1.Series.Add(chartText);

            }
            else
            {
                chart1.Series.Add(chartText);
            }



            //for (int i = 0; i < ar1.Count; i++)
            int ind1 = comboBox1.Items.Count;
            int val1 = Convert.ToInt32(comboBox1.Items[ind1 - 1]);
            int[] seq1 = new int[val1+1];
            int[] seq2 = new int[val1+1];
            for (int i = 1; i <= val1; i++)
            {
                //ar.Add(i+1);
                seq1[i] = i;

            }

            ArrayList ar2 = new ArrayList();
             strCommand = "SELECT DISTINCT TrackID FROM Table1";
            requestQuery(strCommand, ar2);
            //for (int i = 0; i < ar1.Count; i++)
            //{
            //    int val2 = Convert.ToInt32(ar1[i]);
            //    seq2[i] = val2;

            //}
            int cnt=0;
            foreach (object i in ar2)
            {
                int va = Convert.ToInt32(i);
                seq2[va] = Convert.ToInt32(ar1[cnt]);
                cnt++;
            }



            for (int i = 1; i <= val1; i++)
            {
                //seq1[i] = i + 1;
                //chart1.Series[myGate].Points.AddY(ar1[i]);
                //chart1.Series[myGate].Points.AddXY(ar1[i], seq1[i]);
                chart1.Series[chartText].Points.AddXY(seq1[i], seq2[i]);
            }


            chart1.Series[chartText].ChartType = SeriesChartType.Column;
            chart1.Series[chartText].Color = Color.CadetBlue;


            chart1.Series[chartText].IsValueShownAsLabel = true;
        }

        string path2 = @"D:\12112015\company\company\PeopleCount.xls";
        object[,] valueArray;
        private void LoadFromExcellFile()
        {
            //string path1 = @"D:\12112015\company\company\myExcell.xls";
            // Reference to Excel Application.
            Excel.Application xlApp = new Excel.Application();

            // Open the Excel file.
            // You have pass the full path of the file.
            // In this case file is stored in the Bin/Debug application directory.
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(path2));

            // Get the first worksheet.
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);

            // Get the range of cells which has data.
            Excel.Range xlRange = xlWorksheet.UsedRange;

            // Get an object array of all of the cells in the worksheet with their values.
            //object[,] valueArray = (object[,])xlRange.get_Value(
            //            Excel.XlRangeValueDataType.xlRangeValueDefault);
            valueArray = (object[,])xlRange.get_Value(
                    Excel.XlRangeValueDataType.xlRangeValueDefault);

            // iterate through each cell and display the contents.
            //for (int row = 1; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
            //{
            //    for (int col = 1; col <= xlWorksheet.UsedRange.Columns.Count; ++col)
            //    {
            //        // Print value of the cell to Console.
            //        Console.WriteLine(valueArray[row, col].ToString());
            //    }
            //}

            // Close the Workbook.
            xlWorkbook.Close(false);

            // Relase COM Object by decrementing the reference count.
            Marshal.ReleaseComObject(xlWorkbook);

            // Close Excel application.
            xlApp.Quit();

            // Release COM object.
            Marshal.FinalReleaseComObject(xlApp);

            Console.ReadLine();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //LoadFromExcellFile();

            int val1 = valueArray.GetLength(0);
            int[] FrameNumbers = new int[val1];
            int[] PeopleCount = new int[val1];

            //int[] num=valueArray(:,9);
            for (int i = 1; i <= val1; i++)
            {
                FrameNumbers[i - 1] = Convert.ToInt32(valueArray[i, 1]);
                PeopleCount[i - 1] = Convert.ToInt32(valueArray[i, 2]);
            }

            string strDisp = "People Count";
            int len = chart1.Series.Count;
            if (len > 0)
            {
                chart1.Series.Clear();
                chart1.Series.Add(strDisp);

            }
            else
            {
                chart1.Series.Add(strDisp);
            }



            for (int i = 0; i < val1; i++)
            {

                chart1.Series[strDisp].Points.AddXY(FrameNumbers[i], PeopleCount[i]);
            }


            chart1.Series[strDisp].ChartType = SeriesChartType.Column;
            chart1.Series[strDisp].Color = Color.Chartreuse;


            //chart1.Series[strDisp].IsValueShownAsLabel = true;
            //MessageBox.Show("done");
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void aboutUsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form ob = new AboutBox1();
            ob.ShowDialog();
        }
        private void handleTrackDetail(object sender, EventArgs e)
        {
            string tID = this.toolTip1.GetToolTip((PictureBox)sender);
            strSetID = tID.Substring(tID.IndexOf('=')+1);
 
        }
        string strSetID = "";

        private void displayTrackDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                double alpha = trackBar1.Value;
                alpha = alpha / 10.0;
                //string tID = "5";
                //string tID = textBox1.Text;
                //string tID = strSelect;
                //PictureBox pt = ((PictureBox)sender);
                string tID = strSetID;
                //string tID = comboBox1.SelectedItem.ToString();
               
                //string tID =  this.toolTip1.GetToolTip((PictureBox)sender);
                flowLayoutPanel1.Controls.Clear();
                PictureBox ob = new PictureBox();
                ob.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
                ob.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

                ob.BackgroundImage = Image.FromFile(tID.ToString() + ".bmp");

                ob.Size = new System.Drawing.Size(75, 100);


                flowLayoutPanel1.Controls.Add(ob);
                //this.toolTip1.SetToolTip(ob, "TrackID=" + tID);

                strCaller = "Track=" + tID.ToString();

                ArrayList ar1 = new ArrayList();
                ArrayList ar2 = new ArrayList();


                NewMethod(tID, ar1, ar2);

                Image im = Image.FromFile(@"mks1.png");
                //Image im = Image.FromFile(@"myGray.png");
                Image im2 = im;
                //pictureBox1.Image = im2;
                //pictureBox1.Refresh();
                float[] px = new float[ar1.Count];
                float[] py = new float[ar2.Count];

                ConvertToFloat(ar1, px);
                ConvertToFloat(ar2, py);



                int xx = 0;

                //Image canvas = HeatMap.NET.HeatMap.GenerateHeatMap(im, px, py);
                //pictureBox2.Image = canvas;

                Image canvas = Class1.GenerateHeatMap(im, px, py, alpha, strCaller);
                pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
                pictureBox2.BackgroundImage = canvas;
            }
            catch(Exception ee)
            {
                MessageBox.Show(ee.Message.ToString());
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string fname = @"outputImg\"+ DateTime.Now.Ticks.ToString() + ".png";
            pictureBox2.BackgroundImage.Save(fname);

            
            MessageBox.Show("Done");
        }

        private void quickSaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fname = @"outputImg\" + DateTime.Now.Ticks.ToString() + ".png";
            chart1.Series[0].AxisLabel = "Demo copy";

            chart1.SaveImage(fname, ChartImageFormat.Png);

            MessageBox.Show("Done");
        }

        
    }
}
