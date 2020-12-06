using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;
using System.Data.SqlServerCe;



using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;

//using Excel = Microsoft.Office.Interop.Excel;


namespace VideoAnalytics4StoreLines
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        object[,] valueArray;
        string path1 = @"D:\12112015\company\company\myExcell.xls";
        string path = @"Data Source = D:\12112015\company\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";


        private void button1_Click(object sender, EventArgs e)
        {
            //string path = @"D:\companywork\project\project\company\myExcell.xls";

            //MyApp = new Excel.Application();
            //MyApp.Visible = false;
            //MyBook = MyApp.Workbooks.Open(path);
            //MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            //lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            //BindingList<Employee> EmpList = new BindingList<Employee>();
            //for (int index = 2; index <= lastRow; index++)
            //{
            //    System.Array MyValues = (System.Array)MySheet.get_Range("A" +
            //       index.ToString(), "D" + index.ToString()).Cells.Value;
            //    EmpList.Add(new Employee
            //    {
            //        Name = MyValues.GetValue(1, 1).ToString(),
            //        Employee_ID = MyValues.GetValue(1, 2).ToString(),
            //        Email_ID = MyValues.GetValue(1, 3).ToString(),
            //        Number = MyValues.GetValue(1, 4).ToString()
            //    });
            //}

            // Load Excel file.
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path1);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[,] allData = new string[rowCount+1, colCount+1];

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount-2; j++)
                {
                    //MessageBox.Show(xlRange.Cells[i, j].Value2.ToString());
                    allData[i, j] = xlRange.Cells[i, j].Value2.ToString();
                    //Console.Write(xlRange.Cells[i, j].Value2.ToString());
                }
                //Console.WriteLine(" ");
            }
            MessageBox.Show("done");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            LoadFromExcellFile();
            MessageBox.Show("done");
        }

        private void LoadFromExcellFile()
        {
            //string path1 = @"D:\12112015\company\company\myExcell.xls";
            // Reference to Excel Application.
            Excel.Application xlApp = new Excel.Application();

            // Open the Excel file.
            // You have pass the full path of the file.
            // In this case file is stored in the Bin/Debug application directory.
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(path1));

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

        private void button3_Click(object sender, EventArgs e)
        {
            //valueArray
            Random rdn = new Random();
            for (int i = 0; i < 50; i++)
            {
                //chart1.Series["test1"].Points.AddXY(rdn.Next(0, 10), rdn.Next(0, 10));
                chart1.Series["test2"].Points.AddXY
                                (rdn.Next(0, 10), rdn.Next(0, 10));
            }

            chart1.Series["test1"].ChartType = SeriesChartType.FastLine;
            chart1.Series["test1"].Color = Color.Red;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            int[] numbers = { 5, 4, 1, 3, 9, 8, 6, 7, 2, 0 };

            var numberGroups =
                from n in numbers
                group n by n % 5 into g
                select new { Remainder = g.Key, Numbers = g };

            foreach (var g in numberGroups)
            {
                Console.WriteLine("Numbers with a remainder of {0} when divided by 5:", g.Remainder);
                foreach (var n in g.Numbers)
                {
                    Console.WriteLine(n);
                }
            } 
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //int[] numbers = { 5, 4, 1, 3, 9, 8, 6, 7, 2, 0, 4, 3, 2 };

            int val1= valueArray.GetLength(0);
            int[] numbers = new int[val1];

            //int[] num=valueArray(:,9);
            for (int i = 1; i <= val1; i++)
            {
                numbers[i-1] = Convert.ToInt32(valueArray[i, 9]);
            }


            var numberGroups =
                from n in numbers
                group n by n into g
                select new { Remainder = g.Key, Numbers = g };
            ArrayList arr = new ArrayList();
            
            foreach (var g in numberGroups)
            {
                Console.WriteLine("Numbers :", g.Remainder);
                Console.WriteLine(g.Remainder);
                //foreach (var n in g.Numbers)
                //{
                //    //Console.WriteLine(n);
                //}
                arr.Add(g.Remainder);
            }

            int[] seq1 = new int[arr.Count];
            ArrayList ar = new ArrayList();

            for (int i = 0; i < arr.Count; i++)
            {
                //ar.Add(i+1);
                seq1[i] = i + 1;
            }

            int[] seq2 = arr.ToArray(typeof(int)) as int[];
            chart1.Series.Add("test1");
            chart1.Series.Add("test2");

            for (int i = 0; i < arr.Count; i++)
            {
                //chart1.Series["test1"].Points.AddXY(seq1[i], seq2[i]);
                //chart1.Series["test2"].Points.AddXY(seq1[i], seq2[i]);
                chart1.Series["test1"].Points.AddXY(seq2[i], seq1[i]);
                chart1.Series["test2"].Points.AddXY(seq2[i], seq1[i]);
                //chart1.Series["test2"].Points.AddY(seq2[i]);
                //chart1.Series[0].IsValueShownAsLabel = true;
                comboBox1.Items.Add(arr[i]);
               
               
            }

            chart1.Series["test1"].ChartType = SeriesChartType.FastLine;
            chart1.Series["test1"].Color = Color.Red;
            //chart1.Series["test2"].ChartType = SeriesChartType.FastLine;
            chart1.Series["test2"].Color = Color.Green;
            //chart1.Series["test1"].Label = "sdfs";
            chart1.Series["test1"].XValueMember = "Account Name";
            chart1.Series["test1"].YValueMembers = "Days Used";
            chart1.Series["test1"].IsValueShownAsLabel = true;
            label1.Text = arr.Count.ToString();
          
         
           
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SqlCeConnection conn = null;
            int r = 0;
            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf;");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);

               
                conn.Open();

                //SqlCeCommand cmd = conn.CreateCommand();
                //cmd.CommandText = "INSERT INTO Table1 ([name], [age]) VALUES('abc', '123')";
                //cmd.CommandText = "INSERT INTO Table1 ([Gate1], [Gate2],[Gate3],[Gate4],[TrackID],[Px],[Py]) VALUES('1',1,1,1,,2,'11.2','14.2')";
                
                //cmd.CommandText = "INSERT INTO Table1 ([Gate1], [Gate2],[Gate3],[Gate4],[TrackID],[Px],[Py]) VALUES('1','1','1','1','2','11.2','14.2')";


                SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Table1 ([Gate1], [Gate2],[Gate3],[Gate4],[TrackID],[Px],[Py]) VALUES (@0, @1, @2, @3,@4,@5,@6)", conn);

                // In the command, there are some parameters denoted by @, you can 
                // change their value on a condition, in my code they're hardcoded.

                cmd.Parameters.Add(new SqlCeParameter("0", 1));
                cmd.Parameters.Add(new SqlCeParameter("1",1));
                cmd.Parameters.Add(new SqlCeParameter("2", 1));
                cmd.Parameters.Add(new SqlCeParameter("3", -1));
                 cmd.Parameters.Add(new SqlCeParameter("4", 1));
                 cmd.Parameters.Add(new SqlCeParameter("5", "13.1"));
                 cmd.Parameters.Add(new SqlCeParameter("6", "16.2"));

                r = cmd.ExecuteNonQuery();
            }
            finally
            {
                conn.Close();
            }
            MessageBox.Show(r.ToString());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ////string source = @"Server=(local);" + "integrated security=SSPI;" + "database=Database1.sdf";
            //string source = @"database=|DataDirectory|\Database1.sdf";
            //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
            //conn = new SqlCeConnection(path);

            string select = "select * from table1";
            SqlCeConnection conn = new SqlCeConnection(path);
            conn.Open();
            SqlCeCommand cmd = new SqlCeCommand(select, conn);
            SqlCeDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                Console.WriteLine("val1= {0,-20} val2: {1}", reader[0], reader[1]);
            }

            //using (SqlConnection con = new SqlConnection(ConnectionStrings[@"Data Source=|DataDirectory|\Database1.sdf"].ConnectionStrings)) ;


        }

        private void button8_Click(object sender, EventArgs e)
        {
           

            SqlCeConnection conn = null;

            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf; Persist Security Info=False");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);
                conn.Open();

                SqlCeCommand cmd = conn.CreateCommand();
                //cmd.CommandText = "SELECT * FROM Table1";
                //sqlcedatawr 

                //SqlCeDataReader myReader = null;
                //SqlCeDataReader myReader = cmd.ExecuteReader();
                string strCommand = "SELECT * FROM Table1";
                DataTable dt = new DataTable();
                SqlCeDataAdapter da = new SqlCeDataAdapter(strCommand, conn);
                da.Fill(dt);
                //return dt;
                dataGridView1.DataSource = dt;




                //while (myReader.Read())
                //{
                //    string myReaderData = (myReader[1].ToString());


                //    string myColumn3 = (myReader["Column3"].ToString());

                //    //string myColumn4 = (myReader["Column4"].ToString());

                //    //string myColumn5 = (myReader["Column5"].ToString());

                //    //string myColumn6 = (myReader["Column6"].ToString());

                //    ////string bodyMsg = "Hi " + myColumn2 + ", your MOT on vehicle " + myColumn3 + " reg number '" + myColumn4 +
                //    //"' is due to expire on " + myColumn5.Remove(10, 9) + ". Please call us to book an appointment";


                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {
                MessageBox.Show("success");

                conn.Close();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

            SqlCeConnection conn = null;
            int r = 0;
            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf;");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);


                conn.Open();

                for (int i = 1; i <= valueArray.GetLength(0); i++)
                {
                    SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Table1 ([Gate1], [Gate2],[Gate3],[Gate4],[TrackID],[Px],[Py]) VALUES (@0, @1, @2, @3,@4,@5,@6)", conn);
                    cmd.Parameters.Add(new SqlCeParameter("0", valueArray[i,2]));
                    cmd.Parameters.Add(new SqlCeParameter("1", valueArray[i, 4]));
                    cmd.Parameters.Add(new SqlCeParameter("2", valueArray[i, 6]));
                    cmd.Parameters.Add(new SqlCeParameter("3", valueArray[i, 8]));
                    cmd.Parameters.Add(new SqlCeParameter("4", valueArray[i, 9]));
                    cmd.Parameters.Add(new SqlCeParameter("5", valueArray[i, 10].ToString()));
                    cmd.Parameters.Add(new SqlCeParameter("6", valueArray[i, 11].ToString()));

                    // In the command, there are some parameters denoted by @, you can 
                    // change their value on a condition, in my code they're hardcoded.

                    //cmd.Parameters.Add(new SqlCeParameter("0", 1));
                    //cmd.Parameters.Add(new SqlCeParameter("1", 1));
                    //cmd.Parameters.Add(new SqlCeParameter("2", 1));
                    //cmd.Parameters.Add(new SqlCeParameter("3", -1));
                    //cmd.Parameters.Add(new SqlCeParameter("4", 1));
                    //cmd.Parameters.Add(new SqlCeParameter("5", "13.1"));
                    //cmd.Parameters.Add(new SqlCeParameter("6", "16.2"));

                    r = cmd.ExecuteNonQuery();
                }
            }
            finally
            {
                conn.Close();
            }
            MessageBox.Show(r.ToString());
        }

        private void button10_Click(object sender, EventArgs e)
        {
            SqlCeConnection conn = null;
            int r = 0;
            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf;");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);


                conn.Open();

                //SqlCeCommand cmd = new SqlCeCommand("DELETE Gate1, Gate2,Gate3,Gate4,TrackID,Px,Py FROM Table1)", conn);
                SqlCeCommand cmd = new SqlCeCommand("DELETE * FROM Table1)", conn);
                   
                    r=cmd.ExecuteNonQuery();
              
            }
            finally
            {
                conn.Close();
            }
            MessageBox.Show(r.ToString());
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string tID = textBox2.Text;
            ArrayList ar = new ArrayList();
            string GateToInvistigate = "1";
            string myGate = "Gate" + GateToInvistigate;
            //string myGate = "Gate2";

            TrackCrossingGate(tID, ar, myGate);

            int cnt = ar.Count;
            int output = 0;
            int outputAt = 0;
            for (int i = 1; i < cnt; i++)
            {
                output = Convert.ToInt32(ar[i - 1].ToString()) * Convert.ToInt32(ar[i].ToString());
                if ((output == -1)||(output == 0))
                {
                    outputAt = i;
                    break;
                }
            }


            //
            string myStr= "output=\t"+output+"\tmyGate=\t"+myGate+"\ttID=\t"+tID+"\r\n";
            richTextBox1.AppendText(myStr);

           
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
            for (int i = 0; i < ar.Count; i++)
            {
               
                chart1.Series[myGate].Points.AddY(ar[i]);

            }
          

            chart1.Series[myGate].ChartType = SeriesChartType.FastLine;
            chart1.Series[myGate].Color = Color.Red;


            chart1.Series[myGate].IsValueShownAsLabel = true;
          

        }

        private void TrackCrossingGate(string tID, ArrayList ar, string myGate)
        {
            SqlCeConnection conn = null;

            DataTable dt = new DataTable();

            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf; Persist Security Info=False");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);
                conn.Open();

                //string strCommand = "SELECT Gate1  FROM Table1 where TrackID=" + tID;
                string strCommand = "SELECT " + myGate+  " FROM Table1 where TrackID=" + tID;
                               
                SqlCeCommand cmd = new SqlCeCommand(strCommand, conn);

                SqlCeDataReader myReader = cmd.ExecuteReader();

                //DataTable dt = new DataTable();
                SqlCeDataAdapter da = new SqlCeDataAdapter(strCommand, conn);
                da.Fill(dt);
                //return dt;
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

        private void button12_Click(object sender, EventArgs e)
        {
            //string tID = textBox2.Text;
            string tID = "";
            string GateToInvistigate = "";
            //string GateToInvistigate = "1";
            SqlCeConnection conn = null;
            int r = 0;
            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf;");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);


                conn.Open();

                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    tID = comboBox1.Items[i].ToString();

                    int output = 0;
                    string strOutputColl = "";
                    ArrayList arr1 = new ArrayList();

                    for (int j = 0; j < comboBox2.Items.Count; j++)
                    {
                        GateToInvistigate = comboBox2.Items[j].ToString();
                        NewMethod(tID, GateToInvistigate, ref output);
                        strOutputColl += output + "\t";
                        arr1.Add(output);


                    }
                    //string myStr = "output=\t" + output + "\tmyGate=\t" + myGate + "\ttID=\t" + tID + "\r\n";
                    string myStr = "ID=\t" + tID + "\toutput=\t" + strOutputColl + "\r\n";
                    richTextBox1.AppendText(myStr);

                    SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Table2 ([TrackID],[Gate1], [Gate2],[Gate3],[Gate4]) VALUES (@0, @1, @2, @3,@4)", conn);
                    cmd.Parameters.Add(new SqlCeParameter("0", tID));
                    cmd.Parameters.Add(new SqlCeParameter("1", arr1[0].ToString()));
                    cmd.Parameters.Add(new SqlCeParameter("2", arr1[1].ToString()));
                    cmd.Parameters.Add(new SqlCeParameter("3", arr1[2].ToString()));
                    cmd.Parameters.Add(new SqlCeParameter("4", arr1[3].ToString()));



                    r = cmd.ExecuteNonQuery();
                }
            }

            finally
            {
                conn.Close();
            }
            MessageBox.Show(r.ToString());

            //NewMethod1();

           
        }

        private void NewMethod1()
        {
            SqlCeConnection conn = null;
            int r = 0;
            try
            {
                //conn = new SqlCeConnection(@"Data Source = |DataDirectory|\Database1.sdf;");
                //string path = @"Data Source = D:\companywork\project\project\company\VideoAnalytics4StoreLines\VideoAnalytics4StoreLines\Database1.sdf";
                conn = new SqlCeConnection(path);


                conn.Open();

                for (int i = 1; i <= valueArray.GetLength(0); i++)
                {
                    SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Table1 ([TrackID],[Gate1], [Gate2],[Gate3],[Gate4]) VALUES (@0, @1, @2, @3,@4,@5,@6)", conn);
                    cmd.Parameters.Add(new SqlCeParameter("0", valueArray[i, 2]));
                    cmd.Parameters.Add(new SqlCeParameter("1", valueArray[i, 4]));
                    cmd.Parameters.Add(new SqlCeParameter("2", valueArray[i, 6]));
                    cmd.Parameters.Add(new SqlCeParameter("3", valueArray[i, 8]));
                    cmd.Parameters.Add(new SqlCeParameter("4", valueArray[i, 9]));
       


                    r = cmd.ExecuteNonQuery();
                }
            }
            finally
            {
                conn.Close();
            }
            MessageBox.Show(r.ToString());
        }

        private void NewMethod(string tID, string GateToInvistigate, ref int output)
        {
            ArrayList ar = new ArrayList();
            string myGate = "Gate" + GateToInvistigate;
            //string myGate = "Gate2";

            TrackCrossingGate(tID, ar, myGate);

            int cnt = ar.Count;
            //int output = 0;
            int outputAt = 0;
            for (int i = 1; i < cnt; i++)
            {
                output = Convert.ToInt32(ar[i - 1].ToString()) * Convert.ToInt32(ar[i].ToString());
                if ((output == -1) || (output == 0))
                {
                    outputAt = i;
                    break;
                }
            }


            //
            //string myStr = "output=\t" + output + "\tmyGate=\t" + myGate + "\ttID=\t" + tID + "\r\n";
            //richTextBox1.AppendText(myStr);


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
            for (int i = 0; i < ar.Count; i++)
            {

                chart1.Series[myGate].Points.AddY(ar[i]);

            }


            chart1.Series[myGate].ChartType = SeriesChartType.FastLine;
            chart1.Series[myGate].Color = Color.Red;


            chart1.Series[myGate].IsValueShownAsLabel = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Form obj = new Form2();
            obj.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            ArrayList ar = new ArrayList();
            string GateToInvistigate = "";

            for (int i = 1; i <= 4;i++ )
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
            chart1.Series[myGate].Color = Color.Red;
            


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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
