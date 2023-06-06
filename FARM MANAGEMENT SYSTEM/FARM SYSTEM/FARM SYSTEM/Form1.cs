using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace FARM_SYSTEM
{
    public partial class Form1 : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        public Form1()
        {
            InitializeComponent();
            connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/FARM MANAGEMENT SYSTEM/FARM SYSTEM/FARMDB.accdb";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                if (textBox3.Text.Equals("") || textBox3.Text.Equals("1"))
                {
                    MessageBox.Show("Input Item name!!!", "MSG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;

                }



                if (textBox2.Text.Equals("") || textBox2.Text.Equals("1"))
                {
                    MessageBox.Show("Input location where the item will be stored!!!", "Msg", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;

                }


                if (comboBox3.Text.Equals("") || comboBox3.Text.Equals("1"))
                {
                    MessageBox.Show("Input item weight type in kgs, tonnes etc!!!", "Msg", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;

                }

                if (textBox5.Text.Equals("") || textBox5.Text.Equals("Password"))
                {
                    MessageBox.Show("Input item actual weight!!!", "Msg", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;

                }

                if (textBox1.Text.Equals("") || textBox1.Text.Equals("Password"))
                {
                    MessageBox.Show("Input item quantity(the number of items)!!!", "Msg", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;

                }


                string ITEM = Convert.ToString(textBox3.Text.Trim());
                string LOCATION = Convert.ToString(textBox2.Text.Trim());
                string WEIGHTTYPE = Convert.ToString(comboBox3.Text.Trim());
                double WEIGHT = 0.0;
                WEIGHT = double.Parse(textBox5.Text.Trim());
                WEIGHT = Math.Round(WEIGHT, 2);
                double QUANTITY = 0.0;
                QUANTITY = double.Parse(textBox1.Text.Trim());
                QUANTITY = Math.Round(QUANTITY, 2);

       
              
                    string date = DateTime.Now.ToString("MM/dd/yyyy");
                    connection.Close();
                    connection.Open();
                    OleDbCommand commandinsert = new OleDbCommand();
                    commandinsert.Connection = connection;
                    commandinsert.CommandText = "INSERT INTO [warehousetb] ([DATE],[ITEM], [LOCATION], [QUANTITY], [UNIT WEIGHT], [WEIGHT TYPE] ) VALUES ('" + date + "','" + ITEM + "','" + LOCATION + "'," + QUANTITY + "," + WEIGHT + " ,'" + WEIGHTTYPE + "')";
                    commandinsert.ExecuteNonQuery();
                    connection.Close();
                    MessageBox.Show("Record Inserted successfully");

                }

                catch (Exception ex)
                {
                    MessageBox.Show("ENTER CORRECT DATA FORMAT"+ ex);
                    connection.Close();


                }

            try
            {
                connection.Close();
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }
                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString();


            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        


            
          

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.InitialDirectory = "C:";
                saveFileDialog1.Title = "Save as Excel File";
                saveFileDialog1.FileName = "";
                saveFileDialog1.Filter = "Excel Files (2013)| * .xlsx; * .xls; * .xlsm | Excel Files (2007)| * .xlsx;  * .xls; * .xlsm";
                if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
                {

                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    ExcelApp.Application.Workbooks.Add(Type.Missing);

                    //change property of work book
                    ExcelApp.Columns.ColumnWidth = 20;
                    // storing header part in Excel
                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        ExcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }
                    //storing Each row and column value to excel sheet
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        try
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {

                                if (dataGridView1.Rows[i].Cells[j].Value == null)
                                {
                                    dataGridView1.Rows[i].Cells[j].Value = "BORN ACORD FARM";
                                }
                                ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                            }



                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("EMPTY");
                            connection.Close();
                        }

                    }
                    ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                    ExcelApp.ActiveWorkbook.Saved = true;
                    ExcelApp.Quit();
                    MessageBox.Show("Excel file created");
                    // ExcelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DGVPrinterHelper.DGVPrinter printer = new DGVPrinterHelper.DGVPrinter();
                printer.Title = "WAREHOUSE STOCK\n\n";
                printer.SubTitle = string.Format("Date: {0}", DateTime.Now.Date + "\n\n\n");
                printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
                printer.PageNumbers = true;
                printer.PageNumberInHeader = false;
                printer.PorportionalColumns = true;
                printer.HeaderCellAlignment = StringAlignment.Near;
                printer.Footer = "BORN ACORD FARM";
                printer.FooterSpacing = 15;
                printer.FooterAlignment = StringAlignment.Center;
                printer.PrintDataGridView(dataGridView1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
             try
            {
                connection.Close();
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }
                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString(); 

                
            }
             catch (Exception ex)
             {
                 MessageBox.Show("EMPTY");
                 connection.Close();


             }
        }

        private void button7_Click(object sender, EventArgs e)
        {
             
        

            try
            {

                connection.Open();

                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb WHERE ITEM = '" + Convert.ToString(textBox4.Text) + "' ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }

                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString();

                

            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();
            }
        
        }
        public double id = 0;
        public double QT = 0;
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];


            comboBox3.Text = row.Cells[6].Value.ToString();
            id = double.Parse(row.Cells[0].Value.ToString());
            QT = double.Parse(row.Cells[4].Value.ToString());
            textBox1.Text = row.Cells[4].Value.ToString();
            textBox2.Text = row.Cells[3].Value.ToString();
            textBox3.Text = row.Cells[2].Value.ToString();
            textBox4.Text = row.Cells[2].Value.ToString();
            textBox5.Text = row.Cells[5].Value.ToString();
           
        }

        private void button4_Click(object sender, EventArgs e)
        {

            try
            {
                if (dataGridView1.SelectedRows.Count >= 1 || dataGridView1.SelectedCells.Count >= 1)
                {

                        string titlesave = "Message";
                    MessageBoxButtons buttonssave = MessageBoxButtons.YesNo;
                    DialogResult resultsave = MessageBox.Show("ARE YOU SURE YOU WANT TO UPDATE THIS RECORD !!!!", titlesave, buttonssave);
                    if (resultsave == DialogResult.No)
                    {
                    }
                    else
                    {

                        connection.Close();
                        connection.Open();
                        OleDbCommand cmd = connection.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = " update [warehousetb] set  ITEM =  '" + textBox3.Text + "' , LOCATION = '" + textBox2.Text + "' , QUANTITY = " + textBox1.Text + ", [UNIT WEIGHT] = " + textBox5.Text + " , [WEIGHT TYPE] = '" + comboBox3.Text + "' WHERE ID = " + id + "";
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Record Updated successfully");

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error");
                connection.Close();


            }
            try
            {
                connection.Close();
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }
                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString(); 

                
            }
             catch (Exception ex)
             {
                 MessageBox.Show("EMPTY");
                 connection.Close();


             }
        

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count >= 1 || dataGridView1.SelectedCells.Count >= 1)
                {

                    double fn = 0;

                    double pp = double.Parse(textBox1.Text.Trim());                  
                    fn = QT - pp;
                    fn = QT - pp;
                    if (fn > 0)
                    {
                         string titlesave = "Message";
                    MessageBoxButtons buttonssave = MessageBoxButtons.YesNo;
                    DialogResult resultsave = MessageBox.Show("ARE YOU SURE YOU WANT TO REMOVE SOME QUANTITY OF THIS ITEM THIS PERMANENTLY!!!!", titlesave, buttonssave);
                    if (resultsave == DialogResult.No)
                    {
                    }
                    else
                    {
                        connection.Close();
                        connection.Open();
                        OleDbCommand cmd = connection.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = " update [warehousetb] set  QUANTITY = " + fn + " WHERE ID = " + id + "";
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Record Removed successfully");
                    }
                    }
                    else if (fn <= 0)
                    {
                             string titlesave = "Message";
                    MessageBoxButtons buttonssave = MessageBoxButtons.YesNo;
                    DialogResult resultsave = MessageBox.Show("ARE YOU SURE YOU WANT TO REMOVE ALL RECORDS OF THIS ITEM PERMANENTLY!!!!", titlesave, buttonssave);
                    if (resultsave == DialogResult.No)
                    {
                    }
                    else
                    {
                        connection.Close();
                        connection.Open();
                        OleDbCommand cmd = connection.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = " delete * from warehousetb  WHERE ID = " + id + "";
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Records Removed successfully");
                    }
                    }



                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error");
                connection.Close();


            }
           
            try
            {
                connection.Close();
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }
                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString();


            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {

                 string titlesave = "Message";
                    MessageBoxButtons buttonssave = MessageBoxButtons.YesNo;
                    DialogResult resultsave = MessageBox.Show("ARE YOU SURE YOU WANT TO DELETE ALL RECORDS PERMANENTLY!!!!", titlesave, buttonssave);
                    if (resultsave == DialogResult.No)
                    {
                    }
                    else
                    {

                        connection.Close();
                        connection.Open();
                        OleDbCommand cmd = connection.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = " delete * from warehousetb";
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        MessageBox.Show("Record deleted successfully");
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        
      
            try
            {
                connection.Close();
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }
                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString();


            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            try
            {
                connection.Close();
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("select * from warehousetb ", connection);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // paiddatagridview1.Rows[0].DefaultCellStyle.Font = new Font("Tahoma",12,FontStyle.Bold,ForeColor.Black);
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    dataGridView1.Rows[y].DefaultCellStyle.ForeColor = Color.Black;

                }
                label12.Text = Convert.ToString(dataGridView1.Rows.Count - 1);
                connection.Close();
                double a;
                double b = 0;
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {

                    a = Convert.ToDouble(r.Cells[4].Value);
                    b = b + a;
                }
                label9.Text = b.ToString();


            }
            catch (Exception ex)
            {
                MessageBox.Show("EMPTY");
                connection.Close();


            }
        
        }
    }
}
