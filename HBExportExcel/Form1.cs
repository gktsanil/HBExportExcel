using HBExportExcel.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HBExportExcel
{
    public partial class Form1 : Form
    {
        Export exp = new Export();
        static MySQL mysql = new MySQL();
        String filePathExp = "";
        List<ExcellList> excellLists = new List<ExcellList>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePathExp);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
            int columnCount = xlWorksheet.UsedRange.Columns.Count;
            List<string> columnNames = new List<string>();
            for (int c = 1; c < columnCount; c++)
            {
                if (xlWorksheet.Cells[1, c].Value2 != null)
                {
                    string columnName = xlWorksheet.Columns[c].Address;
                    Regex reg = new Regex(@"(\$)(\w*):");
                    if (reg.IsMatch(columnName))
                    {
                        Match match = reg.Match(columnName);
                        columnNames.Add(match.Groups[2].Value);
                    }
                    Console.WriteLine(columnName.ToString());

                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //GetWhere();
            //mysql.FillDatatable("SELECT p.productName, p.brands, p.productCode, pa.barcode, p.productDescription, pp.picture, pa.amountImage, p.id " +
            //   "FROM product AS p " +
            //   "INNER JOIN product_amount AS pa " +
            //   "ON p.id = pa.idProduct " +
            //   "INNER JOIN product_pictures AS pp " +
            //   "ON pa.idProduct = pp.idProduct " +
            //   "WHERE pa.amountImage != '' LIMIT 5", dataGridView1);
            //mysql.FillDatatable("SELECT * FROM hb_listings", dataGridView1);
            mysql.FillDatatable("SELECT * FROM hb_listings WHERE categoryId=132 OR categoryId=207",dataGridView1);
            MessageBox.Show(dataGridView1.Rows.Count.ToString());
        }

        private string GetWhere()
        {
            string where = "";
            foreach (var item in listBox3.Items)
            {
                where += item.ToString();
            }
            return "";
        }

        private string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();
            //XLSX Excel File
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = filePathExp;

            //XLS Excel File
            //props["Provider"] = "Microsoft.JET.OLEDB.4.0;";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = filePathExp;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            return sb.ToString();
        }

        private void WriteExcelFile()
        {
            string connectionString = GetConnectionString();

            using (var con = new OleDbConnection(connectionString))
            {
                con.Open();
                using (var cmd = new OleDbCommand("select * from [" + listBox1.SelectedItem.ToString() + "]", con))
                using (var reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly))
                {
                    var table = reader.GetSchemaTable();
                    var nameCol = table.Columns["ColumnName"];
                    foreach (DataRow row in table.Rows)
                    {
                        Console.WriteLine(row[nameCol]);
                    }
                }
            }

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;

                    //cmd.CommandText = "CREATE TABLE [table1] (id INT, name VARCHAR, datecol DATE );";
                    //cmd.ExecuteNonQuery();
                    try
                    {
                        //foreach (DataGridViewRow row in dataGridView1.Rows)
                        //{
                        //    string query = "INSERT INTO [" + listBox1.SelectedItem.ToString() + "](Temel, NoName, NoName1, NoName3, NoName4, NoName5, NoName6, NoName7, NoName8, NoName9) VALUES('" + row.Cells[0].Value.ToString() + "', '" + row.Cells[1].Value.ToString() + "-" + row.Cells[2].Value.ToString() + "', '" + row.Cells[2].Value.ToString() + "', '" + row.Cells[2].Value.ToString() + "', '" + brandTxt.Text.ToString() + "', '" + desiTxt.Text.ToString() + "', '" + taxNum.Value.ToString() + "', '" + warantyNum.Value.ToString() + "', '" + row.Cells[5].Value.ToString() + "', '" + row.Cells[4].Value.ToString() + "');";
                        //    Console.WriteLine(query);
                        //    cmd.CommandText = query;
                        //    cmd.ExecuteNonQuery();
                        //}
                        foreach (ExcellList item in excellLists)
                        {
                            string query = "INSERT INTO [" + listBox1.SelectedItem.ToString() + "](" +
                                "Temel, " +
                                "NoName, " +
                                "NoName1, " +
                                "NoName3, " +
                                "NoName4, " +
                                "NoName5, " +
                                "NoName6, " +
                                "NoName7, " +
                                "NoName8, " +
                                "NoName9, " +
                                "NoName10, " +
                                "NoName11, " +
                                "NoName12, " +
                                "Varyant, " +
                                "NoName13" +
                                ") VALUES(" +
                                "'" + item.ProductName.ToString() + "', " +
                                "'" + item.MerchantSku.ToString() + "', " +
                                "'" + item.Barcode.ToString() + "', " +
                                "'" + StripHTML(item.ProductDescriptiom.ToString()) + "', " +
                                "'" + item.BrandName.ToString() + "', " +
                                "'" + item.Desi.ToString() + "', " +
                                "'" + item.Vax.ToString() + "', " +
                                "'" + item.Waranty.ToString() + "', " +
                                "'" + item.Img1.ToString() + "', " +
                                "'" + item.Img2.ToString() + "', " +
                                "'" + item.Img3.ToString() + "', " +
                                "'" + item.Img4.ToString() + "', " +
                                "'" + item.Img5.ToString() + "', " +
                                "'', " +
                                "'" + item.CompModel.ToString() + "'" +
                                " );";
                            Console.WriteLine(query);
                            cmd.CommandText = query;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        listBox2.Items.Add(ex.Message);
                    }


                    conn.Close();
                }
                catch (Exception ex)
                {
                    listBox2.Items.Add(ex.Message);
                }
            }
        }

        public string StripHTML(string input)
        {
            //input = Regex.Replace(input, "<.*?>", String.Empty);
            //input = Regex.Replace(input, "'", String.Empty);
            //string veri = "";

            //for (int i = 0; i < 10; i++)
            //{
            //    veri += Regex.Replace(input, "<.*?>", String.Empty).Split(' ')[i];
            //}
            //return Regex.Replace(input, "<.*?>", String.Empty).Substring(0,40);
            //return input.Replace("\"", "&#34;");

            if (input.Length<255)
            {
                return input;
            }
            else
            {
                return input.Substring(0, 255);
            }
        }

        class ExcellList
        {
            public string ProductName { get; set; }
            public string MerchantSku { get; set; }
            public string Barcode { get; set; }
            public int VariantGrId { get; set; }
            public string ProductDescriptiom { get; set; }
            public string BrandName { get; set; }
            public int Desi { get; set; }
            public int Vax { get; set; }
            public int Waranty { get; set; }
            public string Img1 { get; set; }
            public string Img2 { get; set; }
            public string Img3 { get; set; }
            public string Img4 { get; set; }
            public string Img5 { get; set; }
            public string Color { get; set; }
            public string CompModel { get; set; }
            public int MyProperty { get; set; }
            //Product Specs

        }

        private void button2_Click(object sender, EventArgs e)
        {
            WriteExcelFile();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    mysql.FillDatatable("SELECT * FROM category LIMIT 100", metroGrid6);
            //    //mysql.FetchDataToList("SELECT * FROM category LIMIT 10");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    throw;
            //}

            //exp.ImportExcell(metroGrid1);
            filePathExp = exp.GetFilePath();
            listBox1.Items.Clear();
            foreach (var item in exp.GetExcelSheetNames(filePathExp))
            {
                listBox1.Items.Add(item);
            }
        }

        private void brandTxt_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            exp.ImportExcellFile(dataGridView1, listBox1.SelectedItem.ToString(), filePathExp);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //MessageBox.Show(row.Cells[1].Value.ToString().Split(',').Last());
                //foreach (DataGridViewCell cell in row.Cells)
                //{
                //    MessageBox.Show(cell.Value.ToString());
                //}
                ExcellList exList = new ExcellList();
                //MessageBox.Show(row.Cells[0].Value.ToString());
                exList.ProductName = ProductNameFix(row.Cells[1].Value.ToString());
                exList.MerchantSku = row.Cells[2].Value.ToString(); //+ "-" + row.Cells[3].Value.ToString();
                exList.CompModel = row.Cells[18].Value.ToString();//getBrandModel(row.Cells[1].Value.ToString().Split(',')[0]);
                exList.Barcode = row.Cells[3].Value.ToString();
                exList.ProductDescriptiom = row.Cells[7].Value.ToString();
                exList.BrandName = row.Cells[8].Value.ToString();//BrandNameFix(row.Cells[1].Value.ToString().Split(',').Last());//brandTxt.Text.ToString();
                exList.Desi = Convert.ToInt32(row.Cells[9].Value);//Convert.ToInt32(desiTxt.Text);
                exList.Vax = Convert.ToInt32(row.Cells[10].Value);//taxNum.Value);
                exList.Waranty = Convert.ToInt32(row.Cells[11].Value);//warantyNum.Value);
                exList.Img1 = "https://cdn.akilliphone.com/8004/1500x1500" + row.Cells[12].Value.ToString().Replace("img/", "/");
                exList.Img2 = "https://cdn.akilliphone.com/8004/1500x1500" + row.Cells[13].Value.ToString().Replace("img/", "/");
                exList.Img3 = "https://cdn.akilliphone.com/8004/1500x1500" + row.Cells[14].Value.ToString().Replace("img/", "/");
                exList.Img4 = "https://cdn.akilliphone.com/8004/1500x1500" + row.Cells[15].Value.ToString().Replace("img/", "/");
                exList.Img5 = "https://cdn.akilliphone.com/8004/1500x1500" + row.Cells[16].Value.ToString().Replace("img/", "/");
                exList.Color = row.Cells[17].Value.ToString();

                excellLists.Add(exList);
                Console.WriteLine(excellLists.Count);
            }

            foreach (ExcellList item in excellLists)
            {
                Console.WriteLine(item.ProductName.ToString());
                Console.WriteLine(item.BrandName.ToString());
                Console.WriteLine(item.Color.ToString());
                Console.WriteLine(item.Desi.ToString());
                Console.WriteLine(item.Waranty.ToString());
                Console.WriteLine(item.Vax.ToString());
                Console.WriteLine(item.Img1.ToString());
                Console.WriteLine(item.Img2.ToString());
                Console.WriteLine(item.Img3.ToString());
                Console.WriteLine(item.MerchantSku.ToString());
            }

            Console.WriteLine("bitti");
        }

        private string BrandNameFix(String brand)
        {
            if (brand == "0") {
                return brandTxt.Text.ToString();
            }
            else
            {
                return getBrandModel(brand);
            }
        }

        private string ProductNameFix(string productName)
        {
            if (productName.Split(' ')[0].ToLower() == brandTxt.Text.ToLower())
            {
                return productName;
            }
            else
            {
                return brandTxt.Text.ToString() + " " + productName;
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            mysql.FillDatatable("SELECT hb.*,pa.amountImage FROM hb_products AS hb INNER JOIN product_amount AS pa ON pa.barcode = hb.akilli_barcode ",dataGridView1);
            progressBar1.Maximum = dataGridView1.Rows.Count;

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                progressBar1.Value++;
                //MessageBox.Show(item.Cells[0].Value.ToString());
                mysql.InsertData("INSERT INTO hb_listings(productName,merchantSku,barcode,productId,desi,kdv,waranty) VALUES('" + item.Cells[3].Value.ToString() + "','" + item.Cells[5].Value.ToString() + "','" + item.Cells[2].Value.ToString() + "','" + item.Cells[1].Value.ToString() + "',1,18,24)");
                Thread.Sleep(5);
            }

        }

        private string getBrandModel(string id)
        {
            if (id=="0" || id=="")
            {
                return "Ally";
            }
            else
            {
                return mysql.SelectQuery("SELECT brandName FROM brands WHERE id=" + id);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            listBox3.Items.Add(textBox1.Text.ToString());
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBox3.Items.RemoveAt(listBox3.SelectedIndex);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            //mysql.FillDatatable("SELECT p.id,p.brands,pa.amountImage FROM product AS p INNER JOIN product_amount AS pa ON pa.idProduct = p.id WHERE pa.amountImage!=''", dataGridView1);
            //mysql.FillDatatable("SELECT hb.productName,hb.productId AS hbID,pc.idProduct AS pcID,pc.idCategory FROM product_categories AS pc INNER JOIN hb_listings AS hb ON hb.productId = pc.idProduct WHERE hb.img1!='' AND (pc.idCategory=132 OR pc.idCategory=207)", dataGridView1);
            //mysql.FillDatatable("SELECT * FROM hb_listings WHERE categoryId=207 OR categoryId = 132",dataGridView1);
            mysql.FillDatatable("SELECT hb.productName,hb.productId, pp.idProduct,pp.picture FROM hb_listings AS hb INNER JOIN product_pictures AS pp ON hb.productId = pp.idProduct WHERE (hb.categoryId=132 OR hb.categoryId=207)", dataGridView1);

            Thread.Sleep(2000);
            progressBar1.Maximum = dataGridView1.Rows.Count;
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                progressBar1.Value++;
                //Console.WriteLine(getBrandModel(item.Cells[1].Value.ToString().Split(',')[0]));
                //Console.WriteLine(item.Cells[2].Value.ToString());
                //Console.WriteLine(item.Cells[1].Value.ToString().Split(',')[0]);
                string queries = "UPDATE hb_listings SET img3 = '" + item.Cells[3].Value.ToString() + "' WHERE productId = " + item.Cells[1].Value.ToString();
                //Console.WriteLine(queries);
                mysql.UpdateData(queries);
                Thread.Sleep(5);
            }
            MessageBox.Show(dataGridView1.Rows.Count.ToString());
        }
    }
}
