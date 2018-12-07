using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HBExportExcel.Model
{
    class MySQL
    {
        private static string Host = "46.45.139.109";
        private static string User = "akilliph_wecart";
        private static string DBname = "akilliph_wecart";
        private static string Password = "@i_rURz?+pbP";
        private static string Port = "5432";
        public MySqlDataAdapter baglayici = new MySqlDataAdapter();
        public MySqlCommand komut = new MySqlCommand();
        public DataTable table = new DataTable();
        public MySqlConnection baglan;
        public MySqlDataReader dataReader;
        public int count = 0;

        public MySQL()
        {
            this.SqlConnection();
        }

        private void SqlConnection()
        {
            try
            {
                this.baglan = new MySqlConnection("Server=" + Host + "; Database=" + DBname + "; uid=" + User + "; Password=" + Password + ";Persist Security Info=True;Convert Zero Datetime=True");
                this.baglan.Open();
            }
            catch (Exception msg)
            {
                // something went wrong, and you wanna know why
                MessageBox.Show(msg.ToString());
                //throw;
            }
        }

        public void CreateMySqlCommand(string myExecuteQuery, MySqlConnection myConnection)
        {
            MySqlCommand myCommand = new MySqlCommand(myExecuteQuery, myConnection);
            myCommand.Connection.Open();
            myCommand.ExecuteNonQuery();
            myConnection.Close();
        }

        public void FillDatatable(String query, DataGridView view)
        {
            //this.table.Reset();
            try
            {

                //this.komut.CommandText = query;
                //this.komut.Connection = this.baglan;
                //this.baglayici.SelectCommand = this.komut;
                //this.baglan.Open();
                //this.baglayici.Fill(this.table);

                //view.DataSource = this.table;
                //view.Refresh();
                //this.baglan.Close();

                if (this.baglan.State.ToString() == "Closed")
                {
                    this.baglan.Open();
                }

                MySqlDataAdapter MyDA = new MySqlDataAdapter();
                string sqlSelectAll = query;
                MyDA.SelectCommand = new MySqlCommand(sqlSelectAll, this.baglan);
                DataTable table = new DataTable();
                MyDA.Fill(table);
                BindingSource bSource = new BindingSource();
                bSource.DataSource = table;
                view.DataSource = bSource;
                this.baglan.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
        }

        public void UpdateData(String query)
        {
            Console.WriteLine("update data****************************************************************");
            this.count = 0;
            int numRowsUpdated = 0;
            try
            {
                this.baglan.Open();
            }
            catch (Exception ex)
            {
                this.baglan.Close();
                this.baglan.Dispose();
                Console.WriteLine(ex.Message + "update hatası");
                this.UpdateData(query);
            }
            try
            {
                this.komut.Connection = this.baglan;
                this.komut.CommandText = query;
                numRowsUpdated = this.komut.ExecuteNonQuery();
                this.komut.Dispose();
                if (numRowsUpdated > 0)
                {
                    Console.WriteLine(numRowsUpdated.ToString() + " Satır Güncellendi");
                    this.count = numRowsUpdated;
                }
                this.baglan.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            //MessageBox.Show(numRowsUpdated.ToString() + " Satır etkilendi.");

        }

        public void updateNew(String query)
        {
            this.baglan.Close();
            try
            {
                this.baglan.Open();
                MySqlCommand cmd = new MySqlCommand(query, this.baglan);
                cmd.ExecuteNonQuery();
                Console.WriteLine("successfully updated");
                this.baglan.Close();
            }
            catch (Exception ex)
            {
                this.baglan.Close();
                Console.WriteLine("1" + ex.Message);
                this.updateNew(query);
            }
        }

        public void InsertData(String query)
        {
            //this.komut.Connection = this.baglan;
            //this.komut.CommandText = query;
            //this.baglan.CreateCommand();

            //this.komut.ExecuteNonQuery();

            try
            {
                this.baglan.Open();
                this.komut.Connection = this.baglan;
                this.komut.CommandText = query;
                this.komut.Prepare();

                this.komut.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (this.baglan != null)
                    this.baglan.Close();
            }
        }

        public string SelectQuery(string query)
        {

            try
            {
                //Display query  
                MySqlCommand MyCommand2 = new MySqlCommand(query, this.baglan);
                //  MyConn2.Open();  
                //For offline connection we weill use  MySqlDataAdapter class.  
                MySqlDataAdapter MyAdapter = new MySqlDataAdapter();
                MyAdapter.SelectCommand = MyCommand2;
                DataTable dTable = new DataTable();
                MyAdapter.Fill(dTable);
                return dTable.Rows[0][0].ToString(); // here i have assign dTable object to the dataGridView1 object to display data.               
                                                   // MyConn2.Close();  
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public void FetchDataToList(String query)
        {
            this.komut.CommandText = query;
            try
            {
                this.baglan.Open();
                this.dataReader = komut.ExecuteReader();
                while (dataReader.Read())
                {
                    MessageBox.Show(dataReader.GetString(0));
                }
                this.baglan.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.baglan.Close();
            }
        }
    }
}
