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
using MySql.Data.MySqlClient;
namespace ImportDataToMysql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void SelectFileButton_Click(object sender, EventArgs e)
        {
            if(cbImportType.SelectedIndex == -1)
            {
                MessageBox.Show("請先選擇欲匯入的資料種類");
                return;
            }
            else
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    
                    
                    FileNameTextBox.Text = openFileDialog1.FileName;
                   
                }
            }
            
        }
        private void ImportButton_Click(object sender, EventArgs e)
        {
            if(cbImportType.SelectedIndex == -1)
            {
                MessageBox.Show("請先選擇欲匯入的資料種類");
                return;
            }
            else
            {
                if (FileNameTextBox.Text == "")
                {
                    MessageBox.Show("請先選擇檔案");
                    return;
                }
            }
            if(cbImportType.SelectedIndex == 2)
            {
                //當選擇組距資料,將多選的檔案一一匯入
                foreach(var file in openFileDialog1.FileNames)
                {
                    ReadExcel(file);
                }
            }
            else
            {
                ReadExcel(openFileDialog1.FileName);
            }
            ImportButton.Enabled = true;
       
        }
        private void InsertDataToMySql(DataTable import_data, string DataName)
        {
            //MySQL連線資訊
            string dbHost = "localhost";
            string dbUser = "scpa";
            string dbPass = "0813";
            string dbName = "scpa";



            string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
            MySqlConnection mysqlConn = new MySqlConnection(connStr);
            mysqlConn.Open();
            string delcomStr = "";
            string comdStr = "";
            if (DataName.IndexOf("學年度四技甄選簡章資料") >= 0)
            {
                delcomStr = "Delete from thisyearbrief";
                string cols = "";
                //利用迴圈產生78個欄位參數
                for (int c = 1; c <= 78; c++)
                {
                    if (c == 78)
                    {
                        cols += "?Col" + c;
                    }
                    else {
                        cols += "?Col" + c + ",";
                    }

                }
                comdStr = "insert into thisyearbrief values(" + cols + ")";
            }
            else if (DataName.IndexOf("名額分數合併") >= 0)
            {
                delcomStr = "Delete from quoscore";
                comdStr = "insert into quoscore values(?Col1,?Col2,?Col3,?Col4,?Col5,?Col6,?Col7,?Col8,?Col9,?Col10,?Col11,?Col12,?Col13,?Col14,?Col15,?Col16)";
            }
            else if (DataName.IndexOf("是否限選填一校一系") >= 0)
            {
                delcomStr = "Delete from onlyone";
                comdStr = "insert into onlyone values(?Col1,?Col2,?Col3,?Col4,?Col5,?Col6)";
            }


            int i = 0;

            MySqlTransaction transaction = mysqlConn.BeginTransaction();
            MySqlCommand mysqlCom = new MySqlCommand();
            mysqlCom.Transaction = transaction;
            foreach (DataRow dr in import_data.Rows)
            {
                //匯入前，先將資料庫裡的資料清除
                if (i == 0)
                {
                    MySqlCommand delmysqlCom = new MySqlCommand(delcomStr, mysqlConn);
                    delmysqlCom.ExecuteNonQuery();
                    delmysqlCom.Dispose();
                }


                mysqlCom.Connection = mysqlConn;
                mysqlCom.CommandText = comdStr;
                mysqlCom.Parameters.Clear();
                if (DataName.IndexOf("學年度四技甄選簡章資料") >= 0)
                {
                    //匯入今年學年度四技甄選簡章資料
                    ThisyearbriefDataImport(mysqlCom, dr);
                }
                else if (DataName.IndexOf("名額分數合併") >= 0)
                {
                    //匯入名額分數資料
                    quoscoreDataImport(mysqlCom, dr);
                }
                else if (DataName.IndexOf("是否限選填一校一系") >= 0)
                {
                    //匯入是否限選填一校一系資料

                    if(dr[0].ToString() =="")
                    {
                        continue;
                    }
                    else
                    {
                        OnlyoneDataImport(mysqlCom, dr);
                    }
                        
                    

                }


               
                    mysqlCom.ExecuteNonQuery();
                    mysqlCom.Dispose();
                
                

                i++;
                //進度條
                // Importprogress.Value += 1;
            }
            transaction.Commit();

            mysqlConn.Close();
            i = 0;



        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void quoscoreDataImport(MySqlCommand mysqlCom, DataRow dr)
        {
            int checkisnum;
            for (int j = 0; j <= 15; j++)
            {
                if (j <= 5)
                {
                    mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), dr[j].ToString());
                }
                else {
                    if (int.TryParse(dr[j].ToString(), out checkisnum))
                    {
                        mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), Convert.ToInt32(dr[j]));
                    }
                    else {
                       if(dr[j].ToString().IndexOf("106")>=0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -106);
                       }
                        else if (dr[j].ToString().IndexOf("105") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -105);
                        }
                        else if (dr[j].ToString().IndexOf("104") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -104);
                        }
                        else if (dr[j].ToString().IndexOf("103") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -103);
                        }
                        else if (dr[j].ToString() == "")
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -1);
                        }
                        else if (dr[j].ToString().IndexOf("其他方式計分") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -2);
                        }
                        else if (dr[j].ToString().IndexOf("推甄不招生") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -3);
                        }
                        else if (dr[j].ToString().IndexOf("未公告級分") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -4);
                        }
                        else if (dr[j].ToString().IndexOf("無一般生名額") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -5);
                        }
                        else if (dr[j].ToString().IndexOf("--") >= 0)
                        {
                            mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -6);
                        }
                        //else
                        //{
                        //    mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), -7);
                        //}


                    }
                }

            }
        }

        private void OnlyoneDataImport(MySqlCommand mysqlCom, DataRow dr)
        {
            mysqlCom.Parameters.AddWithValue("?Col1", dr[0].ToString());
            mysqlCom.Parameters.AddWithValue("?Col2", dr[1].ToString());
            for (int k = 2; k <= 3; k++)
            {
                if (dr[k].ToString() == "是")
                {
                    mysqlCom.Parameters.AddWithValue("?Col" + (k + 1), true);
                }
                else {
                    mysqlCom.Parameters.AddWithValue("?Col" + (k + 1), false);
                }
            }
            if(dr[4].ToString() =="")
            {
                mysqlCom.Parameters.AddWithValue("?Col" + (5), -1);
            }
            else
            {
                mysqlCom.Parameters.AddWithValue("?Col" + (5), dr[4].ToString());
            }
            switch(dr[5].ToString())
            {
                case "國立科大":
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), 1);
                    break;
                case "私立科大":
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), 2);
                    break;
                case "私立學院":
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), 3);
                    break;
                case "國立專校":
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), 4);
                    break;
                case "國立大學":
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), 5);
                    break;
                case "私立大學":
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), 6);
                    break;
                default:
                    mysqlCom.Parameters.AddWithValue("?Col" + (6), -1);
                    break;

            }

        }

        private void ThisyearbriefDataImport(MySqlCommand mysqlCom, DataRow dr)
        {
            int checkisnum;

            for (int j = 0; j <= 77; j++)
            {

                if (int.TryParse(dr[j].ToString(), out checkisnum))
                {
                    mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), Convert.ToInt32(dr[j]));
                }
                else {


                    mysqlCom.Parameters.AddWithValue("?Col" + (j + 1), dr[j].ToString());


                }


            }
        }
        private void ImportAccountsData(DataTable data)
        {
            MysqlConn mysqlconn = new MysqlConn();
            mysqlconn.Open();
            //string dbHost = "localhost";
            //string dbUser = "scpa";
            //string dbPass = "0813";
            //string dbName = "scpa";


            int i;
            int index=0;
            string c="";
           // string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
           // MySqlConnection mysqlConn = new MySqlConnection(connStr);
          //  mysqlConn.Open();
            string insertcmd = "Insert into accounts values ({0})";
            string importcmd = "";
           
            for (i = 0; i < 19;i++)
            {
                if( i == 18)
                {
                    c += "?Col" + (i + 1) ;
                }
                else
                {
                    c += "?Col" + (i + 1) + ",";
                }
            }
            importcmd = string.Format(insertcmd, c);
            MySqlCommand cmd = new MySqlCommand(importcmd, mysqlconn.conn);
            foreach ( DataRow dr in data.Rows)
            {
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("?Col1", dr["UserID"].ToString());
                cmd.Parameters.AddWithValue("?Col2", dr["Password"].ToString());
                cmd.Parameters.AddWithValue("?Col8", "09 商業與管理群");
                for(index = 3;index <= 19; index++)
                {
                    if( index >= 3 && index <= 6)
                    {
                        cmd.Parameters.AddWithValue("?Col"+index, null);
                    }
                    else
                    {
                        if(index == 7 || (index >= 9 && index <= 18))
                        {
                            cmd.Parameters.AddWithValue("?Col" + index, 0);
                        }
                        else
                        {
                            if(index == 19)
                            {
                                cmd.Parameters.AddWithValue("?Col" + index, 1);
                            }
                        }
                    }
                }
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                
            }
            // mysqlConn.Close();
            mysqlconn.Close();
            
           

        }
        private void ImportGroupDisData(DataTable data,string filename,string year)
        {
            //string dbHost = "localhost";
            // string dbUser = "scpa";
            // string dbPass = "0813";
            // string dbName = "scpa";



            // string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
            // MySqlConnection mysqlConn = new MySqlConnection(connStr);
            // mysqlConn.Open();
            MysqlConn mysqlconn = new MysqlConn();
            mysqlconn.Open();
            string delSql = "Delete from grpdis where Year ="+year;
            string InsertSql = "insert into grpdis values (?Col1,?Col2,?Col3,?Col4,?Col5,?Col6)";
           
            MySqlCommand cmd = new MySqlCommand(InsertSql, mysqlconn.conn);
            MySqlCommand delcmd = new MySqlCommand(delSql, mysqlconn.conn);
           
           // delcmd.ExecuteNonQuery();
            foreach (DataRow dr in data.Rows)
            {
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("?Col1", year);
                cmd.Parameters.AddWithValue("?Col2", filename);
                cmd.Parameters.AddWithValue("?Col3", dr[0]);
                cmd.Parameters.AddWithValue("?Col4", dr[1]);
                if(dr[2].ToString()=="")
                {
                    cmd.Parameters.AddWithValue("?Col5", -1);
                }
                else
                {
                    cmd.Parameters.AddWithValue("?Col5", dr[2]);
                }
                
                cmd.Parameters.AddWithValue("?Col6", dr[3]);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            
            
            mysqlconn.Close();
        }
        private void ImportDayData(DataTable data)
        {
            MysqlConn mysqlconn = new MysqlConn();
            
            mysqlconn.Open();
            MySqlTransaction ts = mysqlconn.conn.BeginTransaction();
            string Insert_sql = "Insert into distribution values ({0})";
            //匯入前先將原有資料刪除
            string Del_sql = "Delete from distribution";
            MySqlCommand del_cmd = new MySqlCommand(Del_sql,mysqlconn.conn);
            del_cmd.ExecuteNonQuery();
            del_cmd.Dispose();
            int checkisnum;
            string c = "";
            int i ;
            //利用迴圈產生9個欄位
            for (i = 1; i <= 9; i++) 
            {
                if(i==9)
                {
                    c += "?Col" + i;
                }
                else
                {
                    c += "?Col" + i + ",";
                }
               
            }
            Insert_sql = string.Format(Insert_sql, c);
            MySqlCommand cmd = new MySqlCommand(Insert_sql, mysqlconn.conn);
            cmd.Transaction = ts;
            foreach(DataRow dr in data.Rows)
            {
                cmd.Parameters.Clear();
                //匯入前三個欄位(全中文)
                for(int j=1;j<=3;j++)
                {
                    cmd.Parameters.AddWithValue("?Col"+j,dr[j].ToString());
                }
                for(int j=4;j<=9;j++)
                {
                    //判斷內容是否為數值
                    if (int.TryParse(dr[j].ToString(), out  checkisnum))
                    {
                        cmd.Parameters.AddWithValue("?Col" + j, dr[j].ToString());
                    }
                    else
                    {
                        string year = (DateTime.Now.Year-1911)+"新增";
                        if(dr[j].ToString().IndexOf(year)>=0)
                        {
                            cmd.Parameters.AddWithValue("?Col" + j, -106);
                        }
                        else
                        {
                            switch (dr[j].ToString())
                            {
                                case "":
                                    cmd.Parameters.AddWithValue("?Col" + j, -1);
                                    break;
                                case "分發無招生":
                                    cmd.Parameters.AddWithValue("?Col" + j, -7);
                                    break;
                                case "其他方式計分":
                                    cmd.Parameters.AddWithValue("?Col" + j, -2);
                                    break;
                                case "--":
                                    cmd.Parameters.AddWithValue("?Col" + j, -6);
                                    break;
                                case "無一般生名額":
                                    cmd.Parameters.AddWithValue("?Col" + j, -5);
                                    break;
                                case "甄選無招生":
                                    cmd.Parameters.AddWithValue("?Col" + j, -3);
                                    break;
                                default:
                                    cmd.Parameters.AddWithValue("?Col" + j, dr[j]);
                                    break;

                            }
                        }
                       
                    }
                }
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            ts.Commit();
            mysqlconn.Close();
        }
        private void ImportAdmissionData(DataTable data)
        {
            MysqlConn mysqlconn = new MysqlConn();
            mysqlconn.Open();
            string insert_sql = "insert into estimates values ({0})";
            string del_sql = "Delete from estimates";
            //匯入前先把原有資料刪除
            MySqlCommand del_cmd = new MySqlCommand(del_sql,mysqlconn.conn);
            del_cmd.ExecuteNonQuery();
            del_cmd.Dispose();
            int checkisnum;
            int i;
            string c = "";
            for(i=1;i<=10;i++)
            {
                if( i == 10)
                {
                    c += "?Col"+i;
                }
                else
                {
                    c += "?Col" + i + ",";
                }
            }
            insert_sql = string.Format(insert_sql, c);
            MySqlCommand cmd = new MySqlCommand(insert_sql, mysqlconn.conn);
            MySqlTransaction ts = mysqlconn.conn.BeginTransaction();
            cmd.Transaction = ts;
            int index = 0;
            foreach(DataRow dr in data.Rows)
            {
                cmd.Parameters.Clear();
                if(index > 0)
                {
                    for (int j = 1; j <= 5; j++) 
                    {
                        cmd.Parameters.AddWithValue("?Col"+j,dr[j-1].ToString());
                    }
                    for (int j = 6; j <= 10; j++) 
                    {
                        if (int.TryParse(dr[j-1].ToString(), out checkisnum))
                        {
                            cmd.Parameters.AddWithValue("?Col" + j, dr[j-1].ToString());
                        }
                        else
                        {
                            string year = (DateTime.Now.Year - 1911) + "新增";
                            if (dr[j-1].ToString().IndexOf(year) >= 0)
                            {
                                cmd.Parameters.AddWithValue("?Col" + j, -106);
                            }
                            else
                            {
                                switch (dr[j-1].ToString())
                                {
                                    case "":
                                        cmd.Parameters.AddWithValue("?Col" + j, -1);
                                        break;
                                    case "分發無招生":
                                        cmd.Parameters.AddWithValue("?Col" + j, -7);
                                        break;
                                    case "其他方式計分":
                                        cmd.Parameters.AddWithValue("?Col" + j, -2);
                                        break;
                                    case "--":
                                        cmd.Parameters.AddWithValue("?Col" + j, -6);
                                        break;
                                    case "無一般生名額":
                                        cmd.Parameters.AddWithValue("?Col" + j, -5);
                                        break;
                                    case "甄選無招生":
                                        cmd.Parameters.AddWithValue("?Col" + j, -3);
                                        break;
                                    case "#N/A":
                                        cmd.Parameters.AddWithValue("?Col" + j, -8);
                                        break;
                                    default:
                                        cmd.Parameters.AddWithValue("?Col" + j, dr[j-1]);
                                        break;

                                }
                            }
                        }
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                
                index++;
            }
            ts.Commit();
            mysqlconn.Close();

        }
        private void ImportCollateData(DataTable data)
        {
            MysqlConn mysqlconn = new MysqlConn();
            mysqlconn.Open();
            string del_sql = "Delete from scorecollate";
            string insert_sql = "Insert into scorecollate values({0})";
            //匯入前先把原有資料刪除
            MySqlCommand del_cmd = new MySqlCommand(del_sql,mysqlconn.conn);
            del_cmd.ExecuteNonQuery();
            del_cmd.Dispose();
            int i;
            string c = "";
            int checkisnum;
            for (i=1;i<=26;i++)
            {
                if (i == 26)
                {
                    c += "?Col" + i;
                }
                else
                {
                    c += "?Col" + i + ",";
                }
            }
            insert_sql = string.Format(insert_sql, c);
            MySqlCommand cmd = new MySqlCommand(insert_sql, mysqlconn.conn);
            MySqlTransaction st = mysqlconn.conn.BeginTransaction();
            cmd.Transaction = st;
            int index = 0;
            foreach(DataRow dr in data.Rows)
            {
                cmd.Parameters.Clear();
                if(index > 0)
                {
                    for (int j = 0; j < 8; j++) 
                    {
                        cmd.Parameters.AddWithValue("?Col" + (j + 1), dr[j].ToString());
                    }
                    for (int j = 8; j < 26; j++) 
                    {
                        string year = (DateTime.Now.Year - 1911)+"";
                        if (dr[j].ToString().IndexOf((year + "新增")) >= 0)
                        {
                            cmd.Parameters.AddWithValue("?Col" + (j + 1), "-"+year);
                        }
                        else
                        {
                            year = ((DateTime.Now.Year - 1911)-1) + "";
                            if (dr[j].ToString().IndexOf((year + "新增")) >= 0)
                            {
                                cmd.Parameters.AddWithValue("?Col" + (j + 1), "-"+year);
                            }
                            else
                            {
                                year = ((DateTime.Now.Year - 1911) - 2) + "";
                                if (dr[j].ToString().IndexOf((year + "新增")) >= 0)
                                {
                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), "-" + year);
                                }
                                else
                                {
                                    year = ((DateTime.Now.Year - 1911) - 3) + "";
                                    if (dr[j].ToString().IndexOf((year + "新增")) >= 0)
                                    {
                                        cmd.Parameters.AddWithValue("?Col" + (j + 1), "-" + year);
                                    }
                                    else
                                    {
                                        year = ((DateTime.Now.Year - 1911) - 4) + "";
                                        if (dr[j].ToString().IndexOf((year + "新增")) >= 0)
                                        {
                                            cmd.Parameters.AddWithValue("?Col" + (j + 1), "-" + year);
                                        }
                                        else
                                        {
                                            switch (dr[j].ToString())
                                            {
                                                case "":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -1);
                                                    break;
                                                case "分發無招生":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -7);
                                                    break;
                                                case "其他方式計分":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -2);
                                                    break;
                                                case "--":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -6);
                                                    break;
                                                case "無一般生名額":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -5);
                                                    break;
                                                case "甄選無招生":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -3);
                                                    break;
                                                case "#N/A":
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), -8);
                                                    break;
                                                default:
                                                    cmd.Parameters.AddWithValue("?Col" + (j + 1), dr[j]);
                                                    break;

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                index++;
            }
            st.Commit();
            mysqlconn.Close();

        }
        private void ReadExcel(string filepath)
        {
            MsgtextBox.Text = "正在讀取檔案...........\r\n";
            ImportButton.Enabled = false;
           
            //利用OLEDB 讀取Excel檔
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + filepath + ";Extended Properties='Excel 12.0; HDR=YES; IMEX=1'";
            OleDbConnection OleConn = new OleDbConnection(strConn);
            OleConn.Open();
            DataTable SheetName_data = new DataTable();
            DataTable data = new DataTable();
            OleDbDataAdapter OleAdapter;
            //取得Excel工作表名稱
            SheetName_data = OleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetname;
            string strComm;
            //讀取Excel每個分頁的資料
            for (int i = 0; i < SheetName_data.Rows.Count; i++)
            {
                data = new DataTable();
                sheetname = SheetName_data.Rows[i]["TABLE_NAME"].ToString();
                strComm = "Select * from [{0}]";
                if (!(sheetname.IndexOf("FilterDatabase") >= 0))
                {
                    OleAdapter = new OleDbDataAdapter(string.Format(strComm, sheetname), strConn);
                    OleAdapter.Fill(data);
                    sheetname = sheetname.Replace("$", "");
                    MsgtextBox.Text += "正在匯入" + sheetname + "..................";
                    switch(cbImportType.SelectedIndex)
                    {
                        case 0: //落點分析資料
                            InsertDataToMySql(data, sheetname);
                            break;
                        case 1://帳戶資料
                            ImportAccountsData(data);
                            break;
                        case 2: //組距資料
                            int indexS;
                            int indexS2;
                            string name;
                            string year;
                            indexS = System.IO.Path.GetFileName(filepath).IndexOf("-") + 1;
                            indexS2 = System.IO.Path.GetFileName(filepath).IndexOf("(") + 1;
                            name = System.IO.Path.GetFileName(filepath).Substring(indexS, (System.IO.Path.GetFileName(filepath).Length - indexS));
                            year = System.IO.Path.GetFileName(filepath).Substring(indexS2, 3);
                            name = name.Replace(".xlsx", "");
                            ImportGroupDisData(data, name,year);
                            break;
                        case 3://甄選+分發資料
                            ImportDayData(data);
                            break;
                        case 4://分發預估錄取分數
                            ImportAdmissionData(data);
                            break;
                        case 5://歷年甄選分數對照
                            ImportCollateData(data);
                            break;

                    }
                                   
                    MsgtextBox.Text += "OK\r\n";
                }
                //組距資料只讀取Excel的第一分頁
                if(cbImportType.SelectedIndex == 2 || cbImportType.SelectedIndex == 3 || cbImportType.SelectedIndex == 4 || cbImportType.SelectedIndex == 5)
                {
                    break;
                }

            }



            MsgtextBox.Text += "匯入完成";
            //MessageBox.Show("匯入完成");
            if(cbImportType.SelectedIndex != 2)
            {
                ImportButton.Enabled = true;
            }
            
            OleConn.Close();
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cbImportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            //當選擇組距資料時,可多選資料檔
            if(cbImportType.SelectedIndex == 2)
            {
                openFileDialog1.Multiselect = true;
            }
            else
            {
                openFileDialog1.Multiselect = false;
            }
        }
    }
}