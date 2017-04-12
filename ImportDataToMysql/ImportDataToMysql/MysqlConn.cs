using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
namespace ImportDataToMysql
{
    public  class MysqlConn
    {
        
        public MySqlConnection conn = new MySqlConnection("server=localhost;uid=root;pwd=0813;database=scpa");
        public  void Open()
        {           
            conn.Open();
        }
        public void Close()
        {
            conn.Close();
        }
    }
}
