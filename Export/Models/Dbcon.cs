using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Export.Models
{
    public class Dbcon
    {
        public class Database
        {
            static string server = "108.60.209.65";
            static string user = "aaqa_shared_dba";
            static string pass = "Sh0hel123";
            static string database = "RCLWEB_DB";

            //static string server = "159.21.2.16";
            //static string user = "sa";
            //static string pass = "R0yal123";
            //static string database = "RCLWEB_DB_NEW";

            //static string server = "LAPTOP-G3Q6J4HR";
            //static string user = "sa";
            //static string pass = "A@qaTech123";
            //static string database = "RCLWEB_DB";

            public static string DbCon = "User ID=" + user + "; Password=" + pass + "; Initial Catalog=" + database + "; Data Source=" + server + "";

            public static string DbPrefix = "wc_";
        }
    }
}