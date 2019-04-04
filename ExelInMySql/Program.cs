using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;


namespace ExelInMySql
{
    class Program
    {
        static void Main(string[] args)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;user=root;database=export;password=;");
            con.Open();
            string sql = "select * from user";
            MySqlCommand com = new MySqlCommand(sql, con);
            MySqlDataReader tabl = com.ExecuteReader();
            /*
            while (tabl.Read())
            {
                Console.WriteLine(String.Format("|{0:d3} | {1} | {2} | {3:d3}", tabl[0], tabl[1], tabl[2], tabl[3]));
            }
            */
            /*
            //con.ConnectionString = @"";
            Excel.Application excel = new Excel.Application();
            //excel.Visible = true;
            excel.Workbooks.Open("d:\\test.xlsx");
            int row = 2,kol=1;

            Excel.Worksheet currentSheet = (Excel.Worksheet)excel.Workbooks[1].Worksheets[1];

            while (currentSheet.Range["A" + row].Value2 != null)
            {
                List<string> tmp = new List<string>();
                for (int j = 1; j < 5; j++)
                {
                    tmp.Add( currentSheet.Cells[row,j].Text);
                }
                string sql = "insert into user (id,name,fam,kol) values (" + tmp[0] + ",'" + tmp[1] + "','" + tmp[2] + "'," + tmp[3] + ");";
                MySqlCommand add=new MySqlCommand  (sql, con);
                add.ExecuteNonQuery();
                Console.WriteLine(sql);
                row++;

            }
            excel.Quit();//*/
            con.Close();
            Console.ReadLine();
        }
    }
}
