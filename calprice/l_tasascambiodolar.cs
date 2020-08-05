using System;
using System.Windows.Forms;
using System.Data;
using System.Data.SQLite;

namespace calprice
{
    class l_tasascambiodolar
    {

        public  int         id;
        public  DateTime    fecha;
        public  decimal     tasacambio;

        public l_tasascambiodolar()
        {
            id         = 0;
            fecha      = DateTime.Now;
            tasacambio = 0;
        }

        public void insert()
        {
            string sql = "insert into tasascambiodolar (fecha,tasacambio) " +
                         "values (@p1,@p2)";

            SQLiteCommand cmdSql = new SQLiteCommand(sql, misclases.conexLite);

            cmdSql.Parameters.Add(new SQLiteParameter("@p1", fecha));
            cmdSql.Parameters.Add(new SQLiteParameter("@p2", tasacambio));

            try
            {
                cmdSql.ExecuteNonQuery();
            }
            catch (SQLiteException err)
            {
                MessageBox.Show(err.Message.ToString());
            }
        }

        public DataTable select(string condi)
        {
            DataTable dt1 = new DataTable();
            string    sql = "select *,(strftime('%d/%m/%Y %H:%M',fecha)||'  tasa  '||tasacambio) fechaytasa " +
                            "from tasascambiodolar " +
                            "where "+condi;

            try
            {
                SQLiteDataAdapter da = new SQLiteDataAdapter(sql, misclases.conexLite);
                da.Fill(dt1);
                da.Dispose();
            }
            catch (SQLiteException err)
            {
                MessageBox.Show(err.Message.ToString());
            }

            return dt1;
        }

    }
}
