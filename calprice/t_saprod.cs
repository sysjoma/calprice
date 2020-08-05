using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace calprice
{
    class t_saprod
    {

        public  decimal     CostAct;
        public  decimal     CostPro;
        public  decimal     CostAnt;
        public  decimal     Precio1;
        public  decimal     Precio2;
        public  decimal     Precio3;
        public  bool        sj_selec;
        public  decimal     sj_tasacambio;
        public  decimal     sj_costodolar;
        public  decimal     sj_p1dolar;
        public  decimal     sj_p2dolar;
        public  decimal     sj_p3dolar;
        public  decimal     sj_putilidad1;
        public  decimal     sj_putilidad2;
        public  decimal     sj_putilidad3;
        public  DateTime    sj_feulactualiza;

        public void update(string condi, string updateset = "")
        {
            string sql = "";

            SqlCommand cmdSql = new SqlCommand();

            if (updateset == "")
            {
                sql = "update saprod set " +
                      "CostAct          = @p1,  CostPro          = @p2, " +
                      "CostAnt          = @p3, " +
                      "Precio1          = @p4,  Precio2          = @p5, " +
                      "Precio3          = @p6,  sj_selec         = @p7, " +
                      "sj_tasacambio    = @p8,  sj_costodolar    = @p9, " +
                      "sj_p1dolar       = @p10,  sj_p2dolar      = @p11," +
                      "sj_p3dolar       = @p12, sj_putilidad1    = @p13," +
                      "sj_putilidad2    = @p14, sj_putilidad3    = @p15," +
                      "sj_feulactualiza = @p16 " +
                      "where " + condi;

                cmdSql.CommandText = sql;
                cmdSql.Connection = misclases.conexSql;

                cmdSql.Parameters.Add(new SqlParameter("@p1",  CostAct));
                cmdSql.Parameters.Add(new SqlParameter("@p2",  CostPro));
                cmdSql.Parameters.Add(new SqlParameter("@p3",  CostAnt));
                cmdSql.Parameters.Add(new SqlParameter("@p4",  Precio1));
                cmdSql.Parameters.Add(new SqlParameter("@p5",  Precio2));
                cmdSql.Parameters.Add(new SqlParameter("@p6",  Precio3));
                cmdSql.Parameters.Add(new SqlParameter("@p7",  sj_selec));
                cmdSql.Parameters.Add(new SqlParameter("@p8",  sj_tasacambio));
                cmdSql.Parameters.Add(new SqlParameter("@p9",  sj_costodolar));
                cmdSql.Parameters.Add(new SqlParameter("@p10", sj_p1dolar));
                cmdSql.Parameters.Add(new SqlParameter("@p11", sj_p2dolar));
                cmdSql.Parameters.Add(new SqlParameter("@p12", sj_p3dolar));
                cmdSql.Parameters.Add(new SqlParameter("@p13", sj_putilidad1));
                cmdSql.Parameters.Add(new SqlParameter("@p14", sj_putilidad2));
                cmdSql.Parameters.Add(new SqlParameter("@p15", sj_putilidad3));
                cmdSql.Parameters.Add(new SqlParameter("@p16", sj_feulactualiza));
            }
            else
            {
                sql = "update saprod set " + updateset + " where " + condi;

                cmdSql.CommandText = sql;
                cmdSql.Connection = misclases.conexSql;
            }

            try
            {
                cmdSql.ExecuteNonQuery();
            }
            catch (SqlException err)
            {
                MessageBox.Show(err.Message.ToString());
            }
        }

        public DataTable select(string condi)
        {
            DataTable dt1 = new DataTable();
            string    sql = "select * from saprod " +
                            "where " + condi;

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(sql, misclases.conexSql);
                da.Fill(dt1);
                da.Dispose();
            }
            catch (SqlException err)
            {
                MessageBox.Show(err.Message.ToString());
            }

            return dt1;
        }

    }
}
