using System.IO;
using System.Data;
using System.Data.SQLite;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;

namespace calprice
{

    class misclases
    {

        public static SQLiteConnection  conexLite = new SQLiteConnection();
        public static SqlConnection     conexSql  = new SqlConnection();

        public static string namecompany;

        public static SQLiteConnection Conexion_Sqlite()
        {
            SQLiteConnection conexion = new SQLiteConnection("Data Source=datapp.db");

            try
            {
                conexion.Open();
            }
            catch (SQLiteException err)
            {
                MessageBox.Show(err.Message.ToString());
            }

            return conexion;
        }

        public static SqlConnection Conexion_Sql()
        {
            SqlConnection conexion = new SqlConnection();
            string        cadconex  = ConfigurationManager.ConnectionStrings["admin"].ToString();

            conexion.ConnectionString = cadconex;

            try
            {
                conexion.Open();
            }
            catch (SqlException err)
            {
                MessageBox.Show(err.Message.ToString());
            }

            return conexion;
        }

        public static DataTable CursorTable(string texSql,
                                            SqlConnection myConexSql = null,
                                            bool msgerr = true)
        {
            DataTable dt = new DataTable();

            myConexSql = (myConexSql == null ? misclases.conexSql : myConexSql);

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(texSql, myConexSql);
                da.Fill(dt);
                da.Dispose();
            }
            catch (SqlException err)
            {
                if (msgerr) MessageBox.Show(err.Message.ToString());
            }

            return dt;
        }

        public static string FileNameTemp(string extension)
        {
            string fileRandom = Path.GetRandomFileName();
            int i = fileRandom.IndexOf(".");

            if (extension != "" && i >= 0)
            {
                fileRandom = fileRandom.Substring(0, (i + 1)) + extension;
            }

            fileRandom = Path.Combine(Path.GetTempPath(), fileRandom);

            return fileRandom;
        }

    }
}
