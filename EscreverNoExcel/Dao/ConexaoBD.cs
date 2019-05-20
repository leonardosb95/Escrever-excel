using EscreverNoExcel.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscreverNoExcel
{
    public class ConexaoBD
    {


        public static SqlConnection conexao()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["conexao"].ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);
            return connection;
        }




        

       


    }
}
