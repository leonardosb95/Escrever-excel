using EscreverNoExcel.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscreverNoExcel.Dao
{
    public class AplicacaoBD
    {
        private readonly ConexaoBD bd = new ConexaoBD();

        public void add(Usuario usuario)
        { 
            string insertInto ="INSERT INTO Leonardo_bispo (documento,dt_nasc,ani) VALUES(@documento,@dt_nasc,@ani)";

            try
            {

                using (SqlConnection conn = ConexaoBD.conexao())
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(insertInto, conn))
                    {
                        cmd.Parameters.AddWithValue("@documento", usuario.Documento);
                        cmd.Parameters.AddWithValue("@dt_nasc", usuario.Dt_nasc);
                        cmd.Parameters.AddWithValue("@ani", usuario.Ani);
                        cmd.ExecuteNonQuery();
                    }

                }
            }
            catch (Exception)
            {
                throw;
            }

            

         }


        public static List<Usuario> selectAll()
        {
            List<Usuario> usuarioLista = new List<Usuario>();
            Usuario usuario = new Usuario();
            string strSelectAll = "SELECT * FROM Leonardo_bispo";


            using (SqlConnection conn = ConexaoBD.conexao())
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(strSelectAll, conn);
                SqlDataReader leitor = cmd.ExecuteReader();

                while (leitor.Read())
                {

                    //usuario.Documento=leitor["documento"].ToString();
                    //usuario.Dt_nasc=leitor["dt_nasc"].ToString();
                    //usuario.Ani=leitor["ani"].ToString();

                    usuarioLista.Add(new Usuario {
                        Documento = leitor.GetString(1),
                        Dt_nasc = leitor.GetString(2),
                        Ani = leitor.GetString(3)
                    });
                }
               
            }


            return usuarioLista;


        }



}
}
