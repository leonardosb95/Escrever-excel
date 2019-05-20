using EscreverNoExcel.Controllers;
using EscreverNoExcel.Dao;
using EscreverNoExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscreverNoExcel
{
    class Program
    {
        static void Main(string[] args)
        {


            AplicacaoBD aplicacao = new AplicacaoBD();


            Console.Write("1-Inserir no excel atraves do banco de dados\n 2-Inserir manual no excel \n--------->");
            string opcao = Console.ReadLine();

            if (opcao.Equals("1"))
            {
                BancoToExcel.JogaDadosDoBancoNoExcel(AplicacaoBD.selectAll());

            }
            else if (opcao.Equals("2"))
            {

                Console.Write("Digite o documento do usuário: ");
                string documento = Console.ReadLine();

                Console.Write("Digite a data de nascimento: ");
                string dt_nasc = Console.ReadLine();

                Console.Write("Digite o numero de telefone do usuário: ");
                string ani = Console.ReadLine();

                var usuario = new Usuario()
                {
                    Documento = documento,
                    Dt_nasc = dt_nasc,
                    Ani = ani
                };
                EscreveNoExcel.EscreverNoExcelManual(usuario);

            }
            else
            {
                Console.WriteLine("Opcao invalida");
            }


            



            

         



           
           







        }
    }
}
