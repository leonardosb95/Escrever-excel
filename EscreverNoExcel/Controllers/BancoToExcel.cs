using EscreverNoExcel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace EscreverNoExcel.Controllers
{
    public class BancoToExcel
    {
        public static void JogaDadosDoBancoNoExcel(List<Usuario> usuarioLista)
        {
            
            Usuario user = new Usuario();

            FileInfo caminhoNomeArquivo = new FileInfo(@"C:\Users\lbispo\Desktop\Usuario.xlsx");
            ExcelPackage arquivoExcel = new ExcelPackage(caminhoNomeArquivo);
            //ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets.Add("Usuario");//Nome da aba
            ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets["Usuario"];
            //planilha.Cells["A1"].Value = "Livraria jujubinhas";//Escrever na linha 1 do excel

                int linha = 1, coluna = 1;
                //CRIANDO O CABEÇALHO
                planilha.Cells[linha, coluna++].Value = "Documento";
                planilha.Cells[linha, coluna++].Value = "Data nascimento";
                planilha.Cells[linha, coluna++].Value = "Ani";

                linha = 2;
                coluna = 1;


            if (UltimaLinhaExcel.ultimaLinha() > 1)
            {

                linha = UltimaLinhaExcel.ultimaLinha();
            }
           


                foreach (var usuario in usuarioLista)
                {
                    //ESCREVENDO NA LINHA DOS CABEÇALHOS
                    planilha.Cells[linha, coluna++].Value = usuario.Documento;
                    planilha.Cells[linha, coluna++].Value = usuario.Dt_nasc;
                    planilha.Cells[linha, coluna++].Value = usuario.Ani;
                    linha++;
                    coluna = 1;

                }

                //SALVAR E FECHAR A PLANILHA
                arquivoExcel.Save();
                arquivoExcel.Dispose();
        }

    }
}
