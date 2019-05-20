using EscreverNoExcel.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscreverNoExcel.Controllers
{
    public class EscreveNoExcel
    {
        public  static void EscreverNoExcelManual(Usuario usuario)
        {
            FileInfo caminhoNomeArquivo = new FileInfo(@"C:\Users\lbispo\Desktop\Usuario.xlsx");
            ExcelPackage arquivoExcel = new ExcelPackage(caminhoNomeArquivo);
            ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets["Usuario"];

            int linha = 2, coluna = 1;

            if (UltimaLinhaExcel.ultimaLinha()>1)
            {

                linha = UltimaLinhaExcel.ultimaLinha();

            }

            planilha.Cells[linha, coluna++].Value = usuario.Documento;
            planilha.Cells[linha, coluna++].Value = usuario.Dt_nasc;
            planilha.Cells[linha, coluna++].Value = usuario.Ani;
            linha++;


            //SALVAR E FECHAR A PLANILHA
            arquivoExcel.Save();
            arquivoExcel.Dispose();
        }

    }
}
