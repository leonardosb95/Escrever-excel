using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscreverNoExcel.Controllers
{
    public class UltimaLinhaExcel
    {

        public static int ultimaLinha()
        {
            FileInfo caminhoNomeArquivo = new FileInfo(@"C:\Users\lbispo\Desktop\Usuario.xlsx");
            ExcelPackage arquivoExcel = new ExcelPackage(caminhoNomeArquivo);
            ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets["Usuario"];
            int linha = 2, coluna = 1;
            int cont = 0;



            for (int i = 2; i < 300; i++)
            {
                if (planilha.Cells[linha, coluna].Value != null)
                {
                    cont++;

                }
                linha++;

            }
            cont += 1;//Proxima linha sem valor que pode ser preenchida

            return cont;
        }




    }
}
