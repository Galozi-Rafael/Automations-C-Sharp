using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;

namespace ExtrairExcel.Services
{
    internal class ExcelService
    {
        public void ExportToExcel(List<List<string>> tableData, string outputPath)
        {
            // Garante que a pasta de destino exista.
            var folder = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(folder))
            {
                Directory.CreateDirectory(folder);
            }

            // Criando o arquivo Excel.
            using var workbook = new XLWorkbook();

            var sheet = workbook.Worksheets.Add("Tabela");

            // Preenchendo os dados na planilha.
            for (int rowIndex = 0; rowIndex < tableData.Count; rowIndex++)
            {
                var row = tableData[rowIndex];

                for (int colIndex = 0; colIndex < row.Count; colIndex++)
                {
                    var cellValue = row[colIndex];

                    sheet.Cell(rowIndex + 1, colIndex + 1).Value = cellValue;
                }
            }

            sheet.Columns().AdjustToContents();

            workbook.SaveAs(outputPath);
        }
    }
}
