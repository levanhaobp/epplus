// See https://aka.ms/new-console-template for more information
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using TestEpplusExel;


var bytyFile = File.ReadAllBytes("./Template_PA.xlsx");
await ReadFromExcel<PreAprovedExcelModel>(bytyFile);

async Task<T> ReadFromExcel<T>(byte[] byteFile, bool hasHeader = true, CancellationToken cancellationToken = default)
{
    try
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var stream = new MemoryStream(byteFile);
        using (var excelPack = new ExcelPackage(stream))
        {
            await excelPack.LoadAsync(stream);
            var ws = excelPack.Workbook.Worksheets[0];
            var wsSettings = excelPack.Workbook.Worksheets[1];

            DataTable excelasTable = new DataTable();
            var startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= wsSettings.Dimension.End.Row; rowNum++)
            {
                var firstRowCell = wsSettings.Cells[rowNum, 1];
                if (!string.IsNullOrEmpty(firstRowCell?.Text))
                {
                    string firstColumn = string.Format("Column {0}", firstRowCell.Start.Column);
                    excelasTable.Columns.Add(hasHeader ? firstRowCell.Text : firstColumn);
                }
            }

            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, excelasTable.Columns.Count];
                DataRow row = excelasTable.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(excelasTable));
        }
    }
    catch (Exception ex)
    {
        return default(T);
    }
}