// See https://aka.ms/new-console-template for more information
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using TestEpplusExel;


var bytyFile = File.ReadAllBytes("./Template_PA.xlsx");
var b = await ReadFromExcel<List<PreAprovedExcelModel>>(bytyFile);
int c = 1;

async Task<T> ReadFromExcel<T>(byte[] byteFile, bool hasHeader = true, CancellationToken cancellationToken = default)
{
    try
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var stream = new MemoryStream(byteFile);
        using (var excelPack = new ExcelPackage(stream))
        {
            await excelPack.LoadAsync(stream, cancellationToken);
            var ws = excelPack.Workbook.Worksheets[0];
            var wsSettings = excelPack.Workbook.Worksheets[1];

            DataTable excelasTable = new DataTable();
            var startRow = hasHeader ? 2 : 1;
            var endRow = wsSettings.Dimension.End.Row;

            var dictionary = new Dictionary<string, string>();
            for (int i = startRow; i <= endRow; i++)
            {
                excelasTable.Columns.Add(wsSettings.Cells[i, 1].Value + string.Empty);
                dictionary.Add(wsSettings.Cells[i, 1].Value + string.Empty, wsSettings.Cells[i, 2].Value + string.Empty);
            }

            //var temp = JsonConvert.SerializeObject(dictionary);
            //var temp1 = JsonConvert.DeserializeObject<PreAprovedExcelModel>(temp);

            //for (int rowNum = startRow; rowNum <= wsSettings.Dimension.End.Row; rowNum++)
            //{
            //    var firstRowCell = wsSettings.Cells[rowNum, 1];
            //    if (!string.IsNullOrEmpty(firstRowCell?.Text))
            //    {
            //        string firstColumn = string.Format("Column {0}", firstRowCell.Start.Column);
            //        excelasTable.Columns.Add(hasHeader ? firstRowCell.Text : firstColumn);
            //    }
            //}

            endRow = ws.Dimension.End.Row;
            //for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            //{
            //    var wsRow = ws.Cells[rowNum, 1, rowNum, excelasTable.Columns.Count];
            //    DataRow row = excelasTable.Rows.Add();
            //    foreach (var cell in wsRow)
            //    {
            //        row[cell.Start.Column - 1] = cell.Text;
            //    }
            //}
            for (int i = startRow; i <= endRow; i++)
            {
                var dRow = excelasTable.Rows.Add();
                foreach (var item in dictionary)
                {
                    dRow[item.Key] = ws.Cells[$"{item.Value}{i}"].Value + string.Empty;
                }
                //excelasTable.Rows.Add(dRow);
            }

            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(excelasTable));
            //return default;
        }
    }
    catch (Exception ex)
    {
        return default(T);
    }
}