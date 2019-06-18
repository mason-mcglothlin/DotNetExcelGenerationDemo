using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace DotNetExcelGenerationDemo
{
	class EPPlusDemo : IDemo
	{
		public void GenerateAndOpenFile(List<Product> products)
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Products");

				sheet.Cells["A1"].Value = "Name";
				sheet.Cells["B1"].Value = "Description";
				sheet.Cells["C1"].Value = "Release Date";
				sheet.Cells["D1"].Value = "Price";
				sheet.Cells["E1"].Value = "Quantity In Stock";

				for (int i = 0; i < products.Count; i++)
				{
					var product = products[i];
					var row = i + 2;
					sheet.Cells[row, 1].Value = product.Name;
					sheet.Cells[row, 2].Value = product.Description;
					sheet.Cells[row, 3].Value = product.ReleaseDate;
					sheet.Cells[row, 4].Value = product.Price;
					sheet.Cells[row, 5].Value = product.QuantityInStock;
				}

				sheet.Cells["C1:C5"].Style.Numberformat.Format = "m/d/yyyy";
				sheet.Cells["D1:D5"].Style.Numberformat.Format = "$#,##0.00";
				sheet.Cells["E1:E5"].Style.Numberformat.Format = "#,##0";

				var range = sheet.Cells[1, 1, products.Count + 1, 5];
				var table = sheet.Tables.Add(range, "ProductsTable");
				table.TableStyle = TableStyles.Medium8;
				range.AutoFitColumns();

				var bytes = package.GetAsByteArray();
				const string filePath = @"C:\temp\EPPlusDemo.xlsx";
				File.WriteAllBytes(filePath, bytes);
				Process.Start(filePath);
			}
		}
	}
}
