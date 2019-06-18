using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DotNetExcelGenerationDemo
{
	public class GridViewDemo : IDemo
	{
		public void GenerateAndOpenFile(List<Product> products)
		{
			var gridView = new GridView();
			gridView.DataSource = products;
			gridView.DataBind();
			StringWriter sw = new StringWriter();
			HtmlTextWriter htw = new HtmlTextWriter(sw);
			gridView.RenderControl(htw);
			var output = sw.ToString();
			const string filePath = @"C:\temp\GridViewDemo.xls";
			File.WriteAllText(filePath, output);
			Process.Start(filePath);
		}
	}
}
