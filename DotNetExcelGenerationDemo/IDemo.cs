using System.Collections.Generic;

namespace DotNetExcelGenerationDemo
{
	interface IDemo
	{
		void GenerateAndOpenFile(List<Product> products);
	}
}
