namespace DotNetExcelGenerationDemo
{
	class Program
	{
		static void Main(string[] args)
		{
			var productRepository = new ProductRepository();
			var products = productRepository.GetAllProducts();
			var demo = new GridViewDemo();
			//var demo = new EPPlusDemo();
			demo.GenerateAndOpenFile(products);
		}
	}
}
