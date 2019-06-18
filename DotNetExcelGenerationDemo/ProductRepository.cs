using System;
using System.Collections.Generic;

namespace DotNetExcelGenerationDemo
{
	class ProductRepository
	{
		public List<Product> GetAllProducts()
		{
			return new List<Product>
			{
			     new Product { Name = "La Croix", Description = "Carbonated water beverage", Price = 2.50m, QuantityInStock = 7000, ReleaseDate = DateTime.Now },
				 new Product { Name = "Logitech M510", Description = "Wireless optical mouse", Price = 30m, QuantityInStock = 17, ReleaseDate = new DateTime(2000, 1, 1) },
				 new Product { Name = "Gigabyte Aero 15", Description = "Colorful laptop", Price = 2000m, QuantityInStock = 3, ReleaseDate = new DateTime(2018, 7, 4) },
				 new Product { Name = "Labrador Retriever", Description = "Fun and cute puppy", Price = 500m, QuantityInStock = 8 }
			};
		}
	}
}
