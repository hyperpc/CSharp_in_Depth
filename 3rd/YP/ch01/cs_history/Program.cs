using System.Xml.Linq;
using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
//dotnet add package System.Configuration.ConfigurationManager --version 4.7.0
namespace cs_history
{
    class Program
    {
        static void Main(string[] args)
        {
            //ProductsSort();
            //ProductsQuery();

            //LinqToObject();
            //LinqToXml();

            SaveToExcelViaCOM();

            Console.ReadLine();
        }

        /// <summary>
        /// Not work on my local
        /// </summary>
        private static void SaveToExcelViaCOM(){
            var app = new Application();
            if(app==null){
                Console.WriteLine("Not found installed Office.Excel!");
            }
            app.Visible=true;
            //var wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            //var sheet = (Worksheet)wb.Worksheets[1];
            var wb = app.Workbooks.Add();
            var sheet = (Worksheet)app.ActiveSheet;
            sheet.Name = "Products";
            int row = 1;
            var products = Product.GetSampleProducts();
            var suppliers = Supplier.GetSampleSuppliers();
            var filtered = from p in products
                           join s in suppliers on p.SupplierId equals s.SupplierId
                           where p.Price.HasValue
                           orderby s.Name, p.Name
                           select new { SupplierName = s.Name, ProductName = p.Name, ProductPrice=p.Price.Value };
            foreach (var product in filtered)
            {
                sheet.Cells[row, 1] = product.SupplierName;
                sheet.Cells[row, 2] = product.ProductName;
                sheet.Cells[row, 3] = product.ProductPrice;
                row++;
            }

            //var filepath = ConfigurationManager.AppSettings["filepath"];
            //choose diff code, dependence on your local excel version
            //wb.SaveAs(Filename: (filepath + "product.xls"), FileFormat: XlFileFormat.xlWorkbookNormal);
            //wb.SaveAs2(Filename: (filepath + "product.xlsx"), FileFormat: XlFileFormat.xlOpenXMLWorkbook);
            wb.SaveAs(Filename: "product.xls",  FileFormat: XlFileFormat.xlWorkbookNormal);
            wb.Close();
            app.Application.Quit();
        }

        private static void LinqToXml()
        {
            var doc = XDocument.Load("App_Data/products.xml");
            var filtered = from p in doc.Descendants("Product")
                           join s in doc.Descendants("Supplier") on (int)p.Attribute("SupplierId") equals (int)s.Attribute("SupplierId")
                           where (decimal)p.Attribute("Price") > 10
                           orderby s.Attribute("Name").ToString(), p.Attribute("Name").ToString()
                           select new { SupplierName = s.Attribute("Name").ToString(), ProductName = p.Attribute("Name").ToString() };
            foreach (var v in filtered)
            {
                Console.WriteLine("Supplier={0}; Product={1}", v.SupplierName, v.ProductName);
            }
        }

        private static void LinqToObject()
        {
            var products = Product.GetSampleProducts();
            var suppliers = Supplier.GetSampleSuppliers();
            var filtered = from p in products
                           join s in suppliers on p.SupplierId equals s.SupplierId
                           where p.Price > 10
                           orderby s.Name, p.Name
                           select new { SupplierName = s.Name, ProductName = p.Name };
            foreach (var v in filtered)
            {
                Console.WriteLine("Supplier={0}; Product={1}", v.SupplierName, v.ProductName);
            }
        }

        private static void ProductsQuery()
        {
            var products = Product.GetSampleProducts();
            foreach (var product in products.Where(p => p.Price.HasValue && p.Price > 10))
            {
                Console.WriteLine($"{product.Name}: {product.Price.Value}");
            }
        }

        private static void ProductsSort()
        {
            var products = Product.GetSampleProducts();
            foreach (var product in products.OrderBy(p => p.Name))
            {
                Console.WriteLine(product.Name);
            }
        }
    }
}
