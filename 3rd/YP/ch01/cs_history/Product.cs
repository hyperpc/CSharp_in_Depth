using System.Collections.Generic;
public class Product{
    readonly string name;
    public string Name{get{return name;}}
    readonly decimal? price;
    public decimal? Price{get{return price;}}
    readonly int supplierId;
    public int SupplierId{get{return supplierId;}}
    public Product(string name, int supplierId, decimal? price=null){
        this.name=name;
        this.supplierId=supplierId;
        this.price=price;
    }
    public static List<Product> GetSampleProducts(){
        return new List<Product>{
            new Product(name:"West Side Story", supplierId:1, price: 9.99m),
            new Product(name:"Assassins", supplierId:2, price: 14.99m),
            new Product(name:"Frogs", supplierId:1, price: 13.99m),
            //new Product("Frogs",supplierId:1),
            new Product(name:"Sweeney Todd", supplierId:3, price: 10.99m)
        };
    }
    public override string ToString()
    {
        return string.Format("{0}: {1}", name, price.HasValue?price.Value.ToString():"null");
    }
}