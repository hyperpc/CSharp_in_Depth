using System.Collections.Generic;
public class Supplier{
    readonly string name;
    public string Name{get{return name;}}
    readonly int supplierId;
    public int SupplierId{get{return supplierId;}}
    public Supplier(string name, int supplierId){
        this.name=name;
        this.supplierId=supplierId;
    }
    public static List<Supplier> GetSampleSuppliers(){
        return new List<Supplier>{
            new Supplier(name:"Solely Sondheim", supplierId: 1),
            new Supplier(name:"CD-by-CD-by-Sondheim", supplierId: 2),
            new Supplier(name:"Barbershop CDs", supplierId: 3)
        };
    }
    public override string ToString()
    {
        return string.Format("{0}: {1}", supplierId, name);
    }
}