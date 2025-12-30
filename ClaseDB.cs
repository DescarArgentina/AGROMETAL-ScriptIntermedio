using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace PruebaArbol
{
    class DataRow //Clase para guardar datos de los nodos
    {
        public string NombreNodo { get; set; }
        public List<string> Atributos { get; set; }
        public XmlNode XmlNode { get; set; }
    }
    class OccurrenceData //Clase para crear Excel o TXT
    {
        public int Id { get; }
        public string InstancedRef { get; }
        public string ParentRef { get; }
        public string Name { get; }
        public string MasterRef { get; }
        public string Product { get; set; }
        public string ParentProduct { get; set; }
        public int InstancedRefCount { get; set; }
        public int Conteo { get; set; }
        public string ParentProduct2 { get; set; }
        public string Revision { get; set; }
        public string quanti { get; set; }
        public string fantasma { get; set; }
        public string medida { get; set; }
        public OccurrenceData(int id, string instancedRef, string parentRef, string name, string masterRef, string product, string parentProduct, string parentproduct2, string revision, string Quanti, string Fantasma, string Cantidad)
        {
            Id = id;
            InstancedRef = instancedRef;
            ParentRef = parentRef;
            Name = name;
            MasterRef = masterRef;
            Product = product;
            ParentProduct = parentProduct;
            InstancedRefCount = 1;
            Conteo = 10;
            ParentProduct2 = parentproduct2;
            Revision = revision;
            quanti = Quanti;
            fantasma = Fantasma;
            medida = Cantidad;
        }
    }
}
