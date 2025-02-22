using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CATIAAssistant.Models
{
    public class CatiaDocResult
    {
        public string DocType { get; set; }  // "DrawingDocument" / "ProductDocument" / "PartDocument"
        public INFITF.Document ActiveDoc { get; set; }
        public DRAFTINGITF.DrawingDocument DrawingDoc { get; set; }
        public ProductStructureTypeLib.ProductDocument ProductDoc { get; set; }
        public MECMOD.PartDocument PartDoc { get; set; }
    }
}
