using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SapLogisticAutomatizaion
{
    public class Product
    {
        public string PartNumber { get; set; }
        public string MaterialDescription { get; set; }
        public string SerialNumber { get; set; }
        public DateTime ManufacturingDate { get; set; }
        public DateTime ReceiptDate { get; set; }
        public string AdditionalData { get; set; }
        public string ContainerCondition { get; set; }
        public string CustomsStatus { get; set; }
        public string ContainerDetails { get; set; }
    }
}