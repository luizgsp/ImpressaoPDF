﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeracaoColetasPDF.Entities
{
    class Order
    {
        public int SellerCode { get; set; }
        public int OrderNumber { get; set; }
        public DateTime dateTime { get; set; }
        public string SalesPerson { get; set; }
        public string OrderReference { get; set; }
        public Customer Customer { get; set; }
        public string Message { get; set; }
        public List<Items> Items { get; set; } = new List<Items>();
        public double LaborValue { get; set; }
        public List<Notes> Notes { get; set; } = new List<Notes>();

        public Order()
        {
        }

        public double GetTotais()
        {
            double sum = 0;
            foreach (Items item in Items)
            {
                sum += item.GetTotalPrice();
            }
            return sum;
        }
    }
}
