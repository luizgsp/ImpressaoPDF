﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeracaoColetasPDF.Entities
{
    class Customer
    {
        public string Name { get; set; }
        public string CnpjCpf { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Uf { get; set; }
        public string PostalCode { get; set; }
        public List<string> Phones { get; set; } = new List<string>();
        public string Email { get; set; }
        public string Contact { get; set; }

        public Customer(string name, string cnpjCpf, string address, string city, string uf, string postalCode, string email, string contact)
        {
            Name = name;
            CnpjCpf = cnpjCpf;
            Address = address;
            City = city;
            Uf = uf;
            PostalCode = postalCode;
            Email = email;
            Contact = contact;
        }
    }
}
