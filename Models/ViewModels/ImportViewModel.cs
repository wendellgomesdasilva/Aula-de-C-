using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Teste.Models.ViewModels
{
    public class ImportViewModel

    {
        public string Nome { get; set; }
        public string Email { get; set; }
        public Seller Seller { get; set; }
        public int Linha { get; set; }
        public int i { get; set; }
        public IXLWorksheet ExcelFile { get; set; }
        public ICollection<Department> Departments { get; set; }

        public ICollection<string>ListNome { get; set; }
        public ICollection<string> ListEmail { get; set; }

        /*public Seller GetSeller()
         {
             return Seller;
         }

         public void SetSeller()
         {
             Seller.Name = Nome;
             Seller.Email = Email;
         }*/
    }
}
