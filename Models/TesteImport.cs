using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Teste.Models
{
    public class TesteImport
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public int Id { get; set; }


        public TesteImport() 
        {
        }

        public TesteImport(string name, string email)
        {
            Name = name;
            Email = email;
        }

        public TesteImport(string name, string email, int id)
        {
            Name = name;
            Email = email;
            Id = id;
        }
    }
}
