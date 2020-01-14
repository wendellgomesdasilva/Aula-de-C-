using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace Teste.Models
{
    public class TesteContext : DbContext
    {
        public TesteContext (DbContextOptions<TesteContext> options)
            : base(options)
        {
        }

        public DbSet<Teste.Models.Department> Department { get; set; }
    }
}
