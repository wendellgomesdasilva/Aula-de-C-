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

        public DbSet<Department> Department { get; set; }
        public DbSet<Seller> Seller { get; set; }
        public DbSet<SalesRecord> SalesRecord { get; set; }
        public DbSet<TesteImport> TesteImport { get; set; }
    }
}
