using System.Collections.Generic;
using System.Linq;
using Teste.Models;

namespace Teste.Services
{
    public class DepartmentService
    {
        private readonly TesteContext _context;

        public DepartmentService(TesteContext context)
        {
            _context = context;
        }

        public List<Department> FindAll()
        {
            return _context.Department.OrderBy(x => x.Name).ToList();
        }
    }
}
