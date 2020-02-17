using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Teste.Models;
using Microsoft.EntityFrameworkCore;
using Teste.Services.Exceptions;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Teste.Models.ViewModels;
using Teste.Services;
using System.IO;

namespace Teste.Services
{
    public class ImportService
    {
        private readonly TesteContext _context;

        public ImportService(TesteContext context)
        {
            _context = context;
        }

        public async Task<List<Seller>> FindAllAsync()
        {
            return await _context.Seller.ToListAsync();
        }

        public async Task InsertAsync(Seller obj)
        {
            _context.Add(obj);
            await _context.SaveChangesAsync();
        }

        public async Task<Seller> FindByIdAsync(int id)
        {
            return await _context.Seller.Include(obj => obj.Department).FirstOrDefaultAsync(obj => obj.Id == id);
        }


        public async Task<IXLWorksheet> Directory(IFormFile excelFile)
        {
            
            string str = Convert.ToString(excelFile.FileName);
            bool b1 = string.IsNullOrEmpty(str);

            if (b1 == true)
            {
                return null;
            }
            else
            {
                if (Path.GetExtension(str) == ".xls" || Path.GetExtension(str) == ".xlsx")
                {
                    var wb = new XLWorkbook("C:\\Users/wgomessi/Desktop/Teste/teste.xlsx");
                    IXLWorksheet planilha = wb.Worksheet(1);
                    return planilha;
                }
            }
            return null;

        }


        public async Task<string[]> ListaAsync(IXLWorksheet planilha, string x)
        {

            List<string> listNome = new List<string>();
            var linha = 2;
            while (true)
            {
                string nome = planilha.Cell(x + linha.ToString()).Value.ToString();

                if (string.IsNullOrEmpty(nome)) break;

                listNome.Add(nome);
                linha++;
            }


            string[] vet = new string[listNome.Count];
            int cont = 0;
            foreach (string y in listNome)
            {
                vet[cont] = y;
                cont++;
            }
                    
            return vet.ToArray();

        }

        public async Task InsertTesteAsync(TesteImport obj)
        {
            _context.Add(obj);
            await _context.SaveChangesAsync();
        }

        public async Task RemoveTesteAsync()
        {
            try
            {
                var obj = _context.TesteImport.Count();
                for (int i = 1; i <= obj; i++)
                {
                    var teste = await _context.TesteImport.FindAsync(i);
                    _context.TesteImport.Remove(teste);
                    await _context.SaveChangesAsync();
                }
            }
            catch (DbUpdateException e)
            {
                throw new IntegrityException("Can't delete seller because he/she has sales");
            }

        }

        public async Task<TesteImport> TesteFindAsync(int i) 
        {
            return TesteImport.Where(t => t.Name = i);
        }
    }
}
