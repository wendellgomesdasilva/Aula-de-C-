using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Teste.Models;
using Microsoft.EntityFrameworkCore;
using Teste.Services.Exceptions;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using System.IO;

namespace Teste.Services
{
    public class SellerService
    {
        private readonly TesteContext _context;

        public SellerService(TesteContext context)
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

        public async Task RemoveAsync(int id)
        {
            try
            {
                var obj = await _context.Seller.FindAsync(id);
                _context.Seller.Remove(obj);
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateException e)
            {
                throw new IntegrityException("Can't delete seller because he/she has sales");
            }

        }

        public async Task UpdateAsync(Seller obj)
        {
            bool hasAny = await _context.Seller.AnyAsync(x => x.Id == obj.Id);
            if (!hasAny)
            {
                throw new NotFoundException("Id not found");
            }
            try
            {
                _context.Update(obj);
                await _context.SaveChangesAsync();
            }
            catch(DbUpdateConcurrencyException e)
            {
                throw new DbConcurrencyException(e.Message);
            }

        }


        


        public async Task<string[]> ListaAsync(IFormFile excelFile, string x, int linha)
        {

            List<string> listNome = new List<string>();
            string str = Convert.ToString(excelFile.FileName);
            bool b1 = string.IsNullOrEmpty(str);
            //FileInfo path = excelFile;

            //Confere se algum arquivo foi selecionado
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


                    //var linha = 2;
                    
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


            }

            return null;
        }

        

    }
}
