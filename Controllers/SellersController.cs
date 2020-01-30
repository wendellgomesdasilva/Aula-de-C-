using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Teste.Models;
using Teste.Models.ViewModels;
using Teste.Services;
using Teste.Services.Exceptions;
using Microsoft.AspNetCore.Http;
using Excel = Microsoft.Office.Interop.Excel;
using IronXL;
using ClosedXML.Excel;
using System.IO;
using System.Linq;
using System.Web;
using Microsoft.AspNetCore.Server;
using System.Web.Http;



namespace Teste.Controllers
{
    public class SellersController : Controller
    {
        private readonly SellerService _sellerService;
        private readonly DepartmentService _departmentService;

        
        public SellersController(SellerService sellerService, DepartmentService departmentService) 
        {
            _sellerService = sellerService;
            _departmentService = departmentService;
        }

        public async Task<IActionResult> Index()
        {
            var list = await _sellerService.FindAllAsync();
            return View(list);
        }

        public async Task<IActionResult> Create() 
        {
            var departments = await _departmentService.FindAllAsync();
            var viewModel = new SellerFormViewModel { Departments = departments };
            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Seller seller)
        {
            
            if (!ModelState.IsValid)
            {
                var departments = await _departmentService.FindAllAsync();
                var viewModel = new SellerFormViewModel { Seller = seller, Departments = departments };
                return View(viewModel);
            }
            await _sellerService.InsertAsync(seller);
            return RedirectToAction(nameof(Index));

        }

        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return RedirectToAction(nameof(Error), new { message = "Id not provided"});
            }

            var obj = await _sellerService.FindByIdAsync(id.Value);
            if (obj == null)
            {
                return RedirectToAction(nameof(Error), new { message = "Id not found" });
            }

            return View(obj);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            try
            {
                await _sellerService.RemoveAsync(id);
                return RedirectToAction(nameof(Index));
            }
            catch (IntegrityException e)
            {
                return RedirectToAction(nameof(Error), new { message = e.Message });
            }
        }

        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return RedirectToAction(nameof(Error), new { message = "Id not provided" });
            }

            var obj = await _sellerService.FindByIdAsync(id.Value);
            if (obj == null)
            {
                return RedirectToAction(nameof(Error), new { message = "Id not found" });
            }

            return View(obj);
        }

        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return RedirectToAction(nameof(Error), new { message = "Id not provided" }); ;
            }

            var obj = await _sellerService.FindByIdAsync(id.Value);
            if (obj == null)
            {
                return RedirectToAction(nameof(Error), new { message = "Id not found" });
            }
            List<Department> departments = await _departmentService.FindAllAsync();
            SellerFormViewModel viewModel = new SellerFormViewModel { Seller = obj, Departments = departments };
            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, Seller seller)
        {
            if (!ModelState.IsValid)
            {
                var departments = await _departmentService.FindAllAsync();
                var viewModel = new SellerFormViewModel { Seller = seller, Departments = departments };
                return View(viewModel);
            }
            if (id != seller.Id)
            {
                return RedirectToAction(nameof(Error), new { message = "Id mismatch" });
            }
            try
            {
                await _sellerService.UpdateAsync(seller);
                return RedirectToAction(nameof(Index));
            }
            catch (ApplicationException e)
            {
                return RedirectToAction(nameof(Error), new { message = e.Message });
            }
        }

        public IActionResult Error(string message)
        {
            var viewModel = new ErrorViewModel
            {
                Message = message,
                RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier
            };
            return View(viewModel);
        }


        public async Task<IActionResult> Import()
        {
            var departments = await _departmentService.FindAllAsync();
            var viewModel = new ImportViewModel { Departments = departments };
            return View(viewModel);
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            List<string> listNome = new List<string>();
            List<string> listEmail = new List<string>();
            string str = Convert.ToString(excelFile.FileName);
            bool b1 = string.IsNullOrEmpty(str);
            //FileInfo path = excelFile;
            
            //Confere se algum arquivo foi selecionado
            if(b1 == true)
            {
                
                return RedirectToAction(nameof(Create));
            }
            else
            {
                if (Path.GetExtension(str) == ".xls" || Path.GetExtension(str) == ".xlsx")
                {
                    var wb = new XLWorkbook("C:\\Users/wgomessi/Desktop/Teste/teste.xlsx");
                    var planilha = wb.Worksheet(1);

                    var linha = 2;
                    while (true)
                    {
                        string nome = planilha.Cell("A" + linha.ToString()).Value.ToString();

                        if (string.IsNullOrEmpty(nome)) break;

                        string email = planilha.Cell("B" + linha.ToString()).Value.ToString();

                        listNome.Add(nome);
                        listEmail.Add(email);
                        linha++;
                    }


                    string[] vetNome = new string[listNome.Count];
                    string[] vetEmail = new string[listEmail.Count];
                    int cont = 0;
                    foreach (string x in listNome)
                    {
                        vetNome[cont] = x;
                        cont++;
                    }
                    cont = 0;
                    foreach (string y in listEmail)
                    {
                        vetEmail[cont] = y;
                        cont++;
                    }
                    

                    for (int i = 0; i <= vetEmail.Length; i++)
                    {
                        var seller = new Seller(vetNome[i], vetEmail[i]);

                        
                        var departments = await _departmentService.FindAllAsync();
                        var viewModel = new ImportViewModel { Seller = seller, Departments = departments };
                        return View(viewModel);
                        
                        //await _sellerService.InsertAsync(seller);
                        //return RedirectToAction(nameof(Index));


                    }
                    return RedirectToAction(nameof(Index));


                    /* var seller = new Seller(listNome[0], listEmail[0]);
                     var departments = await _departmentService.FindAllAsync();
                     ImportViewModel viewModel = new ImportViewModel { Seller = seller, Departments = departments, ListNome = listNome, ListEmail = listEmail };
                     return View(viewModel);*/


                }
               
                
            }
            return await Create();
        }
        public async Task<IActionResult> CreateImport()
        {
            var departments = await _departmentService.FindAllAsync();
            var viewModel = new ImportViewModel { Departments = departments };
            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CreateImport(string[] vetNome, string[] vetEmail)
        {
            for (int i = 0; i <= vetEmail.Length; i++)
            {
                var seller = new Seller(vetNome[i], vetEmail[i]);

                var departments = await _departmentService.FindAllAsync();
                ImportViewModel viewModel = new ImportViewModel { Seller = seller, Departments = departments };
                await _sellerService.InsertAsync(seller);
                return View(viewModel);
                    
                
            }
            return RedirectToAction(nameof(Index));
        }

    }


}
