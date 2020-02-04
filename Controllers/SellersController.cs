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
        //private readonly ImportService _importService;

        public SellersController(SellerService sellerService, DepartmentService departmentService/*, ImportService importService*/) 
        {
            _sellerService = sellerService;
            _departmentService = departmentService;
           // _importService = importService;
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
            //return RedirectToAction(nameof(Index));
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


        /*public async Task<IActionResult> Import()
        {
            var departments = await _departmentService.FindAllAsync();
            var viewModel = new ImportViewModel { Departments = departments };
            return View(viewModel);
        }*/


        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public async Task<IActionResult> Import(IFormFile excelFile, int i)
        //{
            
          //  return await Create();
        //}

        //[HttpPost]
        //[ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile excelFile,  int i, int linha)
        {
            //if (excelFile != null)
            //{
                if (linha < 2) { linha = 2; }
                var vetNome = await _sellerService.ListaAsync(excelFile, "A", linha);
                var vetEmail = await _sellerService.ListaAsync(excelFile, "B", linha);
                var seller = new Seller(vetNome[i], vetEmail[i]);
                
                //var nextName = vetNome[i + 1];
                //var nextEmail = vetEmail[i + 1];
                var departments = await _departmentService.FindAllAsync();
                var viewModel = new ImportViewModel { Seller = seller, Departments = departments, i = i, Linha = linha, ExcelFile = planilha };
                return View(viewModel);

                /*await _sellerService.InsertAsync(seller);

                i++;
                linha++;
                return await Import(excelFile, nome, email, i, linha);*/
           // }
            /*else
            {
                if (linha < 2) { linha = 2; }
                var seller2 = new Seller(nome[i], email[i]);
                var departments2 = await _departmentService.FindAllAsync();
                var viewModel2 = new ImportViewModel { Seller = seller2, Departments = departments2, ExcelFile = excelFile, ListNome = nome, ListEmail = email, i = i, Linha = linha };
                return View(viewModel2);
                //return View("Import");

                /*await _sellerService.InsertAsync(seller2);

                i++;
                linha++;
                return await Import(excelFile, nome, email, i, linha);

            }*/


            //return RedirectToAction(nameof(Index));

        }
        public async Task<IActionResult> CreateImport()
        {
            var departments = await _departmentService.FindAllAsync();
            var viewModel = new ImportViewModel { Departments = departments };
            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CreateImport(Seller seller, IFormFile excelFile, int i, int linha)
        {
            await _sellerService.InsertAsync(seller);

            i++;
            linha++;
            return await Import(excelFile, i, linha);
        }

    }


}
