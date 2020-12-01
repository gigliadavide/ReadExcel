using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ReadExcel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ReadExcel.Controllers
{
    public class UserController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            return View(new List<UserModel>());
        }

        [HttpPost]
        public IActionResult Index(IFormCollection form)
        {
            var filtro = Request.Form["filtro"];
            List<UserModel> users = new List<UserModel>();
            var fileName = "../Users.xlsx";


            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        users.Add(new UserModel
                        {
                            Name = reader.GetValue(0).ToString(),
                            Email = reader.GetValue(1).ToString(),
                            Phone = reader.GetValue(2).ToString()
                        });
                    }
                }
            }
                        
            if (string.IsNullOrEmpty(filtro))
            {
                return View(users);
            }
            else
            {
                return View(users.Where(c => c.Name == filtro || 
                                             c.Phone == filtro || 
                                             c.Email == filtro));
            }
        }
    }

}
