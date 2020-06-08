using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace WebApplication3.Controllers
{

    [ApiController]
    public class DemoController : ControllerBase
    {
        private readonly IHostingEnvironment _hostingEnvironment;
        public DemoController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public class UserInfo
        {
            public string UserName { get; set; }

            public int Age { get; set; }
        }
        [HttpGet]
        [Route("api/export")]
        public async Task<string> Export(CancellationToken cancellationToken)
        {
            string folder = _hostingEnvironment.WebRootPath;
            string excelName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            string downloadUrl = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, excelName);
            FileInfo file = new FileInfo(Path.Combine(folder, excelName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(folder, excelName));
            }

            // query data from database  
            await Task.Yield();

            var list = new List<UserInfo>()
                    {
                        new UserInfo { UserName = "catcher", Age = 18 },
                        new UserInfo { UserName = "james", Age = 20 },
                    };

            using (var package = new ExcelPackage(file))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromCollection(list, true);
                package.Save();
            }

            return downloadUrl;
        }
    }
}
