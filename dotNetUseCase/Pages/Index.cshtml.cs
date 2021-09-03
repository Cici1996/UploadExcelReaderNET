using dotNetUseCase.Constants;
using dotNetUseCase.Helpers;
using dotNetUseCase.Services.Interfaces;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace dotNetUseCase.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly IManageFileUpload _manageFileUpload;

        public IFormFile Upload { get; set; }
        public DataTable DataExcel { get; set; }
        public List<string> ListFile { get; set; } = new List<string>();
        public string MessageError { get; set; }

        public IndexModel(ILogger<IndexModel> logger, IWebHostEnvironment environment, IManageFileUpload manageFileUpload)
        {
            _logger = logger;
            _environment = environment;
            _manageFileUpload = manageFileUpload;
        }

        public void OnGet()
        {
            GetListFile();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                var file = Path.Combine(_environment.ContentRootPath, GeneralConstant.UploadFolder, Upload.FileName);
                await _manageFileUpload.ReadyExcelFile(file, Upload);

                string fileName = ExcelHelper.MakeValidName(Path.GetFileNameWithoutExtension(Upload.FileName));
                DataExcel = _manageFileUpload.SelectData(fileName);
            }
            catch(Exception ex)
            {
                MessageError = ex.Message;
            }
            GetListFile();
            return Page();
        }

        private List<string> GetListFile()
        {
            var file = Path.Combine(_environment.ContentRootPath, GeneralConstant.UploadFolder);
            string[] allfiles = Directory.GetFiles(file);
            foreach (string item in allfiles)
                ListFile.Add(Path.GetFileName(item));

            return ListFile;

        }

        public FileResult OnGetDownloadFile(string fileName)
        {
            string path = Path.Combine(_environment.ContentRootPath, GeneralConstant.UploadFolder, fileName);
            byte[] bytes = System.IO.File.ReadAllBytes(path);
            return File(bytes, "application/octet-stream", fileName);
        }
    }
}