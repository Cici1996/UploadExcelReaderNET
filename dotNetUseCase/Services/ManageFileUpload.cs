using dotNetUseCase.DataAccess.Interfaces;
using dotNetUseCase.Helpers;
using dotNetUseCase.Services.Interfaces;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace dotNetUseCase.Services
{
    public class ManageFileUpload : IManageFileUpload
    {
        private readonly IManageFileRepository _manageFileRepository;

        public ManageFileUpload(IManageFileRepository manageFileRepository)
        {
            _manageFileRepository = manageFileRepository;
        }

        public async Task ReadyExcelFile(string pathFile, IFormFile upload)
        {
            try
            {
                using var fileStream = new FileStream(pathFile, FileMode.Create);
                string fileName = ExcelHelper.MakeValidName(Path.GetFileNameWithoutExtension(upload.FileName));
                byte[] fileBytes = null;
                using (var ms = new MemoryStream())
                {
                    await upload.CopyToAsync(fileStream);
                    await upload.CopyToAsync(ms);
                    fileBytes = ms.ToArray();
                };

                Stream stream = new MemoryStream(fileBytes);

                using var package = new ExcelPackage(stream);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                CreateTableAndInsertData(worksheet, fileName);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        private void CreateTableAndInsertData(ExcelWorksheet worksheet, string fileName)
        {
            int columnCount = ExcelHelper.GetTotalCellCountByAnyNonNullData(worksheet);
            List<string> columnList = new List<string>();
            for (int row = 1; row <= columnCount; row++)
            {
                string columnName = ExcelHelper.MakeValidName(worksheet.Cells[1, row].Value?.ToString().Trim().ToLower());
                columnList.Add(columnName);
            }
            _manageFileRepository.CreateTable(fileName, columnList);
            DataTable dataTableExcel = ExcelHelper.GetDataTableFromExcel(worksheet);
            _manageFileRepository.InsertBulkData(fileName, dataTableExcel);
        }

        public DataTable SelectData(string tableName)
        {
            return _manageFileRepository.SelectData(tableName);
        }
    }
}