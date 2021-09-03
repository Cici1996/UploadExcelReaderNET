using Microsoft.AspNetCore.Http;
using System.Data;
using System.Threading.Tasks;

namespace dotNetUseCase.Services.Interfaces
{
    public interface IManageFileUpload
    {
        Task ReadyExcelFile(string pathFile, IFormFile Upload);

        DataTable SelectData(string tableName);
    }
}