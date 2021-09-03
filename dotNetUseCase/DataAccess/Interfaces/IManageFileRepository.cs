using System.Collections.Generic;
using System.Data;

namespace dotNetUseCase.DataAccess.Interfaces
{
    public interface IManageFileRepository
    {
        void CreateTable(string tableName, List<string> columns);

        void InsertBulkData(string tableName, DataTable data);

        DataTable SelectData(string tableName);
    }
}