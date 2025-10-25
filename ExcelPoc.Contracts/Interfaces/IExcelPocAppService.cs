using ExcelPoc.Contracts.DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPoc.Contracts.Interfaces
{
    public interface IExcelPocAppService
    {
        Task<byte[]> GenerateExcelAsync(ExcelDto dto);

    }
}
