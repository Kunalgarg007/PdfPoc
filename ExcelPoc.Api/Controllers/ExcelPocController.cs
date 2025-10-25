using ExcelPoc.Application;
using ExcelPoc.Contracts.DTO;
using ExcelPoc.Contracts.Interfaces;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace CctPdfPoc.Api.Controllers
{
    [ApiController]
    [Route("api/pdf")]
    public class ExcelPocController : ControllerBase
    {
        private readonly IExcelPocAppService _service;

        public ExcelPocController(IExcelPocAppService service)
        {
            _service = service;
        }

        [HttpPost("generate")]
        public async Task<IActionResult> Generate([FromBody] ExcelDto dto)
        {
            var pdfBytes = await _service.GenerateExcelAsync(dto);
            return File(pdfBytes, "application/pdf", "NEW ALW ISP BLANK.pdf");
        }
    }
}
