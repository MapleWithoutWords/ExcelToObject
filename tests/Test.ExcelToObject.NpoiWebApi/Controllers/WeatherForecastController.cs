using Microsoft.AspNetCore.Mvc;
using OEM.Core;
using Test.OEM.NpoiWebApi.TestClass;

namespace Test.OEM.NpoiWebApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private readonly IExcelFactory _excelFactory;

        public WeatherForecastController(IExcelFactory excelFactory)
        {
            _excelFactory = excelFactory;
        }

        [HttpGet]
        public IActionResult Get()
        {
            using (var excelAppService = _excelFactory.Create(System.IO.File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
            {
                var list = new List<TestImportInput> { new TestImportInput { Code = "", Name = "", Remark = "" } };
                excelAppService.WriteListByNameManager(list, "Sheet2");
                excelAppService.Write("./files/Test_copy1.xlsx");
            }
            return Ok();
        }
    }
}