using OEM.Core;
using OEM.Core.ExcelOperator;
using Test.OEM.NpoiWebApi.TestClass;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
builder.Services.AddExcelToObjectNpoiService();

var app = builder.Build();

// Configure the HTTP request pipeline.

app.UseAuthorization();

app.MapControllers();

var excelFactory = app.Services.GetService<IExcelFactory>();
using var excelAppService = excelFactory.Create(File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite));

var list = excelAppService.ReadListByNameManager<TestImportInput>("Sheet1");

foreach (var item in list)
{
    Console.WriteLine($"���룺��{item.Code}�������ƣ���{item.Name}������ע����{item.Remark}��");
}

var user = excelAppService.ReadByNameManager<TestUserOutput>("Sheet2");
Console.WriteLine($"������{user.TrueName}�����䣺{user.Age}���Ա�{user.Gender}");

user.TrueName = "����";
excelAppService.WriteByNameManager(user, "Sheet2");

user = excelAppService.ReadByNameManager<TestUserOutput>("Sheet2");
Console.WriteLine($"������{user.TrueName}�����䣺{user.Age}���Ա�{user.Gender}");

excelAppService.Write("./files/Test_copy.xlsx");


app.Run();
