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
    Console.WriteLine($"编码：【{item.Code}】，名称：【{item.Name}】，备注：【{item.Remark}】");
}

var user = excelAppService.ReadByNameManager<TestUserOutput>("Sheet2");
Console.WriteLine($"姓名：{user.TrueName}，年龄：{user.Age}，性别：{user.Gender}");

user.TrueName = "王五";
excelAppService.WriteByNameManager(user, "Sheet2");

user = excelAppService.ReadByNameManager<TestUserOutput>("Sheet2");
Console.WriteLine($"姓名：{user.TrueName}，年龄：{user.Age}，性别：{user.Gender}");

excelAppService.Write("./files/Test_copy.xlsx");


app.Run();
