

using OEM.Npoi;
using OEM.Npoi.ExcelOperatorImpl;
using Test.OEM.Npoi.TestClass;

using var excelAppService = new NpoiExcelAppService(File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite));

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

Console.ReadKey();
