# ExcelToObject

#### Description
ExcelToObject，Make operating excel as simple as operating objects.

#### Installation

```shell
dotnet add package ExcelToObject.Npoi --version 1.0.0
```



#### Instructions

1.  The Web program adds the following code to the Program or Startup

```c#
service.AddExcelToObjectNpoiService();
```

2. Console Application

```
IExcelFactory _excelFactory=new NpoiExcelFactory();
```

3. Read a sheet in excel as a list

```c#
using (var excelAppService = _excelFactory.Create(System.IO.File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
{
    var list = excelAppService.ReadListByNameManager<TestImportInput>("Sheet1");
    foreach (var item in list)
    {
        Console.WriteLine($"编码：【{item.Code}】，名称：【{item.Name}】，备注：【{item.Remark}】");
    }
}
```

4. Read a sheet in Excel as a single object

```c#
using (var excelAppService = _excelFactory.Create(System.IO.File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
{
    var user = excelAppService.ReadByNameManager<TestUserOutput>("Sheet2");
    Console.WriteLine($"姓名：{user.TrueName}，年龄：{user.Age}，性别：{user.Gender}");
}
```

5. Write a list object to the Excel Sheet page

```c#
using (var excelAppService = _excelFactory.Create(System.IO.File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
{
    excelAppService.WriteListByNameManager(new List<TestImportInput> { new TestImportInput { Code = "", Name = "", Remark = "" } }, "Sheet2");
    excelAppService.Write("./files/Test_copy1.xlsx");
}
```

6. Write an object to the excel sheet page

```c#
using (var excelAppService = _excelFactory.Create(System.IO.File.Open("./files/Test.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite)))
{
    excelAppService.WriteByNameManager(new TestUserOutput { Age=18, Gender="男", TrueName="赵六" }, "Sheet2");
    excelAppService.Write("./files/Test_copy1.xlsx");
}
```



#### Principle

![image-20221207000351142](.\doc\images\image-20221207000351142.png)

1. Match with the attribute name of the entity through the name manager of Excel
2. Get entity attributes through the passed in entity. Find out the name manager in sheetName and the coordinates of each name. Read.


#### Contribution

1.  Fork the repository
2.  Create Feat_xxx branch
3.  Commit your code
4.  Create Pull Request
