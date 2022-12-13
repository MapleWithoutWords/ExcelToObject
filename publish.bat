dotnet build -c Release
dotnet pack ./src/ExcelToObject.Core/ExcelToObject.Core.csproj -c Release -o .\publish\ExcelToObject.Core
dotnet pack ./src/ExcelToObject.Npoi/ExcelToObject.Npoi.csproj -c Release -o .\publish\ExcelToObject.Npoi
