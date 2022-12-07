using ExcelToHtml.Core;


await using var stream = new FileStream("C:\\Users\\ww_19\\Desktop\\UNI路书项目U币预算详情1111-奖池设定-1123更新.xlsx", FileMode.Open,
    FileAccess.ReadWrite);
var result = await ExcelToHtmlConverter.GenerateUIAsync(stream);
//var html = $"<style>{result.GlobalStyle}</style>"+result.Sheets.First().Html;
File.WriteAllText("D:\\test.html", result.ToString());