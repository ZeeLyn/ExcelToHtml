using ExcelToHtml.Core;

var read = new ExcelReader();
await using var stream = new FileStream("C:\\Users\\ww_19\\Desktop\\UNI路书项目U币预算详情1111-奖池设定-1123更新.xlsx", FileMode.Open,
    FileAccess.ReadWrite);
var result = await read.Read2(stream);
//Console.WriteLine(html);
var html = $"<style>{File.ReadAllText("./style.css")}{result.Style}</style>"+result.Sheets.First().Html;
File.WriteAllText("D:\\test.html", html);