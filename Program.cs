using Microsoft.Office.Interop.Excel;
//Console.WriteLine("Hello, World!");
/// <summary>
/// Using Microsoft.Office.Interop to convert XLS to XLSX format, to work with EPPlus library
/// </summary>
/// <param name="filesFolder"></param>

internal class Program
{
    static void Main(string[] args)
    {
      //pasar el string del path con el nombre de todos los archivos que sean xls 
        var CurrentDirectory = Directory.GetCurrentDirectory();
      
      //algo como  FOReach file in CurrentDirectory  realizar ese metodo
        ConvertXLS_XLSX(CurrentDirectory);
    }

    private static string ConvertXLS_XLSX(FileInfo file)
    {
        var app = new Application();
        var xlsFile = file.FullName;
        var wb = app.Workbooks.Open(xlsFile);
        var xlsxFile = xlsFile + "x";
        wb.SaveAs(Filename: xlsxFile, FileFormat: XlFileFormat.xlOpenXMLWorkbook);
        wb.Close();
        app.Quit();
        return xlsxFile;
    }
}