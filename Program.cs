using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

public class RecursiveFileProcessor
{

    public static async Task Main(string[] args)
    {
       
       foreach(string path in args)
       {
            if (Directory.Exists(path))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var file = new FileInfo(@"C:\Users\potluris\FileNames.xlsx");

                var filenames = ReadFileNamesFromDirectory(path);

                await SaveToExcelFile(filenames, file);
            }
            else
            {
                Console.WriteLine("{0} is not a directory.", path);
            }

       }
    }

    private static async Task SaveToExcelFile(List<string> filenames, FileInfo file)
    {
        DeleteIfFileExists(file);

        using (var package = new ExcelPackage(file))
        {
            var ws = package.Workbook.Worksheets.Add("Bride Album");

            var range = ws.Cells["A2"].LoadFromCollection(filenames);
            range.AutoFitColumns();

            //format the header
            ws.Cells["A1"].Value = "Bride Album";
            ws.Column(col: 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(row: 1).Style.Font.Color.SetColor(Color.Blue);

            await package.SaveAsync();
        };
        
    }

    private static void DeleteIfFileExists(FileInfo file)
    {
        if (file.Exists)
        {
            File.Delete(file.FullName);
        }
    }

    private static List<string> ReadFileNamesFromDirectory(string path)
    {
        List<string> filenames = Directory.GetFiles(path).Select(f => Path.GetFileName(f)).ToList();
       
        foreach(string filename in filenames)
        {
            Console.WriteLine("Processed file {0}", filename);
        }

        return filenames;
    }
}
