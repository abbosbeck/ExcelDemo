using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;
class Program
{
    static async Task Main(string[] args)
    {

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var file = new FileInfo(fileName: @"C:\Users\Abbosjon\Desktop\Sample.xlsx");

        var people = GetSetupData();
        
        await SaveExcelFile(people, file);

        List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

        foreach (var p in peopleFromExcel)
        {
            Console.WriteLine(value: $"{p.Id} {p.FirstName} {p.LastName}");
        }
    }

    private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
    {
        List<PersonModel> output = new();

        using var package = new ExcelPackage(file);

        await package.LoadAsync(file);

        var ws = package.Workbook.Worksheets[PositionID: 0];

        int row = 3;
        int col = 1;

        while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
        {
            PersonModel p = new();
            p.Id = int.Parse(ws.Cells[row, col].Value.ToString());
            p.FirstName = ws.Cells[row, col + 1].Value.ToString();
            p.LastName = ws.Cells[row, col + 1].Value.ToString();
            output.Add(p);
            row++;
        }
        return output;
    }

    private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
    {
        DeleteIfExists(file);

        var package = new ExcelPackage(file);
        var ws = package.Workbook.Worksheets.Add(Name: "MainReport");
        var range = ws.Cells[Address: "A2"].LoadFromCollection(people, PrintHeaders: true);
        range.AutoFitColumns();
        
        Console.Write("Write title: ");
        string title = Console.ReadLine();
        
        //Formats the header 
        ws.Cells[Address: "A1"].Value = title;
        ws.Cells[Address: "A1:C1"].Merge = true;
        ws.Column(col: 1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        ws.Row(row: 1).Style.Font.Size = 24;

        ws.Row(row: 2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        ws.Row(row: 2).Style.Font.Bold = true;

        for (int i = 1; i <= 3; i++)
        {
            ws.Column(col: i).Width = 20;
            ws.Column(col: i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        }


        await package.SaveAsync();

    }

    private static void DeleteIfExists(FileInfo file)
    {
        if (file.Exists)
        {
            file.Delete();
        }
    }

    private static List<PersonModel> GetSetupData()
    {
        List<PersonModel> output = new()
        {
            new() {Id = 1, FirstName = "Ali", LastName = "Valiyev"},
            new() {Id = 2, FirstName = "John", LastName = "Johnov" },
            new() {Id = 3, FirstName = "Andery", LastName = "Hasanov"},
        };

        return output;
    }
}
