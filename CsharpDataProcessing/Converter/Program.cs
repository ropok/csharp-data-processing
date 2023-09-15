using OfficeOpenXml;

namespace Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.WriteLine("Hello, World!");

            var file = new FileInfo(@"");
            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets["BHP RT"];
            Console.WriteLine(ws.Cells[1, 1].Value.ToString());

            var texts = new List<string>();
            var textList = new List<string>();

            for (int i = ws.Dimension.Start.Row; i <= ws.Dimension.End.Row; i++)
            {
                if (ws.Cells[i, 2].Value != null)
                {
                    textList.Add(string.Join(',', texts));
                    texts.Clear();
                }

                for (int j = ws.Dimension.Start.Column; j <= ws.Dimension.End.Column; j++)
                {
                    if (ws.Cells[i, j].Value != null)
                    {
                        var value = ws.Cells[i, j].Value.ToString();
                        if (j >= 3)
                        {
                            texts.Add(value);
                        }
                    }
                }
            }

            foreach (var text in textList)
            {
                Console.WriteLine(text);
            }

            using StreamWriter writer = new StreamWriter(@"");
            foreach (var text in textList)
            {
                writer.WriteLine(text);
            }
        }
    }
}