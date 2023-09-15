namespace Converter;

public class Converter
{
    public string Path { get; private set; }
    public string SheetName { get; set; }
    private List<string> OutputList { get; set; }

    public Converter(string path, string sheetName)
    {
        this.Path = path;
        this.SheetName = sheetName;
        OutputList = new List<string>();
    }
    
    
}