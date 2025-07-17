using  Spire.Xls;
using System;
using System.IO;
Console.WriteLine("Enter path to XLS file: ");
string pat = Console.ReadLine();
Console.WriteLine("Enter base path to export xlsx file: ");
string exportBase = Console.ReadLine();



if (!Directory.Exists(exportBase))
{
    Directory.CreateDirectory(exportBase);
    Console.WriteLine("Created base export directory: " + exportBase);
}


DirectoryInfo di = new DirectoryInfo(pat);
DirectoryInfo[] directory = di.GetDirectories();
Workbook wb = new Workbook();

foreach (DirectoryInfo dri in directory)
{
    DirectoryInfo secondir = new DirectoryInfo(dri.FullName);
    
    string name = dri.FullName;
    FileInfo[] Files = secondir.GetFiles();
    string lastFolderName = Path.GetFileName(name.TrimEnd(Path.DirectorySeparatorChar));
    string exportPath = Path.Combine(exportBase, lastFolderName);

    if (!Directory.Exists(exportPath))
    { 
        try
        {
            Directory.CreateDirectory(exportPath);
            Console.WriteLine("Created export folder: " + exportPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to create directory: " + exportPath);
            Console.WriteLine("Error: " + ex.Message);
            continue; // Skip this one
        }
    }

    foreach (FileInfo file in Files)
    {
        try
        {
            string filename = file.Name;
            string fin = filename.Replace("xls", "xlsx");
            wb.LoadFromFile(file.ToString());
            wb.SaveToFile(exportPath + "\\" + fin);
            Console.WriteLine("saved at: " + exportPath + "" + fin);

        }
        catch
        {
            throw new IOException();
        }

    }
}




