// See https://aka.ms/new-console-template for more information

using  Spire.Xls;
using System;
using System.IO;
Console.WriteLine("Enter pat to XLS file: ");
string pat = Console.ReadLine();
Console.WriteLine("Enter pat to export xlsx file: ");
string export = Console.ReadLine();
// String fromdirectory = @"C:\Users\ZENBOOK\RiderProjects\testing\testing\data\Manohara Bridge-20250212T124052Z-001\Manohara Bridge";
// String todirectory = "C:\\Users\\ZENBOOK\\RiderProjects\\testing\\testing\\data\\output\\manahora bridge\\";

DirectoryInfo di = new DirectoryInfo(pat);
FileInfo[] Files = di.GetFiles();

Workbook wb = new Workbook();
foreach (FileInfo file in Files)
{
    try
    {
        string filename = file.Name;
        string fin = filename.Replace("xls", "xlsx");
        wb.LoadFromFile(file.ToString());
        wb.SaveToFile(export + "" + fin);
        Console.WriteLine("saved at: " + export + "" + fin);
    }
    catch
    {
        throw new IOException();
    }

}


