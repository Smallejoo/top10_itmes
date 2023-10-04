using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
class Mozar // מוצר 
{
    public string name { get; set; }
    public int count { set; get; }

}
class Program
{

    static void Main(string[] args)
    {
        string path = "D:/documents/for test"; // change to ur directory 
        List<Mozar> allProd = new List<Mozar>();
        foreach (string File in Directory.GetFiles(path, "*.Xlsx")) // going threw all exel files 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;//its a must 
            using (var package = new ExcelPackage(new FileInfo(File))) //opening file 
            {
                var page = package.Workbook.Worksheets.FirstOrDefault(); 
                if (page != null)
                {
                    var temp_list = new List<Mozar>();  
                    int start = 1;                         // i asume where it starts 
                    int end = page.Dimension.Rows;
                    for (int row = start; row <= end; row++)
                    {
                        string name_of_thing = page.Cells[row, 1].Value?.ToString();  // asuming that the name of the item is on colm 1 
                        var is_it_in_list = temp_list.FirstOrDefault(mozar => mozar.name == name_of_thing);

                        if (is_it_in_list != null) // if its already existing 
                        {
                            is_it_in_list.count++;

                        }
                        else
                        {
                            temp_list.Add(new Mozar { name = name_of_thing, count = 1 }); // if its the first time we find this item 
                        }
                    }
                    allProd.AddRange(temp_list); // merge two lists togther 
                }


            }

        }
        allProd = allProd.OrderByDescending(mozar => mozar.count).ToList(); // sort by count 
        List<Mozar> top10 = allProd.Take(10).ToList(); //take top 10 

        Console.WriteLine("ה 10 מוצרים שלנו "); // print 
        {
            for(int i=0;i<10;i++)
            {
                Console.WriteLine($"{allProd[i].name}:{allProd[i].count}");
            }
        }
    }

}