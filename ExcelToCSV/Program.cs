using ExcelToCSV.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelToCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string directory = Directory.GetCurrentDirectory();

                var filesInDirectory = from file in Directory.GetFiles(directory)
                            where Path.GetExtension(file) == ".xlsx" || Path.GetExtension(file) == ".xls"
                            select file;

                if (filesInDirectory.Count() == 0) throw new Exception("В текущей папке отсутствуют файлы с форматом \".xlsx\", \".xls\"");

                Console.WriteLine("*Введите номера файлов, которые хотите конвертировать в CSV, через пробел и нажмите Enter\n" +
                    "*Или нажмите Enter, чтобы конвертировать все файлы в папке");

                foreach (string file in filesInDirectory)
                {
                    Console.WriteLine(filesInDirectory.ToList().IndexOf(file).ToString() + " " + Path.GetFileNameWithoutExtension(file));
                }

                string selectedFiles =  Console.ReadLine();

                List<string> files;
                try
                {
                    if (selectedFiles.Length != 0)
                    {
                        files = new List<string>();
                        foreach (string index in selectedFiles.Trim().Split(' '))
                        {
                            int i = Int32.Parse(index);
                            files.Add(filesInDirectory.ToList()[i]);
                        }
                    }
                    else throw new Exception();
                }
                catch
                {
                    files = filesInDirectory.ToList();
                };

                foreach (string file in files)
                {
                    Excel excel = new Excel(file);
                    excel.SaveAs();
                    Console.WriteLine(excel.Name);
                }
                Console.WriteLine("Готово");
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
                return;
            }
        }
    }
}
