using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelToCSV.Models
{
    internal class Excel
    {
        public string Path { get; }
        public string Name
        {
            get => System.IO.Path.GetFileNameWithoutExtension(Path);
        }
        public string Extension
        {
            get => System.IO.Path.GetExtension(Path);
        }
        public string Directory
        {
            get => System.IO.Path.GetDirectoryName(Path);
        }
        public List<Sheet> Sheets { get; }

        public Excel(string path)
        {
            Path = path;
            Sheets = GetSheets();
        }

        List<Sheet> GetSheets()
        {
            List<Sheet> sheets = new List<Sheet>();
            using (var stream = File.Open(Path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    foreach (DataTable table in result.Tables)
                    {
                        Sheet sheet = new Sheet(table);
                        sheets.Add(sheet);
                    }

                }
            }
            return sheets;
        }

        public void SaveAs()
        {
            string folder = "CSV";
            System.IO.Directory.CreateDirectory(folder);
            string path = System.IO.Path.ChangeExtension(System.IO.Path.Combine(Directory,folder,Name),".csv");
            using (StreamWriter sw = new StreamWriter(path, false, System.Text.Encoding.Default))
            {
                foreach (var row in Sheets[0].Rows)
                {
                    sw.WriteLine(String.Join(";",row));
                }
            }
        }
    }
}
