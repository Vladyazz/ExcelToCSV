using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelToCSV.Models
{
    internal class Sheet
    {
        public string Name { get; }
        public List<Cell> Cells { get; }
        public List<List<Cell>> Rows
        {
            get
            {
                List<List<Cell>> rows = new List<List<Cell>>();
                foreach (var row in Cells.GroupBy(q => q.RowIndex).OrderBy(q => q.Key))
                {
                    List<Cell> list = new List<Cell>();
                    foreach (var cell in row.OrderBy(q => q.ColumnIndex))
                    {
                        list.Add(cell);
                    }
                    rows.Add(list);
                }
                return rows;
            }
        }

        public Sheet(DataTable table)
        {
            Name = table.TableName;
            Cells = GetCells(table);
        }

        List<Cell> GetCells(DataTable table)
        {
            List<Cell> cells = new List<Cell>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    Cell cell = new Cell(i, j, table.Rows[i][j]);
                    cells.Add(cell);
                }
            }
            return cells;
        }
    }
}
