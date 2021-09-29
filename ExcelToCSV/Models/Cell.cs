using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCSV.Models
{
    internal class Cell
    {
        public int RowIndex { get; }
        public int ColumnIndex { get; }
        public object Value { get; }

        public Cell(int row, int column, object value)
        {
            RowIndex = row;
            ColumnIndex = column;
            Value = value;
        }

        public string AsString()
        {
            return Value.ToString();
        }

        public override string ToString()
        {
            return Value.ToString();
        }
    }
}
