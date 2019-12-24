using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace AuditCapture
{
    public class AddAuditRecordToExcel
    {
        public string EntityName { get; set; }
        public string Action { get; set; }
        public string Operation { get; set; }
        public string CreatedOn { get; set; }
        public List<Attribute> Attributes { get; set; }

        public int ExportToExcel(_Worksheet workSheet, int row, string title)
        {
            // Populate sheet with some real data from the list of attributes            
            foreach (var attribute_ in Attributes)
            {
                workSheet.Cells[row, "A"] = title;
                workSheet.Cells[row, "B"] = EntityName;
                workSheet.Cells[row, "C"] = Action;
                workSheet.Cells[row, "D"] = Operation;
                workSheet.Cells[row, "E"] = CreatedOn;
                workSheet.Cells[row, "F"] = attribute_.AttributeName;
                workSheet.Cells[row, "G"] = attribute_.OldValue;
                workSheet.Cells[row, "H"] = attribute_.NewValue;
                row++;
            }

            return row;
        }
    }

    public class Attribute
    {
        public string AttributeName { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
}
