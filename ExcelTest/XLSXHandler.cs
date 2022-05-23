using ClosedXML;
using ClosedXML.Excel;
using System.Collections.Generic;
using ExcelTest;


namespace ExcelTest;

public class XLSXHandler : ExcelHandler
{
    public string Read(string filePath)
    {
        // 2. create new Workbook with filepath
        var wb = new XLWorkbook(filePath);
        // 3. select worksheet 
        var ws = wb.Worksheet("Tabelle1");
        // 4. select row (start at 1)
        var row = ws.Row(2);
        // 5. check if row is empty
        bool isEmpty = row.IsEmpty();
        // 6. access row
        var cell = row.Cell(3);
        // 7. get value
        string value = cell.GetValue<string>();


        return value;
    }

    public List<string> GetDataToList(string filePath)
    {
        // 2. create new Workbook with filepath
        var wb = new XLWorkbook(filePath);
        // 3. select worksheet 
        var ws = wb.Worksheet("Tabelle1");
        // 4. select row (start at 1)
        var row = ws.Row(2);

        // 5. check if row is empty
        bool isEmpty = row.IsEmpty();
        // 6. access row

        // 7. get value
        string value = row.Cell(1).GetValue<string>();

        // create list
        


        int i = 1;

        while (!isEmpty)
        {
            if (i != 1)
            {
                var t1 = new MyExcelData
                {
                    Date = row.Cell(1).GetValue<string>(),
                    User = row.Cell(2).GetValue<string>(),
                    UserId = row.Cell(3).GetValue<string>(),
                    Department = row.Cell(4).GetValue<string>(),
                    DepartmentId = row.Cell(5).GetValue<string>(),
                };
            }

            i++;
            isEmpty = ws.Row(2).IsEmpty();
        }
        
        return value;
    }

}
