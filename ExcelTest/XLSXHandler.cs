using ClosedXML;
using ClosedXML.Excel;
using System.Collections.Generic;
using ExcelTest;


namespace ExcelTest;

public class XLSXHandler : ExcelHandler
{
    public XLSXHandler()
    {
        
    }
    
    // important column 1 / 2
    public List<MyExcelData> GetXlsxToList(string filePath, int startRow, int impCol1, int impCol2)
    {
        List<MyExcelData> temp = new List<MyExcelData>();
       
        string date;
        string user;
        string rz;
        string worksheetName;
        ClosedXML.Excel.IXLWorksheet ws;
        ClosedXML.Excel.IXLWorkbook wb;
        ClosedXML.Excel.IXLRow row;
        ClosedXML.Excel.IXLCell currentCell;

        int i = startRow;
        int cellNum = 1;
        wb = new XLWorkbook(filePath); // create new Workbook with filepath


        if (Path.GetFileName(filePath).StartsWith("TAG"))
        {
            rz = "tag";
            worksheetName = "Sheet1";
            cellNum = 19;

            wb = new XLWorkbook(filePath); // create new Workbook with filepath
            ws = wb.Worksheet(worksheetName); // select worksheet
            row = ws.Row(startRow); // select row to start at
            currentCell = row.Cell(10); // select cell


            while (true)
            {
                if (ws.Row(i).Cell(10).IsEmpty() != true) break;

                if (ws.Row(i).IsEmpty() != true)
                {
                    date = (ws.Row(i).Cell(impCol1).GetValue<string>());
                    date = date.Substring(0, 10);
                    user = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol2).GetValue<string>());


                    if (i != 1 && !ws.Row(i).Cell(1).IsEmpty())
                    {
                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    }


                    currentCell = ws.Row(i).Cell(cellNum);
                }
                
                i++;
            }
        }
        else if (Path.GetFileName(filePath).StartsWith("QRZ"))
        {
            rz = "qrz";
            worksheetName = "Page 1";
            cellNum = 1;

            wb = new XLWorkbook(filePath); // create new Workbook with filepath
            ws = wb.Worksheet(worksheetName); // select worksheet
            row = ws.Row(startRow); // select row to start at
            currentCell = row.Cell(1); // select cell

            currentCell = row.Cell(10); // select cell

            while (StillValidQRZ(currentCell))
            {

                if (i != 1 && !ws.Row(i).Cell(1).IsEmpty())
                {
                    temp.Add(
                        new MyExcelData
                        {
                            Date = (ws.Row(i).Cell(impCol1).GetValue<string>()).Substring(0, 10),
                            User = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol2).GetValue<string>()),
                            RZ = rz,
                        });
                }

                i++;

                currentCell = ws.Row(i).Cell(cellNum);
            }
        }
        else
        {
            throw new Exception("wrong file name");
        }
        
        return temp;
    } // TAG/QRZ

    public List<MyExcelData> GetXlsxToList(string filePath, int startRow, int impCol1, int impCol2, int impCol3)
    {
        List<MyExcelData> temp = new List<MyExcelData>();

        string worksheetName;
        string date;
        string user;
        string rz;

        ClosedXML.Excel.IXLWorksheet ws;
        ClosedXML.Excel.IXLWorkbook wb;
        ClosedXML.Excel.IXLRow row;
        ClosedXML.Excel.IXLCell currentCell;

        if (filePath != null || filePath != "")
        {
            int i = startRow;
            int cellNum = 1;
            wb = new XLWorkbook(filePath); // create new Workbook with filepath

            if (Path.GetFileName(filePath).StartsWith("Anmeldungen_Sharepoint"))
            {
                rz = "qrz";
                worksheetName = "query (1)";
                cellNum = 1;

                wb = new XLWorkbook(filePath); // create new Workbook with filepath
                ws = wb.Worksheet(worksheetName); // select worksheet
                row = ws.Row(startRow); // select row to start at
                currentCell = row.Cell(10); // select cell

                while (StillValidQRZ(currentCell))
                {

                    if (i != 1 && !ws.Row(i).Cell(1).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(1, 10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol2).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    }


                    i++;

                    currentCell = ws.Row(i).Cell(cellNum);
                }
            }
            else
            {
                throw new Exception("wrong file name");
            }

            return temp;
        }

        return null;
    }  // TAG/QRZ

    public List<MyExcelData> GetXlsxToList(string filePath, string escapeReference, int startRow, int impCol1, int impCol2, int impCol3) // Sharepoint
    {
        List<MyExcelData> temp = new List<MyExcelData>();

        string worksheetName;
        string date;
        string user;
        string rz;

        ClosedXML.Excel.IXLWorksheet ws;
        ClosedXML.Excel.IXLWorkbook wb;
        ClosedXML.Excel.IXLRow row;
        ClosedXML.Excel.IXLCell currentCell;

        if (filePath != null || filePath != "")
        {
            int i = startRow;
            int cellNum = 1;
            wb = new XLWorkbook(filePath); // create new Workbook with filepath

            if (Path.GetFileName(filePath).StartsWith("Anmeldungen_Sharepoint"))
            {
                worksheetName = "query (1)";
                cellNum = 1;

                wb = new XLWorkbook(filePath); // create new Workbook with filepath
                ws = wb.Worksheet(worksheetName); // select worksheet
                row = ws.Row(startRow); // select row to start at
                currentCell = row.Cell(10); // select cell

                while (true)
                {
                    if (i != 1 && !ws.Row(i).Cell(1).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(0,10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol2).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());

                        if (DateTime.Compare(DateTime.Parse(date), DateTime.Parse(escapeReference)) < 0) break; // excapesequenz

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    } // Hauptbesucher
                    if (i != 1 && !ws.Row(i).Cell(8).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(0, 10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(8).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());

                        if (user.Contains(' '))
                        {
                            string[] splitedUser = user.Split(' ');
                            user = splitedUser[0] + ' ' + splitedUser[1];
                        }

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    } // 1. Besucher
                    if (i != 1 && !ws.Row(i).Cell(9).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(0, 10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(9).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());
                        
                        if (user.Contains(' '))
                        {
                            string[] splitedUser = user.Split(' ');
                            user = splitedUser[0] + ' ' + splitedUser[1];
                        }

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    } // 2. Besucher
                    if (i != 1 && !ws.Row(i).Cell(10).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(0, 10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(10).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());

                        if (user.Contains(' '))
                        {
                            string[] splitedUser = user.Split(' ');
                            user = splitedUser[0] + ' ' + splitedUser[1];
                        }

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    } // 3. Besucher
                    if (i != 1 && !ws.Row(i).Cell(11).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(0, 10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(11).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());

                        if (user.Contains(' '))
                        {
                            string[] splitedUser = user.Split(' ');
                            user = splitedUser[0] + ' ' + splitedUser[1];
                        }

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    } // 4. Besucher
                    if (i != 1 && !ws.Row(i).Cell(12).IsEmpty())
                    {
                        date = (GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol1).GetValue<string>())).Substring(0, 10);
                        user = GetDataIntoCorrectFormat(ws.Row(i).Cell(12).GetValue<string>());
                        rz = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol3).GetValue<string>());

                        if (user.Contains(' '))
                        {
                            string[] splitedUser = user.Split(' ');
                            user = splitedUser[0] + ' ' + splitedUser[1];
                        }

                        temp.Add(
                            new MyExcelData
                            {
                                Date = date,
                                User = user,
                                RZ = rz,
                            });
                    } // 5. Besucher

                    i++;

                    currentCell = ws.Row(i).Cell(cellNum);
                }
            }
            else
            {
                throw new Exception("wrong file name");
            }

            return temp;
        }

        return null;
    }

    public string CreateWriteCheckFile(List<MyExcelData> collectedDataList, List<MyExcelData> sharepointList, string pathToSaveTo)
    {
        IXLWorkbook wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
       
        int i = 1;

        string td = System.DateTime.Now.ToString();
        td = td.Substring(6,4) + td.Substring(3,2) + td.Substring(1,2) + '_' + td.Substring(11,2) + '_' + td.Substring(14, 2) + '_' + td.Substring(17, 2) + '_';

        // the range for which you want to add a table style
        var range = ws.Range(1, 1, collectedDataList.Count() + 1, 6);

        // create the actual table
        var table = range.CreateTable();


        ws.Column("A").Width = 15;
        ws.Column("A").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(i).Cell(1).Value = "Datum";

        ws.Column("B").Width = 25;
        ws.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        ws.Row(i).Cell(2).Value = "Mitarbeiter";

        ws.Column("C").Width = 8;
        ws.Column("C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(i).Cell(3).Value = "SharePoint";

        ws.Column("D").Width = 22;
        ws.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(i).Cell(4).Value = "Status";

        ws.Column("E").Width = 18;
        ws.Column("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(i).Cell(5).Value = "Kontrolle";

        ws.Column("F").Width = 35;
        ws.Row(i).Cell(6).Value = "Kommentar";

        // apply style
        table.Theme = XLTableTheme.TableStyleLight12;

        i++;

        foreach (var item in collectedDataList)
        {
            ws.Row(i).Cell(1).Value = item.Date;
            ws.Row(i).Cell(2).Value = item.User;
            ws.Row(i).Cell(3).Value = item.RZ.ToUpper();

            i++;
        }

        i = 1;

        foreach (var item in sharepointList)
        {
            ws.Row(i).Cell(15).Value = item.User;
            ws.Row(i).Cell(16).Value = item.Date;
            ws.Row(i).Cell(17).Value = item.RZ;

            i++;
        }

        pathToSaveTo = pathToSaveTo + "\\TempFileForChecking_" + td + ".xlsx";

        wb.SaveAs(pathToSaveTo);

        return pathToSaveTo;
    }

    
    
    public bool StillValidQRZ(ClosedXML.Excel.IXLCell currentCell)
    {
        if (currentCell.GetValue<string>().Contains("Anzahl Datensätze:"))
        {
            return false;
        }

        return true;
    }

    public bool StillValidTAG(ClosedXML.Excel.IXLCell currentCell)
    {
        //if(currentCell.Style.Fill.BackgroundColor.ToString() == "Color Index: 64")
        //{
        //    return false;
        //}

        if(currentCell.IsEmpty() != true)
        {
            return false;
        }

        return true;
    }

    public string GetEarlyestDate(List<MyExcelData> listToFilter)
    {
        DateTime earlyestDate = DateTime.Now;
        int i = 1;

        listToFilter.Sort((x, y) => y.Date.CompareTo(x.Date));

        earlyestDate = DateTime.Parse(listToFilter[listToFilter.Count() - 1].Date);

        return earlyestDate.ToString().Substring(0,10);
    }

    public string GetDataIntoCorrectFormat(string data)
    {
        if (data.Contains("(GW)")) data = data.Substring(1, data.Length - 5);
        if (data.StartsWith(" ")) data = data.TrimStart(' ');
        if (data.Contains(',')) data = data.Replace(",",String.Empty);

        char[] chars = data.ToCharArray();
        data = null;

        foreach (char item in chars)
        {
            if (item == 'ö')
                data = data + "oe";
            else if (item == 'Ö')
                data = data + "Oe";
            else if (item == 'ü')
                data = data + "ue";
            else if (item == 'Ü')
                data = data + "Ue";
            else if (item == 'ä')
                data = data + "ae";
            else if (item == 'Ä')
                data = data + "Ae";
            else if (item == ',');
            else data = data + item;
        }

        data = data.ToLower();

        return data;
    }

    public bool CompareItems(MyExcelData item1, MyExcelData item2)
    {
        if (item1.Date == item2.Date && item1.User == item2.User && item1.RZ == item2.RZ) return true;
        else return false;
    }
}