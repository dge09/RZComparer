using ClosedXML;
using ClosedXML.Excel;
using System.Collections.Generic;
using ExcelTest;


namespace ExcelTest;

public class XLSXHandler : ExcelHandler
{    
    public List<MyExcelData> GetXlsxToList(string filePath)
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

        int startRow = 17;
        int i = startRow;
        int impCol1 = 0;
        int impCol2 = 0;

        int cellNum = 1;
        wb = new XLWorkbook(filePath); // create new Workbook with filepath


        if (Path.GetFileName(filePath).StartsWith("TAG"))
        {
            rz = "tag"; // default rz
            worksheetName = "Sheet1"; // worksheet used
            cellNum = 19; // starting row

            wb = new XLWorkbook(filePath); // create new Workbook with filepath
            ws = wb.Worksheet(worksheetName); // select worksheet
            row = ws.Row(startRow); // select row to start at
            currentCell = row.Cell(10); // select cell

            // get collum number to get data
            for (int k = 1; k < 26; k++)
            {
                if (ws.Row(16).Cell(k).GetValue<string>() == "Datum") impCol1 = k;
                if (ws.Row(16).Cell(k).GetValue<string>() == "Person") impCol2 = k;
            }


            while (true)
            {
                if (ws.Row(i).Cell(10).IsEmpty() != true) break; // escape reference 

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
                } // read data, write data to list
                
                i++;
            } // read data from file into Temp list
        } // get data to list (TAG format)
        else if (Path.GetFileName(filePath).StartsWith("QRZ"))
        {
            rz = "qrz"; // default rz 
            worksheetName = "Page 1"; // worksheet used
            cellNum = 1; // starting row
            i = 2; // starting row

            wb = new XLWorkbook(filePath); // create new Workbook with filepath
            ws = wb.Worksheet(worksheetName); // select worksheet
            row = ws.Row(startRow); // select row to start at
            currentCell = row.Cell(1); // select cell

            currentCell = row.Cell(10); // select cell

            for (int k = 1; k < 26; k++)
            {
                if (ws.Row(1).Cell(k).GetValue<string>() == "Zeitpunkt") impCol1 = k;
                if (ws.Row(1).Cell(k).GetValue<string>() == "Person") impCol2 = k;
            } // get collumns with needed data

            while (StillValidQRZ(currentCell))
            { 
                if (ws.Row(i).Cell(1).GetValue<string>().Contains("Anzahl Datensätze:")) break; // escape sequenz

                if (i != 1 && !ws.Row(i).Cell(1).IsEmpty())
                {                    
                    temp.Add(
                    new MyExcelData
                    {
                        Date = (ws.Row(i).Cell(impCol1).GetValue<string>()).Substring(0, 10),
                        User = GetDataIntoCorrectFormat(ws.Row(i).Cell(impCol2).GetValue<string>()),
                        RZ = rz,
                    });
                } // read data, write data to temp list

                i++;


                currentCell = ws.Row(i).Cell(cellNum);
            }
        } // get data to list (QRZ format)
        else
        {
            throw new Exception("wrong file name");
        } // exception for not existing file
        
        return temp;
    } // TAG/QRZ

    public List<MyExcelData> GetXlsxToList(string filePath, string escapeReference) 
    {
        List<MyExcelData> temp = new List<MyExcelData>();

        string worksheetName;
        string date;
        string user;
        string rz;

        int startRow = 2;
        int impCol1 = 0;
        int impCol2 = 0;
        int impCol3 = 0;

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
                worksheetName = "query (1)"; // used worksheet
                cellNum = 1; // starting row

                wb = new XLWorkbook(filePath); // create new Workbook with filepath
                ws = wb.Worksheet(worksheetName); // select worksheet
                row = ws.Row(startRow); // select row to start at
                currentCell = row.Cell(10); // select cell

                for (int k = 1; k < 26; k++)
                {
                    if (ws.Row(1).Cell(k).GetValue<string>() == "Datum") impCol1 = k;
                    if (ws.Row(1).Cell(k).GetValue<string>() == "Name des Besuchers") impCol2 = k;
                    if (ws.Row(1).Cell(k).GetValue<string>() == "Rechenzentrum") impCol3 = k;
                } // get collumn number for needed data

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
                    } // main visitor

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
                    } // 1.  visitor
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
                    } // 2.  visitor
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
                    } // 3. visitor
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
                    } // 4. visitor
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
                    } // 5. visitor

                    i++;

                    currentCell = ws.Row(i).Cell(cellNum);
                }
            } // read data, write data to sharepoint list
            else
            {
                throw new Exception("wrong file name");
            } // exception for not existing file

            return temp;
        } // read data, write data to sharepoint list

        return null;
    } // Sharepoint

    public string CreateWriteCheckFile(List<MyExcelData> collectedDataList, List<MyExcelData> sharepointList, string pathToSaveTo)
    {
        IXLWorkbook wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
       
        // create the actual table
        // the range for which you want to add a table style
        var range = ws.Range(1, 1, collectedDataList.Count() + 1, 6);

        var table = range.CreateTable();

        int i = 1;


        // create timestamp for filename
        string td = System.DateTime.Now.ToString();
        td = td.Substring(6,4) + td.Substring(3,2) + td.Substring(1,2) + '_' + td.Substring(11,2) + '_' + td.Substring(14, 2) + '_' + td.Substring(17, 2) + '_'; 

        // formating the new file =======================================================================================================================================================================
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

        
        ws.Row(2).Cell(10).Value = "OK";
        ws.Row(2).Cell(10).RichText.SetFontColor(XLColor.White);
        ws.Row(3).Cell(10).Value = "Fehlend";
        ws.Row(3).Cell(10).RichText.SetFontColor(XLColor.White);


        ws.Row(i).Cell(5).Value = "Kontrolle";
        ws.Column("E").Width = 18;
        //Applying range validation in sheet 1 where drop down list is to be shown
        ws.Range("E2", "E" + (collectedDataList.Count + 1).ToString()).SetDataValidation().List(ws.Range("J2:J3"), true);
        ws.Range("E2", "E" + (collectedDataList.Count + 1)).Value = "Fehlend";
        ws.Column("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("E2", "E" + (collectedDataList.Count + 1).ToString()).AddConditionalFormat().WhenEquals("OK").Fill.SetBackgroundColor(XLColor.FromHtml("#02FE08"));
        ws.Range("E2", "E" + (collectedDataList.Count + 1).ToString()).AddConditionalFormat().WhenNotEquals("OK").Fill.SetBackgroundColor(XLColor.FromHtml("#FF1919"));

        ws.Column("F").Width = 35;
        ws.Column("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(i).Cell(6).Value = "Kommentar";

        // apply style
        table.Theme = XLTableTheme.TableStyleLight12;
        // ==============================================================================================================================================================================================

        i++;

        foreach (var item in collectedDataList)
        {
            ws.Row(i).Cell(1).Value = item.Date;
            ws.Row(i).Cell(2).Value = item.User;
            ws.Row(i).Cell(3).Value = item.RZ.ToUpper();

            i++;
        } // write data to excel table

        table.Sort("Datum"); // sort table date ascending

        pathToSaveTo = pathToSaveTo + "\\TempFileForChecking_" + td + ".xlsx"; // get filename to return 

        wb.SaveAs(pathToSaveTo); // save new excel table 

        return pathToSaveTo; // return new file name
    } // create new excel file with differences

    
    
    public bool StillValidQRZ(ClosedXML.Excel.IXLCell currentCell)
    {
        if (currentCell.GetValue<string>().Contains("Anzahl Datensätze:"))
        {
            return false;
        }

        return true;
    } // check if file end is reached (QRZ)

    public bool StillValidTAG(ClosedXML.Excel.IXLCell currentCell)
    {
        if(currentCell.IsEmpty() != true)
        {
            return false;
        }

        return true;
    } // check if file end is reached (TAG)

    public string GetEarlyestDate(List<MyExcelData> listToFilter)
    {
        DateTime earlyestDate = DateTime.Now;
        int i = 1;

        earlyestDate = DateTime.Parse(listToFilter[0].Date); // list is already sorted, we can just get first entry

        return earlyestDate.ToString().Substring(0,10);
    } // get earlyest date for escape reference in sharepoint file

    public string GetDataIntoCorrectFormat(string data)
    {
        if (data == null || data == "") return "Error";

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
        } // rewritting ä,ö,ü to ae, oe, ue

        data = data.ToLower(); // make non casesensetive

        return data;
    } // reformat read data to compare it easily 
}