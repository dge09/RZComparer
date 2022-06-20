using ExcelTest;
using Spectre.Console;

XLSXHandler xlsxHandler = new XLSXHandler();
CSVHandler csvHandler = new CSVHandler();

Console.WriteLine("Enter path with files to check and press \"Enter\"");
string generalPath = Console.ReadLine();

Console.WriteLine("");


List<string> fileList = new List<string>(); // List of all files in directory
List<MyExcelData> summaryList = new List<MyExcelData>(); // List with data from all temp lists
List<MyExcelData> sharePointList = new List<MyExcelData>(); // list with sharepoint data
List<MyExcelData> differencesList = new List<MyExcelData>(); // List with only the differences from sharepoint and other files

List<MyExcelData> tempQRZ = new List<MyExcelData>(); // List used to store temporary QRZ Data
List<MyExcelData> tempTAG = new List<MyExcelData>(); // List used to store temporary TAG Data
List<MyExcelData> tempZAG = new List<MyExcelData>(); // List used to store temporary ZAG Data

List<int> itemsToDelete = new List<int>(); // List with items to ignore in differences list

fileList = Directory.GetFiles(generalPath).ToList<string>(); // get all files in directory

string sharePointFile = null;
string escapeReference; // last date to get in sharepoint file


foreach (var file in fileList) // get data from files
{
    if (file.Contains("QRZ"))
    {
        tempQRZ = xlsxHandler.GetXlsxToList(file);

        foreach (var item in tempQRZ)
        {
            summaryList.Add(item);
        }
    } // get data from QRZ files and add data to summaryList
    else if (file.Contains("TAG"))
    {
        tempTAG = xlsxHandler.GetXlsxToList(file);

        foreach (var item in tempTAG)
        {
            summaryList.Add(item);
        }
    } // get data from TAG files and add data to summaryList
    else if (file.Contains("ZAG")) 
    {
        tempZAG = csvHandler.GetCsvToList(file, 9);

        foreach (var item in tempZAG)
        {
            summaryList.Add(item);
        }
    } // get data from ZAG files and add data to summaryList
    else if (file.Contains("Anmeldungen_Sharepoint"))
    {
        sharePointFile = file;
    } // get sharepoint file name
}

// make list distinct and sort Date ascending
summaryList = summaryList.DistinctBy(d => new { d.Date, d.User }).ToList();
summaryList.Sort((x, y) => DateTime.Compare(DateTime.Parse(x.Date), DateTime.Parse(y.Date)));

// get earlyest date in summary list to check how much of the sharepoint file is needed
escapeReference = xlsxHandler.GetEarlyestDate(summaryList);

// get sharepoint data
sharePointList = xlsxHandler.GetXlsxToList(sharePointFile, escapeReference);

// make list distinct and sort Date ascending
sharePointList = sharePointList.DistinctBy(d => new { d.Date, d.User }).ToList();
sharePointList.Sort((x, y) => DateTime.Compare(DateTime.Parse(x.Date), DateTime.Parse(y.Date)));


// needed to use .max function
itemsToDelete.Add(0);


for (int i = 0; i < summaryList.Count(); i++) 
{
    int j = 0;
    foreach (var shareItem in sharePointList)
    {
        if (j >= itemsToDelete.Max())
        {
            //if (DateTime.Parse(finalList[i].Date).Date > DateTime.Parse(shareItem.Date).Date) break;

            if (summaryList[i].Date == shareItem.Date && summaryList[i].User == shareItem.User && summaryList[i].RZ == shareItem.RZ)
            {
                itemsToDelete.Add(i);
            }
            else if (summaryList[i].Date == shareItem.Date && summaryList[i].User == "mitarbeiter" && summaryList[i].RZ == shareItem.RZ)
            {
                itemsToDelete.Add(i);
            }
            else if (summaryList[i].Date == shareItem.Date && summaryList[i].User == "cust. delivery coord/nord 2/west (feldk) organisationseinheit" && summaryList[i].RZ == shareItem.RZ)
            {
                itemsToDelete.Add(i);
            }
        }

        if (DateTime.Parse(summaryList[i].Date).Date < DateTime.Parse(shareItem.Date).Date) break;

        j++;
    }
} // Check for differences

for (int i = 0; i < summaryList.Count(); i++) 
{
    if (!itemsToDelete.Contains(i))
    {
        differencesList.Add(summaryList[i]);
    }
} // get list with only differences


// create file with final data and get name of file
generalPath = xlsxHandler.CreateWriteCheckFile(differencesList, sharePointList, generalPath);

// show collected data in Terminal
if (File.Exists(generalPath))
{
    Console.WriteLine("New File created: ");
    Console.WriteLine(generalPath);
}

Console.WriteLine();
Console.WriteLine("Following differences found:");

var table = new Table();

table.Border(TableBorder.Rounded);

table.AddColumn("Date:");
table.AddColumn(new TableColumn("User"));
table.AddColumn(new TableColumn("RZ"));

foreach (var item in differencesList)
{
    table.AddRow(item.Date, item.User, item.RZ);
}

AnsiConsole.Write(table);

Console.ReadKey();















































































//Console.WriteLine("QRZ Data =========================================================================================================");
//foreach (var item in tempQRZ)
//{
//    Console.WriteLine();
//    Console.WriteLine("Date: " + item.Date);
//    Console.WriteLine("User: " + item.User);
//    Console.WriteLine("RZ: " + item.RZ);
//}
//Console.WriteLine("QRZ Data done ====================================================================================================");
//Console.WriteLine();



//Console.WriteLine("TAG Data =========================================================================================================");
//Console.WriteLine();

//foreach (var item in tempTAG)
//{
//    Console.WriteLine("Date: " + item.Date);
//    Console.WriteLine("User: " + item.User);
//    Console.WriteLine("RZ: " + item.RZ);
//    Console.WriteLine();
//}
//Console.WriteLine("TAG Data done ====================================================================================================");


//Console.WriteLine("ZAG Data =========================================================================================================");
//Console.WriteLine();

//foreach (var item in tempZAG)
        //Console.WriteLine("QRZ Data =========================================================================================================");
        //foreach (var item in tempQRZ)
        //{
        //    Console.WriteLine();
        //    Console.WriteLine("Date: " + item.Date);
        //    Console.WriteLine("User: " + item.User);
        //    Console.WriteLine("RZ: " + item.RZ);
        //}
        //Console.WriteLine("QRZ Data done ====================================================================================================");
        //Console.WriteLine();

        //Console.WriteLine("TAG Data =========================================================================================================");
        //Console.WriteLine();

        //foreach (var item in tempTAG)
        //{
        //    Console.WriteLine("Date: " + item.Date);
        //    Console.WriteLine("User: " + item.User);
        //    Console.WriteLine("RZ: " + item.RZ);
        //    Console.WriteLine();
        //}
        //Console.WriteLine("TAG Data done ====================================================================================================");

        //Console.WriteLine("ZAG Data =========================================================================================================");
        //Console.WriteLine();

        //foreach (var item in tempZAG)
        //{
        //    Console.WriteLine("Date: " + item.Date);
        //    Console.WriteLine("User: " + item.User);
        //    Console.WriteLine("RZ: " + item.RZ);
        //    Console.WriteLine();
        //}
        //Console.WriteLine("ZAG Data done ====================================================================================================");
//{
//    Console.WriteLine("Date: " + item.Date);
//    Console.WriteLine("User: " + item.User);
//    Console.WriteLine("RZ: " + item.RZ);
//    Console.WriteLine();
//}
//Console.WriteLine("ZAG Data done ====================================================================================================");