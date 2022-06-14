using ExcelTest;

XLSXHandler xlsxHandler = new XLSXHandler();
CSVHandler csvHandler = new CSVHandler();

string generalPath = @"C:\Users\DominikG\OneDrive\02_GW\IT-Datacenter\2_1\Projekt_Zutrittskontrolle\Zutrittslisten\2021_4.Quartal";
string sharePointFile = null;
string escapeReference;

List<string> fileList = new List<string>();
List<MyExcelData> finalList = new List<MyExcelData>();
List<MyExcelData> sharePointList = new List<MyExcelData>();
List<MyExcelData> differencesList = new List<MyExcelData>();


List<MyExcelData> tempQRZ = new List<MyExcelData>();
List<MyExcelData> tempTAG = new List<MyExcelData>();
List<MyExcelData> tempZAG = new List<MyExcelData>();

fileList = Directory.GetFiles(generalPath).ToList<string>();

foreach (var file in fileList)
{
    if (file.Contains("QRZ"))
    {
        tempQRZ = xlsxHandler.GetXlsxToList(file, 2, 1, 3);

        foreach (var item in tempQRZ)
        {
            finalList.Add(item);
        }
    }
    else if (file.Contains("TAG"))
    {
        tempTAG = xlsxHandler.GetXlsxToList(file, 17, 1, 9);

        foreach (var item in tempTAG)
        {
            finalList.Add(item);
        }
    }
    else if (file.Contains("ZAG")) 
    {
        tempZAG = csvHandler.GetCsvToList(file, 9);

        foreach (var item in tempZAG)
        {
            finalList.Add(item);
        }
    }
    else if (file.Contains("Anmeldungen_Sharepoint"))
    {
        sharePointFile = file;
    }

}

finalList = finalList.DistinctBy(d => new { d.Date, d.User }).ToList();

escapeReference = xlsxHandler.GetEarlyestDate(finalList);
sharePointList = xlsxHandler.GetXlsxToList(sharePointFile, escapeReference, 2, 2, 4, 3);

sharePointList.Sort((x, y) => y.Date.CompareTo(x.Date));
sharePointList = sharePointList.DistinctBy(d => new { d.Date, d.User }).ToList();



List<int> testToDel = new List<int>();

for (int i = 0; i < finalList.Count(); i++)
{
    foreach (var shareItem in sharePointList)
    {
        if (finalList[i].Date == shareItem.Date && finalList[i].User == shareItem.User && finalList[i].RZ == shareItem.RZ)
        {
            testToDel.Add(i);
        }
    }
}

for (int i = 0; i < finalList.Count(); i++)
{
    if (!testToDel.Contains(i))
    {
        differencesList.Add(finalList[i]);
    }
}


generalPath = xlsxHandler.CreateWriteCheckFile(differencesList, sharePointList, generalPath);

Console.WriteLine();

if (File.Exists(generalPath)) Console.WriteLine("New File created: " + generalPath);











Console.ReadKey();





// lst kann auch xlsx sein      
// Mitarbeiter einfach auf Tag und Mitarbeiter prüfen












































































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