using System.IO;
using System.Text;

namespace ExcelTest
{
    public class CSVHandler
    {
        public List<MyExcelData> GetCsvToList(string zagPath, int startIndex)
        {
            string stringToSplit;
            string date;
            string defaultDate = "1.1.1999";
            string time;
            string user;
            int i = 1;


            List<MyExcelData> temp = new List<MyExcelData>();

            var reader = new StreamReader(zagPath, encoding: Encoding.UTF8);


            while (!reader.EndOfStream)
            {
                stringToSplit = reader.ReadLine();
                
                if (i >= startIndex && IsStringEmpty(stringToSplit))
                {
                    date = stringToSplit.Substring(0,10);
                    //time= stringToSplit.Substring(10,11).TrimEnd(' ');
                    user = stringToSplit.Substring(75,31).TrimEnd(' ');

                    if (date != null && date != "" && date != "          ") defaultDate = date;
                    else date = defaultDate;

                    user = GetDataIntoCorrectFormat(user);

                    //time = time.Replace('.', ':');

                    temp.Add(
                        new MyExcelData
                        {
                            //Date = date + " " + time,
                            Date = date,
                            User = user,
                            RZ = "zag"
                        });
                }

                i++;
            }
            

            return temp;
        }

        public string GetDataIntoCorrectFormat(string data)
        {
            if (data.Contains("(GW)")) data = data.Substring(1, data.Length - 5);
            if (data.StartsWith(" ")) data = data.TrimStart(' ');
            if (data.EndsWith(" ")) data = data.TrimEnd(' ');
            
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
                else data = data + item;
            }

            data = data.ToLower();

            return data;
        }

        public bool IsStringEmpty(string stringToCheck)
        {
            if (stringToCheck != null && stringToCheck != "" && stringToCheck != " ") return true;
            else return false;
        }
    }
}
