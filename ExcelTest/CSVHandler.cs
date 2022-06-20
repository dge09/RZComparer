using System.IO;
using System.Text;

namespace ExcelTest
{
    public class CSVHandler
    {
        public List<MyExcelData> GetCsvToList(string zagPath, int startIndex)
        {
            List<MyExcelData> temp = new List<MyExcelData>(); 
            var reader = new StreamReader(zagPath, encoding: Encoding.UTF8);

            string stringToSplit;
            string date;
            string defaultDate = "1.1.0001"; // first date reference 
            string time;
            string user;
            int i = 1;


            while (!reader.EndOfStream)
            {
                stringToSplit = reader.ReadLine(); 
                
                if (i >= startIndex && IsStringEmpty(stringToSplit)) // check if line is valid
                {
                    date = stringToSplit.Substring(0,10);

                    if (date != null && date != "" && date != "          ")
                    {
                        string[] tempDate = date.Split('.');

                        if (int.Parse(tempDate[0]) < 10 && !tempDate[0].StartsWith('0'))
                        {
                            tempDate[0] = tempDate[0].TrimStart(' ');
                            tempDate[0] = '0' + tempDate[0];
                        }

                        if (int.Parse(tempDate[1]) < 10 && !tempDate[1].StartsWith('0'))
                        {
                            tempDate[1] = tempDate[1].TrimStart(' ');
                            tempDate[1] = '0' + tempDate[1];
                        }


                        date = tempDate[0] + '.' + tempDate[1] + '.' + tempDate[2];
                    } // correct formating from Date

                    user = stringToSplit.Substring(75,31).TrimEnd(' ');

                    // checks if date is empty if yes use last date used
                    if (date != null && date != "" && date != "          ") defaultDate = date; 
                    else date = defaultDate;

                    user = GetDataIntoCorrectFormat(user); // get correct format for user

                    temp.Add(
                        new MyExcelData
                        {
                            //Date = date + " " + time,
                            Date = date,
                            User = user,
                            RZ = "zag"
                        }); // write read data to temp list 
                } // read data, wtrite to list 

                i++;
            } // read data, wtrite to list


            return temp;
        } // read data, wtrite to list

        public string GetDataIntoCorrectFormat(string data)
        {
            // get rid of unnecessary parts of string
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
            } // rewritting ä,ö,ü to ae, oe, ue

            data = data.ToLower(); // make non casesensetive

            return data;
        } // reformat string to usable format

        public bool IsStringEmpty(string stringToCheck)
        {
            if (stringToCheck != null && stringToCheck != "" && stringToCheck != " ") return true;
            else return false;
        } // check if string is empty
    }
}
