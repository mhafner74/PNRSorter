using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PNRSorter.KE24Update
{
    public class UpdateKE24
    {
        private string _FOLDER = @"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\NewKE24";
        private string _KE24_EXTRACT_TOTAL = @"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Data\KE24_Extract_Total.xlsx";
        private string _BACKUP = @"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Archives\BackUpKE24";
        private List<string> wantedCol = new List<string>() { "Product", "Period", "Revenue", "Discount", "Direct material costs", "Direct Resource", "Direct Overhead", "Billing Quantity", "COGS", "Sales", "Gross Profit" };

        public List<List<object>> LoadNewData()
        {
            string newFile = "";
            try
            {
                newFile = Directory.GetFiles(_FOLDER)[0];
            }
            catch
            {
                MessageBox.Show("There are currently no file in " + _FOLDER + ". Please add a KE24 file in it for this function to work");
                return new List<List<object>>();
            }
            Console.Write("\n Reading the new KE24 file");
            FileInfo file = new FileInfo(newFile);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                Console.Write("\n\t gathering colums position");
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                List<int> colIds = new List<int>(GetColIDs(worksheet, wantedCol));
                //colIds[1] = period col id
                List<string> uniqueDatesNewKE24 = new List<string>(UniqueElement(worksheet, colIds[1]));
                List<string> uniqueDatesKE24Total = new List<string>(KE24Dates());
                List<string> uniqueDates = DateToBeAdded(uniqueDatesNewKE24, uniqueDatesKE24Total);
                if(uniqueDates.Count() == 0) 
                {
                    MessageBox.Show("Looks like the KE24 is already up to date");
                    Clear();
                    return new List<List<object>>();
                }
                Console.Write("\n\t Extracting data");
                List<List<object>> extracted = new List<List<object>>(GetDataWithValue(worksheet, uniqueDates, colIds[1], colIds));
                return extracted;
            }
        }

        private List<string> KE24Dates()
        {
            List<string> uniqueDates;
            FileInfo file = new FileInfo(_KE24_EXTRACT_TOTAL);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                uniqueDates = new List<string>(UniqueElement(ws, 2, "date")); 
            }
            //foreach (string date in uniqueDates)
            //    Console.WriteLine(date);
            return uniqueDates;
        }

        private List<string> DateToBeAdded(List<string> uniqueDatesNewKE24, List<string> uniqueDatesKE24Total)
        {
            List<string> ToBeAdded = new List<string>();
            SortedDictionary<int, List<int>> datesNewKE24 = ListToDic(uniqueDatesNewKE24);
            SortedDictionary<int, List<int>> datesKE24Total = ListToDic(uniqueDatesKE24Total);

            int maxYearNew = datesNewKE24.Keys.Max();
            int maxYearTot = datesKE24Total.Keys.Max();
            //Compares max years
            while (maxYearNew > maxYearTot)
            {
                foreach (int date in datesNewKE24[maxYearNew])
                    ToBeAdded.Add(String.Join(".", date.ToString().PadLeft(3,'0'), maxYearNew.ToString()));
                maxYearNew--;
            }

            //Compares max months
            if(datesNewKE24.Keys.Contains(maxYearTot))
            {
                int maxMonthNew = datesNewKE24[maxYearTot].Max();
                int maxMonthTot = datesKE24Total[maxYearTot].Max();
                if(maxMonthNew > maxMonthTot)
                {
                    for(int i = datesNewKE24[maxYearNew].Count()-1; i>=0; i--)
                    {
                        if (maxMonthTot == datesNewKE24[maxYearNew][i])
                            break;
                        else
                                ToBeAdded.Add(String.Join(".", datesNewKE24[maxYearNew][i].ToString().PadLeft(3,'0'), maxYearNew.ToString()));
                    }
                }
            }

            Console.WriteLine("\r\n****TOTAL");
            foreach (string value in ToBeAdded)
                Console.WriteLine(value);
            return ToBeAdded;
        }

        //Insert new data in excel summary sheet
        public void InsertData(List<List<object>> data)
        {
            Console.Write("\n Inserting new data into KE24 global file");
            FileInfo save = new FileInfo(_KE24_EXTRACT_TOTAL);
            using (ExcelPackage package = new ExcelPackage(save))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                FormatColumns(worksheet);
                int nbRows = worksheet.Dimension.Rows;
                Console.Write("\n Writing ...  ");
                for (int i = 1; i < data.Count + 1; i++)
                {
                    //always the same number of items in data[x]
                    for (int j = 0; j < data[0].Count; j++)
                    {
                        worksheet.Cells[i + nbRows, j + 1].Value = data[i - 1][j];
                    }
                }
                Console.Write("\n DONE !");
                Console.Write("\n Saving");
                try
                {
                    package.Save();
                }
                catch
                {
                    MessageBox.Show(@"The file \\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Data\KE24_Extract_Total.xlsx is already openned. Updates cannot be saved. Close the file and restart the procedure"); 
                }
                //File.WriteAllBytes(_KE24_EXTRACT_TOTAL, package.GetAsByteArray());
                Clear();
            }
        }

        //Clearing NewKE24 folder and storing data in Archives\BackUpKE24
        private void Clear()
        {
            //Moving the file out of the repertory to a backup directory
            string oldFile = _FOLDER + "\\" + Directory.GetFiles(_FOLDER)[0].Split('\\').Last();
            string newFile = _BACKUP + "\\" + Directory.GetFiles(_FOLDER)[0].Split('\\').Last();
            if (File.Exists(newFile))
            {
                int index = 1;
                while (File.Exists(_BACKUP + "\\" + Path.GetFileNameWithoutExtension(newFile) + "_" + index + ".xlsx"))
                    index++;
                newFile = _BACKUP + "\\" + Path.GetFileNameWithoutExtension(newFile) + "_" + index + ".xlsx";
            }
            File.Move(oldFile, newFile);
        }

        // return the col id for a given col name. Returns -1 if nothing was found
        private int SearchColumnId(ExcelWorksheet worksheet, string colName)
        {
            for (int i = 1; i < worksheet.Dimension.Columns + 1; i++)
            {
                if (worksheet.Cells[1, i].Value.ToString() == colName)
                {
                    return i;
                }
            }
            MessageBox.Show("Whoops, could not locate the \"Period\" column in the Excel file. Please make sure it is present in " + _FOLDER);
            Environment.Exit(0);
            return -1;
        }

        //Return several column IDs
        private List<int> GetColIDs(ExcelWorksheet ws, List<string> names)
        {
            List<int> colIds = new List<int>();
            foreach (string name in names)
            {
                colIds.Add(SearchColumnId(ws, name));
            }
            return colIds;
        }

        //Return list of unique element from a given column. Main problem is the cell format hence the flag "date" or "double"
        private List<string> UniqueElement(ExcelWorksheet ws, int ColNb, string flag = "double")
        {
            List<string> unique = new List<string>();
            int nbRows = ws.Dimension.Rows;
            //starts at 2 to avoid header
            if(flag == "double")
            {
                for (int i = 2; i < nbRows + 1; i++)
                {
                    if (!unique.Contains(ws.Cells[i, ColNb].Value.ToString()))
                    {
                        unique.Add(ws.Cells[i, ColNb].Value.ToString());
                    }
                }
            }
            else
            {
                for (int i = 2; i < nbRows + 1; i++)
                {
                    //The date needs to be converted into a string. The value of the cell has to be converted from an object to a string to a double to a date to a string.
                    if (!unique.Contains(DateTime.FromOADate(Convert.ToDouble(ws.Cells[i, ColNb].Value)).ToString().Split(' ')[0]))
                    {
                        unique.Add(DateTime.FromOADate(Convert.ToDouble(ws.Cells[i, ColNb].Value)).ToString().Split(' ')[0]);
                    }
                }
                //The output has to be standardised
                for(int i = 0; i< unique.Count(); i++)
                {
                    unique[i] = String.Join(".", unique[i].Split('.')[1], unique[i].Split('.')[2]);
                }
            }
            return unique;
        }

        //Trasnfrom a liast of string dates into a sorted dictionnary
        private SortedDictionary<int, List<int>> ListToDic(List<string> unique)
        {
            //Data will be sorted in dictionnary to be able to easily sort years and date later dicDate[Year][Month1, Month2,...]
            SortedDictionary<int, List<int>> dicDate = new SortedDictionary<int, List<int>>();
            foreach (string date in unique)
            {
                int curMonth = Convert.ToInt32(date.Split('.')[0].ToString());
                int curYear = Convert.ToInt32(date.Split('.')[1].ToString());
                //add entry to dicDate
                if (!dicDate.ContainsKey(curYear))
                    dicDate.Add(curYear, new List<int>() { curMonth });
                else
                    dicDate[curYear].Add(curMonth);
            }
            foreach (int year in dicDate.Keys)
                dicDate[year].Sort();
            return dicDate;
            //return temp[0];
        }

        //Get all line with the proper value
        private List<List<object>> GetDataWithValue(ExcelWorksheet ws, List<string> dates, int sortingCol, List<int> colIds)
        {
            List<List<object>> extracted = new List<List<object>>();
            int nbRows = ws.Dimension.Rows;
            for (int i = 2; i < nbRows + 1; i++)
            {
                if (dates.Contains(ws.Cells[i, sortingCol].Value.ToString()))
                {
                    List<object> temp = new List<object>();
                    foreach (int colId in colIds)
                    {
                        if (ws.Cells[i, colId].Value != null)
                            //triming the first number in the date
                            if (colId == colIds[1])
                                temp.Add(DateTime.Parse("01." + ws.Cells[i, colId].Value.ToString().Substring(1).Replace('.', '/')));
                            else if (colId == colIds[0])
                                temp.Add(ws.Cells[i, colId].Value.ToString());
                            else
                                temp.Add((ws.Cells[i, colId].Value));
                        else
                            temp.Add("N/A");
                    }
                    extracted.Add(temp);
                }
            }
            return extracted;
        }

        //Format worksheet
        private void FormatColumns(ExcelWorksheet ws)
        {
            //Number formatting
            //Proper date format for the second column
            ws.Column(2).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            //Two digits for the rest 
            for (int i = 3; i <= 11; i++)
                ws.Column(i).Style.Numberformat.Format = "0.00";

            //Background color
            ws.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Column(1).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(100, 255, 255, 153));
            ws.Column(9).Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Column(9).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(100, 248, 203, 173));
            ws.Column(10).Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Column(10).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(100, 248, 203, 173));
            ws.Column(11).Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Column(11).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(100, 248, 203, 173));
        }
    }
}
