using OfficeOpenXml;
using OfficeOpenXml.Style;
using PNRSorter.MVVM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using PNRSorter.Utility;
using System.Xml.Serialization;

namespace PNRSorter.KE24Update
{
    public class UpdateKE24:VMBase
    {
        //struture dossier
        //    Archives
        //    BackUpKE24
        //    Data
        //    NewKE24

        //private string NewKE24Directory = @"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\NewKE24";
        //private string KE24ExtractTotalFile = @"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Data\KE24_Extract_Total.xlsx";
        //private string BackupDirectory = @"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Archives\BackUpKE24";
        private List<string> wantedCol = new List<string>() { "Product", "Period", "Revenue", "Discount", "Direct material costs", "Direct Resource", "Direct Overhead", "Billing Quantity", "COGS", "Sales", "Gross Profit" };
        private double _COGS;
        private double _revenue;
        public string BackupDirectory { get
            {
                return Path.Combine(DashboardDirectory, "BackUpKE24");
            }
        }
        public string NewKE24Directory
        {
            get
            {
                return Path.Combine(DashboardDirectory, "NewKE24");
            }
        }
        public string KE24ExtractTotalFile
        {
            get
            {
                return Path.Combine(DashboardDirectory, "Data\\KE24_Extract_Total.xlsx");
            }
        }
        public string DashboardDirectory   { get; private set; }
        public UpdateKE24(string dashboardDirectory) {
            _COGS = 1;
            _revenue = 1;
            DashboardDirectory = dashboardDirectory;
        }
        public UpdateKE24(double COGS, double Revenue)
        {
            _COGS = COGS;
            _revenue = Revenue;
        }
        public List<List<object>> LoadNewData()
        {
            string newFile = "";
            try
            {
                newFile = Directory.GetFiles(NewKE24Directory)[0];
            }
            catch
            {
                MessageBox.Show("There are currently no file in " + NewKE24Directory + ". Please add a KE24 file in it for this function to work");
                return new List<List<object>>();
            }
            Console.Write("\n Reading the new KE24 file");
            FileInfo file = new FileInfo(newFile);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                Console.Write("\n\t gathering colums position");
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "KE24");
                List<int> colIds = new List<int>(GetColIDs(worksheet, wantedCol));
                //colIds[1] = period col id
                List<string> uniqueDatesNewKE24 = new List<string>(UniqueElement(worksheet, colIds[1]));
                List<string> uniqueDatesKE24Total = new List<string>(KE24Dates());
                if (uniqueDatesKE24Total.Count == 0)
                {
                    MessageBox.Show("Error, procedure aborted", "", MessageBoxButton.OK, MessageBoxImage.Error);
                    return new List<List<object>>();
                }
                else
                {
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
        }

        public int CheckValues(double COGS, double revenue)
        {
            MessageBoxResult decision = MessageBox.Show("New data will be added to the KE24 file using the following multipliers: \nRevenue = " + revenue.ToString() + "\nCOGS = " + COGS.ToString(), "", MessageBoxButton.YesNoCancel);
            if (decision == MessageBoxResult.Cancel)
                return -1;
            if (decision == MessageBoxResult.No)
                return 0;
            else
                return 1;
        }

        private List<string> KE24Dates()
        {
            List<string> uniqueDates;
            FileInfo file = new FileInfo(KE24ExtractTotalFile);
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
            FileInfo save = new FileInfo(KE24ExtractTotalFile);
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
                //File.WriteAllBytes(KE24ExtractTotalFile, package.GetAsByteArray());
                Clear();
            }
        }

        //Clearing NewKE24 folder and storing data in Archives\BackUpKE24
        private void Clear()
        {
            //Moving the file out of the repertory to a backup directory
            string oldFile = NewKE24Directory + "\\" + Directory.GetFiles(NewKE24Directory)[0].Split('\\').Last();
            string newFile = BackupDirectory + "\\" + Directory.GetFiles(NewKE24Directory)[0].Split('\\').Last();
            if (!Directory.Exists(BackupDirectory))
            {
                Directory.CreateDirectory(BackupDirectory);
            }
            if (File.Exists(newFile))
            {
                int index = 1;
                while (File.Exists(BackupDirectory + "\\" + Path.GetFileNameWithoutExtension(newFile) + "_" + index + ".xlsx"))
                    index++;
                newFile = BackupDirectory + "\\" + Path.GetFileNameWithoutExtension(newFile) + "_" + index + ".xlsx";
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
            MessageBox.Show("Whoops, could not locate the \"Period\" column in the Excel file. Please make sure it is present in " + NewKE24Directory);
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
            if (flag == "date")
            {
                for (int i = 2; i < nbRows + 1; i++)
                {
                    //The date needs to be converted into a string. The value of the cell has to be converted from an object to a string to a double to a date to a string.
                    object value = ws.Cells[i, ColNb].Value;
                    if (value!=null && value.GetType() == typeof(DateTime))
                    {
                        string test = Convert.ToString(value).ToString().Split(' ')[0];
                        if (!unique.Contains(test))
                        { unique.Add(test); }
                    }
                    else
                    {
                        //DateTime test1 = DateTime.FromOADate(Convert.ToDouble(value));
                        //string test = DateTime.FromOADate(Convert.ToDouble(value)).ToString();
                        if (!unique.Contains(DateTime.FromOADate(Convert.ToDouble(value)).ToString().Split(' ')[0]))
                        {
                            unique.Add(DateTime.FromOADate(Convert.ToDouble(value)).ToString().Split(' ')[0]);
                        }
                    }
                     
                }
                //The output has to be standardised with separator being "." between days, month and years
                for(int i = 0; i< unique.Count(); i++)
                {
                    try
                    {
                        //string[] test = unique[i].Split('/');
                        //string[] test1 = unique[i].Split('.');
                        if (unique[i].Split('/').Length > 1)
                        {
                            unique[i] = String.Join(".", unique[i].Split('/')[1], unique[i].Split('/')[2]);
                        }
                        else if (unique[i].Split('.').Length > 1)
                        {
                            unique[i] = String.Join(".", unique[i].Split('.')[1], unique[i].Split('.')[2]);
                        }
                        else
                        {
                            unique[i] = String.Join(".", unique[i].Split(' ')[1], unique[i].Split(' ')[2]);
                        }
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("Your default Windows region is not correct (that is what you get for using Excel). Please change it to English (Switzerland) so the date is with the format dd/MM/yyyy or dd.MM.yyyy","",MessageBoxButton.OK,MessageBoxImage.Error);
                        unique.Clear();
                        return unique;
                    }
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
                            else if(colId == colIds[2])
                            {
                                double pouet = Convert.ToDouble(ws.Cells[i, colId].Value);
                                double tempValue = Convert.ToDouble(ws.Cells[i, colId].Value) * _revenue;
                                temp.Add(tempValue);
                            }
                            else if (colId == colIds[8])
                            {
                                double pouet2 = Convert.ToDouble(ws.Cells[i, colId].Value);
                                double tempValue = Convert.ToDouble(ws.Cells[i, colId].Value) * _COGS;
                                temp.Add(tempValue);
                            }
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
