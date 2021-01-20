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
                List<string> uniqueDates = new List<string>(UniqueElement(worksheet, colIds[1]));
                Console.Write("\n\t Extracting data");
                List<List<object>> extracted = new List<List<object>>(GetDataWithValue(worksheet, LatestEntries(uniqueDates), colIds[1], colIds));
                return extracted;
            }
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
                for (int i = 1; i < data.Count; i++)
                {
                    //always the same number of items in data[x]
                    for (int j = 0; j < data[0].Count; j++)
                    {
                        worksheet.Cells[i + nbRows, j + 1].Value = data[i - 1][j];
                    }
                }
                Console.Write("\n DONE !");
                Console.Write("\n Saving");
                package.Save();
                //File.WriteAllBytes(_KE24_EXTRACT_TOTAL, package.GetAsByteArray());
            }

            //Moving the file out of the repertory to a backup directory
            string newFile = _BACKUP + "\\" + Directory.GetFiles(_FOLDER)[0].Split('\\').Last();
            if (File.Exists(newFile))
            {
                int index = 1;
                while (File.Exists(_BACKUP + "\\" + Path.GetFileNameWithoutExtension(newFile) + "_" + index + ".xlsx"))
                    index++;
                newFile = _BACKUP + "\\" + Path.GetFileNameWithoutExtension(newFile) + "_" + index + ".xlsx";
            }
            File.Move(Directory.GetFiles(_FOLDER)[0], newFile);
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

        //Return several column ideas
        private List<int> GetColIDs(ExcelWorksheet ws, List<string> names)
        {
            List<int> colIds = new List<int>();
            foreach (string name in names)
            {
                colIds.Add(SearchColumnId(ws, name));
            }
            return colIds;
        }

        //Return list of unique element from a given column
        private List<string> UniqueElement(ExcelWorksheet ws, int ColNb)
        {
            List<string> unique = new List<string>();
            int nbRows = ws.Dimension.Rows;
            //starts at 2 to avoid header
            for (int i = 2; i < nbRows + 1; i++)
            {
                if (!unique.Contains(ws.Cells[i, ColNb].Value.ToString()))
                    unique.Add(ws.Cells[i, ColNb].Value.ToString());
            }
            return unique;
        }

        //Return latest date from the string
        private string LatestEntries(List<string> unique)
        {
            //Get latest year
            int maxYear = 0;
            int maxMonth = 0;
            string maxString = "";
            ////need to duplique unique for the foreach loop. Cannot modify the list it is based on.
            //List<string> temp = new List<string>(unique);
            List<int> years = new List<int>();
            foreach (string date in unique)
            {
                int curMonth = Convert.ToInt32(date.Split('.')[0].ToString());
                int curYear = Convert.ToInt32(date.Split('.')[1].ToString());
                if (curYear >= maxYear)
                {
                    if (curYear > maxYear)
                        maxMonth = 0;
                    if (curMonth > maxMonth)
                    {
                        //if ((maxYear != 0) && (maxMonth != 0))
                        //    temp.Remove(maxString);
                        maxMonth = curMonth;
                        maxString = date;
                    }
                    //else
                    //    temp.Remove(date);
                    maxYear = curYear;
                }
                //else
                //    temp.Remove(date);
            }
            return maxString;
            //return temp[0];
        }

        //Get all line with the proper value
        private List<List<object>> GetDataWithValue(ExcelWorksheet ws, string value, int sortingCol, List<int> colIds)
        {
            List<List<object>> extracted = new List<List<object>>();
            int nbRows = ws.Dimension.Rows;
            for (int i = 2; i < nbRows + 1; i++)
            {
                if (ws.Cells[i, sortingCol].Value.ToString() == value)
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
            ws.Column(2).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
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
