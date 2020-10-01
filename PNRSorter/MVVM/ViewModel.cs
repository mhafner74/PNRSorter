using OfficeOpenXml;
using PNRSorter.Utility;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace PNRSorter.MVVM
{
    public class ViewModel : VMBase
    {
        
        #region Fields
        private ObservableCollection<MyPNR> _myPNR;
        private List<MyPNR> _PMS;
        private List<MyPNR> _SNC;
        private List<MyPNR> _KE24;
        private List<string> _PNRPMS;
        private List<string> _PNRSNC;
        #endregion

        #region Properties
        public ObservableCollection<MyPNR> MyPNR
        {
            get => _myPNR;
            set { _myPNR = value; OnPropertyChanged("MyPNR"); }
        }
        public List<MyPNR> PMS
        {
            get => _PMS;
            set { _PMS = value; OnPropertyChanged("PMS"); }
        }
        public List<MyPNR> SNC
        {
            get => _SNC;
            set { _SNC = value; OnPropertyChanged("SNC"); }
        }
        public List<MyPNR> KE24
        {
            get => _KE24;
            set { _KE24 = value; OnPropertyChanged("KE24"); }
        }
        public List<string> PNRPMS
        {
            get => _PNRPMS;
            set { _PNRPMS = value; OnPropertyChanged("PNRPMS"); }
        }
        public List<string> PNRSNC
        {
            get => _PNRSNC;
            set { _PNRSNC = value; OnPropertyChanged("PNRSNC"); }
        }
        #endregion

        #region Commands
        public ICommand TestCommand { get; set; }
        #endregion

        #region Initialise
        public void Initialise()
        {
            //Required for EEPLUS to be used
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Commands
            TestCommand = new RelayCommand(o => test(), o => true);

            //Extracting data
            FileInfo SNCFile = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\Families_S_C_v3.xlsx");
            FileInfo PMSFile = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\Families_PMS_v3.xlsx");
            FileInfo KE24File = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\KE24_2020.xlsx");
            PNRPMS = new List<string>();
            PNRSNC = new List<string>();
            SNC = ExtractedPNR(SNCFile, "S&C");
            PMS = ExtractedPNR(PMSFile, "PMS");
            KE24 = ExtractedPNR(KE24File, "/", "Product");
        }
        #endregion

        #region Constructor
        public ViewModel()
        {
            Initialise();
        }
        #endregion

        public void test()
        {
            //var file = new FileInfo(@"L:\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Families_PMS_v2.xlsx");
            List<string> uniquePNR = FindUnique(KE24);
            List<string> toBeSorted = FilterExistingPNR(uniquePNR);
            Console.WriteLine("total unique: " + uniquePNR.Count());
            Console.WriteLine("total to be sorted: " + toBeSorted.Count());
        }

        #region Private Methods
        private List<MyPNR> ExtractedPNR(FileInfo file, string group, string columnName = "")
        {
            List<MyPNR> extractedPNR = new List<MyPNR>();
            using (var package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;
                for (int i = 1; i < columns + 1; i++)
                {
                    if(columnName == "")
                    {
                        for (int j = 1; i < rows + 1; j++)
                        {
                            try
                            {
                                MyPNR curPNR = new MyPNR();
                                curPNR.Group = group;
                                curPNR.Family = worksheet.Cells[1, i].Value.ToString();
                                curPNR.PNR = worksheet.Cells[j, i].Value.ToString();
                                extractedPNR.Add(curPNR);
                                if (group == "S&C")
                                    PNRSNC.Add(curPNR.PNR);
                                else
                                    PNRPMS.Add(curPNR.PNR);
                            }
                            catch
                            {
                                break;
                            }
                        }
                    }

                    else
                    {
                        if(worksheet.Cells[1,i].Value.ToString() == columnName)
                        {
                            for (int j = 1; i < rows +1; j++)
                            {
                                try
                                {
                                    MyPNR curPNR = new MyPNR();
                                    curPNR.Group = "unknown";
                                    curPNR.Family = "unknown";
                                    curPNR.PNR = worksheet.Cells[j, i].Value.ToString();
                                    extractedPNR.Add(curPNR);
                                }
                                catch
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return extractedPNR;
        }

        private List<string> FindUnique(List<MyPNR> pnrList)
        {
            List<string> uniquePNR = new List<string>();
            foreach(var pnr in pnrList)
            {
                if (!uniquePNR.Contains(pnr.PNR))
                {
                    uniquePNR.Add(pnr.PNR);
                }
            }

            return uniquePNR;
        }

        private List<string> FilterExistingPNR(List<string> uniquePNR)
        {
            List<string> toBeSorted = new List<string>();
            foreach(var pnr in uniquePNR)
            {
                if (!(PNRPMS.Contains(pnr) || PNRSNC.Contains(pnr)))
                    toBeSorted.Add(pnr);
            }

            return toBeSorted;
        }

        private double FindColumn(FileInfo excel, string name)
        {
            double columnNb = 0;
            using (var package = new ExcelPackage(excel))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int columns = worksheet.Dimension.Columns;
                for (int i = 1; i < columns + 1; i++)
                {
                    if (name == "")
                    {
                        columnNb = i;
                        break;
                    }
                }
            }

            return columnNb;
        }
        #endregion
    }
}
