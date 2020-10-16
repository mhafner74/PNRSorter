using OfficeOpenXml;
using PNRSorter.Utility;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
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
        private List<string> _uniquePNR;
        private ObservableCollection<string> _toBeSorted;
        private string _PNRSearch;
        private ObservableCollection<string> _PNRList;
        private int _salesCol;
        private int _productCol;
        private int _unitSoldCol;
        private int _colNb;
        private Dictionary<string, int> _colInfo = new Dictionary<string, int>();
        private Dictionary<string, MyPNR> _PNRDic;
        private double _previousMinSales;
        private ObservableCollection<string> _toBeSortedIni;
        //GUI Variables
        private double _curSales;
        private double _curQT;
        private string _curGroup;
        private string _curFam;
        private string _errorMsg;
        private double _minSales;
        private string _PNRFilter;
        //test variables
        private List<string> pouet;
        #endregion

        #region Properties
        #region Backend prop
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
        public string PNRSearch
        {
            get => _PNRSearch;
            set { _PNRSearch = value; OnPropertyChanged("PNRSearch"); }
        }
        public List<string> UniquePNR
        {
            get => _uniquePNR;
            set { _uniquePNR = value; OnPropertyChanged("UniquePNR"); }
        }
        public ObservableCollection<string> ToBeSorted
        {
            get => _toBeSorted;
            set { _toBeSorted = value; OnPropertyChanged("ToBeSorted"); }
        }
        public ObservableCollection<string> PNRList
        {
            get => _PNRList;
            set { _PNRList = value; OnPropertyChanged("PNRList"); }
        }
        public Dictionary<string, int> colInfo
        {
            get => _colInfo;
            set { _colInfo = value; OnPropertyChanged("colInfo"); }
        }
        public Dictionary<string, MyPNR> PNRDic
        {
            get => _PNRDic;
            set { _PNRDic = value; OnPropertyChanged("PNRDic"); }
        }
        #endregion
        #region GUI prop
        public double CurSales
        {
            get => _curSales;
            set { _curSales = value; OnPropertyChanged("CurSales"); }
        }
        public double CurQt
        {
            get => _curQT;
            set { _curQT = value; OnPropertyChanged("CurQt"); }
        }
        public string CurFam
        {
            get => _curFam;
            set { _curFam = value; OnPropertyChanged("CurFam"); }
        }
        public string CurGroup
        {
            get => _curGroup;
            set { _curGroup = value; OnPropertyChanged("CurGroup"); }
        }
        public string ErrorMsg
        {
            get => _errorMsg;
            set { _errorMsg = value;OnPropertyChanged("ErrorMsg"); }
        }
        public double MinSales
        {
            get => _minSales;
            set { _minSales = value; OnPropertyChanged("MinSales"); }
        }
        public string PNRFilter
        {
            get => _PNRFilter;
            set { _PNRFilter = value; OnPropertyChanged("PNRFilter"); }
        }
        #endregion
        #endregion

        #region Commands
        public ICommand TestCommand { get; set; }
        public ICommand DisplayDataCommand { get; set; }
        public ICommand FilterCommand { get; set; }
        #endregion

        #region Initialise
        public void Initialise()
        {
            //Required for EEPLUS to be used
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Commands
            TestCommand = new RelayCommand(o => test(), o => true);
            DisplayDataCommand = new RelayCommand(o => DisplayData(), o => true);
            FilterCommand = new RelayCommand(o => Filter(), o => true);
            //Extracting data
            FileInfo SNCFile = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\Families_S_C_v3.xlsx");
            FileInfo PMSFile = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\Families_PMS_v3.xlsx");
            FileInfo KE24File = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\KE24_Extract_Total.xlsx");
            //List for the autocompletion
            PNRDic = new Dictionary<string, MyPNR>();
            //List of already classified PNRs
            PNRPMS = new List<string>();
            PNRSNC = new List<string>();
            SNC = ExtractHierarchy(SNCFile, "S&C");
            PMS = ExtractHierarchy(PMSFile, "PMS");
            //Selecting the desired values
            colInfo = new Dictionary<string, int>();
            colInfo.Add("Product", -1);
            colInfo.Add("Sales", -1);
            colInfo.Add("Billing Quantity", -1);
            FindColumn(KE24File, colInfo);
            foreach (var key in colInfo.Keys)
                Console.WriteLine("key: " + key + ", colNb: " + colInfo[key].ToString());
            _productCol = colInfo["Product"];
            _salesCol = colInfo["Sales"];
            _unitSoldCol = colInfo["Billing Quantity"];
            ToBeSorted = new ObservableCollection<string>();
            _toBeSortedIni = new ObservableCollection<string>();
            GenerateDB(KE24File);
            PNRList = new ObservableCollection<string>(PNRDic.Keys.ToList());
            //Isolating PNRs
            //UniquePNR = FindUnique(KE24);
            //ToBeSorted = new ObservableCollection<string>(FilterExistingPNR(PNRDic.Keys.ToList()));
            //GUI Initialisation
            CurGroup = "N/A";
            CurFam = "N/A";
            CurSales = 0.0;
            CurQt = 0.0;
            ErrorMsg = "";
            MinSales = 0;
            PNRFilter = "";
            //Verification procedures
        }
        #endregion

        #region Constructor
        public ViewModel()
        {
            Initialise();
        }
        #endregion

        #region Command Methods
        private void DisplayData()
        {
            try
            {
                CurGroup = PNRDic[PNRSearch].Group;
                CurFam = PNRDic[PNRSearch].Family;
                CurSales = PNRDic[PNRSearch].Sales;
                CurQt = PNRDic[PNRSearch].UnitSold;
                ErrorMsg = "";
            }
            catch
            {
                ErrorMsg = "Something went wrong, please check PNR";
            }
        }
        #endregion

        public void test()
        {
            //var file = new FileInfo(@"L:\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Families_PMS_v2.xlsx");
            //Console.WriteLine("total unique: " + UniquePNR.Count());
            Console.WriteLine("total to be sorted: " + ToBeSorted.Count());
        }

        #region Private Methods
        private List<MyPNR> ExtractHierarchy(FileInfo file, string group)
        {
            pouet = new List<string>();
            List<MyPNR> ExtractHierarchy = new List<MyPNR>();
            using (var package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rows = worksheet.Dimension.Rows;
                int col = worksheet.Dimension.Columns;
                for (int i = 1; i < col + 1; i++)
                {
                    for (int j = 2; j < rows + 1; j++)
                    {
                        MyPNR curPNR = new MyPNR();
                        try
                        {
                            if (!string.IsNullOrEmpty(worksheet.Cells[j, i].Text))
                            {
                                string PNR = worksheet.Cells[j, i].Value.ToString();
                                if (!PNRDic.Keys.Contains(PNR))
                                {
                                    if (group == "S&C")
                                        PNRSNC.Add(curPNR.PNR);
                                    else
                                        PNRPMS.Add(curPNR.PNR);
                                    curPNR.Family = worksheet.Cells[1, i].Value.ToString();
                                    curPNR.Group = group;
                                    curPNR.PNR = worksheet.Cells[j, i].Value.ToString();
                                    curPNR.Sales = 0;
                                    curPNR.UnitSold = 0;
                                    PNRDic.Add(PNR, curPNR);
                                }
                                else
                                    Console.WriteLine("DOUBLON: " + PNR);
                            }
                        }
                        catch(Exception ex) 
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }
                }
            }
            return ExtractHierarchy;
        }

        private void GenerateDB(FileInfo file)
        {
            using (var package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rows = worksheet.Dimension.Rows;
                //rows have to start at 2, we don't want headers
                for (int j = 2; j < rows + 1; j++)
                {
                    if(worksheet.Cells[j, _productCol].Value != null)
                    {
                        string PNR = worksheet.Cells[j, _productCol].Value.ToString();
                        MyPNR curPNR = new MyPNR();
                        if (PNRDic.Keys.Contains(PNR))
                        {
                            //To be uncommented if data is also wanted for PNRs that have already been classified 
                            PNRDic[PNR].Sales += Convert.ToDouble(worksheet.Cells[j, _salesCol].Value);
                            PNRDic[PNR].UnitSold += Convert.ToDouble(worksheet.Cells[j, _unitSoldCol].Value);
                        }
                        else
                        {
                            curPNR.PNR = PNR;
                            curPNR.Family = "unknown";
                            curPNR.Group = "unknown";
                            if (PNRDic.Keys.Contains("PNR"))
                            {
                                PNRDic[PNR].Sales += Convert.ToDouble(worksheet.Cells[j, _salesCol].Value);
                                PNRDic[PNR].UnitSold += Convert.ToDouble(worksheet.Cells[j, _unitSoldCol].Value);
                            }
                            else
                            {
                                curPNR.Sales = Convert.ToDouble(worksheet.Cells[j, _salesCol].Value);
                                curPNR.UnitSold = Convert.ToDouble(worksheet.Cells[j, _unitSoldCol].Value);
                                PNRDic.Add(PNR, curPNR);
                                ToBeSorted.Add(PNR);
                                _toBeSortedIni.Add(PNR);
                            }
                        }
                        if ((PNR == "200-582-200-021"))
                        {
                            Console.WriteLine("sales:" + worksheet.Cells[j, _salesCol].Text);
                            Console.WriteLine("billing qt:" + worksheet.Cells[j, _unitSoldCol].Text);
                            pouet.Add("o");
                        }
                    }
                } 
            }
            Console.WriteLine(pouet.Count().ToString());
            Console.WriteLine("Total sales:" + PNRDic["200-582-200-021"].Sales.ToString());
            Console.WriteLine("Total billing qt:" + PNRDic["200-582-200-021"].UnitSold.ToString());
        }

        private void Filter()
        {
            List<string> tempString = new List<string>();
            if (MinSales < _previousMinSales)
                ToBeSorted = _toBeSortedIni;
            foreach(var pnr in ToBeSorted)
            {
                if ((PNRDic[pnr].Sales > MinSales) && pnr.Contains(PNRFilter))
                    tempString.Add(pnr);
            }
            //ToBeSorted.Clear();
            _previousMinSales = MinSales;
            ToBeSorted = new ObservableCollection<string>(tempString);
        }

        private List<string> FindUnique(List<MyPNR> pnrList)
        {
            List<string> uniquePNR = new List<string>();
            foreach (var pnr in pnrList)
            {
                if (!uniquePNR.Contains(pnr.PNR))
                {
                    uniquePNR.Add(pnr.PNR);
                    PNRList.Add(pnr.PNR);
                }
            }

            return uniquePNR;
        }

        private List<string> FilterExistingPNR(List<string> uniquePNR)
        {
            List<string> toBeSorted = new List<string>();
            foreach (var pnr in uniquePNR)
            {
                if (!(PNRPMS.Contains(pnr) || PNRSNC.Contains(pnr)))
                    toBeSorted.Add(pnr);
            }

            return toBeSorted;
        }

        private void FindColumn(FileInfo excel, Dictionary<string, int> name)
        {
            using (var package = new ExcelPackage(excel))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int columns = worksheet.Dimension.Columns;
                for (int i = 1; i < columns + 1; i++)
                {
                    if (name.Keys.Contains(worksheet.Cells[1, i].Value.ToString()))
                    {
                        name[worksheet.Cells[1, i].Value.ToString()] = i;
                    }
                }
            }

            return;
        }
        #endregion

        #region Research
        private MyPNR PNRFetcher(FileInfo excel, string pnrNb)
        {
            using (var package = new ExcelPackage(excel))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                for (int j = 1; j < (_colNb + 1); j++)
                {
                    //if(worksheet.Cells[j, _productCol].Value.ToString() == pnrNb)
                }
            }
            MyPNR pnr = new MyPNR();
            return pnr;
        }
        #endregion
    }
}
