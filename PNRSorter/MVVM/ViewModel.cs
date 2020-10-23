﻿using Microsoft.Expression.Interactivity.Core;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using PNRSorter.Utility;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Runtime.Serialization;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PNRSorter.MVVM
{
    public class ViewModel : VMBase
    {

        #region Fields
        //Config file chaos
        private FileInfo _configFile;
        private ObservableCollection<FileInfo> _fileList;
        private FileInfo _selectedFile;
        #region PNR manipulation
        private ObservableCollection<MyPNR> _myPNR;
        private ObservableCollection<string> _groups;
        private ObservableCollection<string> _families;
        private ObservableCollection<Group> _hierarchy;
        private List<MyPNR> _PMS;
        private List<MyPNR> _SNC;
        private List<MyPNR> _KE24;
        private List<string> _PNRPMS;
        private List<string> _PNRSNC;
        private List<string> _uniquePNR;
        private ObservableCollection<ListItems> _toBeSorted;
        private string _PNRSearch;
        private ObservableCollection<string> _PNRList;
        private int _salesCol;
        private int _productCol;
        private int _unitSoldCol;
        private int _colNb;
        private Dictionary<string, int> _colInfo = new Dictionary<string, int>();
        private Dictionary<string, MyPNR> _PNRDic;
        private double _previousMinSales;
        private string _previousPNR;
        private ObservableCollection<ListItems> _toBeSortedIni;
        private Dictionary<string, List<string>> _groupToFam;
        #endregion
        #region GUI variables
        private double _curSales;
        private double _curQT;
        private string _curGroup;
        private string _curFam;
        private string _errorMsg;
        private double _minSales;
        private string _PNRFilter;
        private int _nbTBS;
        private string _selectedGroup;
        private string _selectedFam;
        private string _clickedPNR;
        #endregion
        #region Edit window
        private Object _selectedTreeItem;
        private string _modifySelection;
        private string _newGroup;
        private string _newFamily;
        private ObservableCollection<string> _groupList;
        private string _myGroup;
        #endregion
        //test variables
        private string pouet;
        #endregion

        #region Properties
        #region Config File Chaos
        public FileInfo ConfigFile
        {
            get => _configFile;
            set { _configFile = value; OnPropertyChanged("ConfigFile"); }
        }
        public ObservableCollection<FileInfo> FileList
        {
            get => _fileList;
            set { _fileList = value; OnPropertyChanged("FileList"); }
        }
        public FileInfo SelectedFile
        {
            get => _selectedFile;
            set { _selectedFile = value; OnPropertyChanged("SelectedFile"); }
        }
        #endregion
        #region Backend prop
        public ObservableCollection<MyPNR> MyPNR
        {
            get => _myPNR;
            set { _myPNR = value; OnPropertyChanged("MyPNR"); }
        }
        public ObservableCollection<string> Groups
        {
            get => _groups;
            set { _groups = value; GetAssociatedFam(); OnPropertyChanged("Groups"); }
        }
        public ObservableCollection<string> Families
        {
            get => _families;
            set { _families = value; GetAssociatedGroup(); OnPropertyChanged("Families"); }
        }
        public ObservableCollection<Group> Hierarchy
        {
            get => _hierarchy;
            set 
            { 
                _hierarchy = value;  
                OnPropertyChanged("Hierarchy");
            }
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
        public ObservableCollection<ListItems> ToBeSorted
        {
            get => _toBeSorted;
            set { _toBeSorted = value; UpdateCount(); OnPropertyChanged("ToBeSorted"); }
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
        public Dictionary<string, List<string>> GroupToFam
        {
            get => _groupToFam;
            set { _groupToFam = value; OnPropertyChanged("GroupToFam"); }
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
            set { _errorMsg = value; OnPropertyChanged("ErrorMsg"); }
        }
        public double MinSales
        {
            get => _minSales;
            set { _minSales = value; UpdateMainWin(); OnPropertyChanged("MinSales"); }
        }
        public string PNRFilter
        {
            get => _PNRFilter;
            set { _PNRFilter = value; UpdateMainWin(); OnPropertyChanged("PNRFilter"); }
        }
        public int NbTBS
        {
            get => _nbTBS;
            set { _nbTBS = value; OnPropertyChanged("NbTBS"); }
        }
        public string SelectedGroup
        {
            get => _selectedGroup;
            set { _selectedGroup = value; GetAssociatedFam(); OnPropertyChanged("SelectedGroup"); }
        }
        public string SelectedFam
        {
            get => _selectedFam;
            set { _selectedFam = value; GetAssociatedGroup(); OnPropertyChanged("SelectedFam"); }
        }
        public string ClickedPNR
        {
            get => _clickedPNR;
            set { _clickedPNR = value; OnPropertyChanged("ClickedPNR"); }
        }
        #endregion
        #region Edit win
        public Object SelectedTreeItem
        {
            get => _selectedTreeItem;
            set { _selectedTreeItem = value; UpdateEditWin(); OnPropertyChanged("SelectedTreeItem"); }
        }
        public string ModifySelection
        {
            get => _modifySelection;
            set { _modifySelection = value; OnPropertyChanged("ModifySelection"); }
        }
        public string NewGroup
        {
            get => _newGroup;
            set { _newGroup = value; OnPropertyChanged("NewGroup"); }
        }
        public string NewFamily
        {
            get => _newFamily;
            set { _newFamily = value; OnPropertyChanged("NewFamily"); }
        }
        public ObservableCollection<string> GroupList
        {
            get => _groupList;
            set { _groupList = value; OnPropertyChanged("GroupList"); }
        }
        public string MyGroup
        {
            get => _myGroup;
            set { _myGroup = value; OnPropertyChanged("MyGroup"); }
        }
        #endregion
        #region Test
        public string Pouet
        {
            get => pouet;
            set { pouet = value; Console.WriteLine("POUETPOUETPOUETPOUET" + Pouet); OnPropertyChanged("Pouet"); }
        }
        #endregion
        #endregion

        #region Commands
        public ICommand TestCommand { get; set; }
        public ICommand DisplayDataCmd { get; set; }
        public ICommand ResetCmd { get; set; }
        public ICommand SelectAllCmd { get; set; }
        public ICommand EditCmd { get; set; }
        public ICommand DeleteItemCmd { get; set; }
        public ICommand RenameItemCmd { get; set; }
        public ICommand AddGroupCmd { get; set; }
        public ICommand AddFamilyCmd { get; set; }
        #endregion

        #region Initialise
        public void Initialise()
        {
            //Required for EEPLUS to be used
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Commands
            TestCommand = new RelayCommand(o => test(), o => {return SelectedTreeItem != null; });
            DisplayDataCmd = new RelayCommand(o => DisplayData(), o => true);
            ResetCmd = new RelayCommand(o => Reset(), o => true);
            SelectAllCmd = new RelayCommand(o => SelectAll(), o => true);
            EditCmd = new RelayCommand(o => Edit(), o => true);
            DeleteItemCmd = new RelayCommand(o => DeleteItem(), o => { return SelectedTreeItem != null; });
            RenameItemCmd = new RelayCommand(o => RenameItem(), o => { return SelectedTreeItem != null; });
            AddGroupCmd = new RelayCommand(o => AddGroup(), o => { return NewGroup != null; });
            AddFamilyCmd = new RelayCommand(o => AddFamily(), o => { return ((MyGroup != null) && (NewFamily != "")); });
            //Config file location
            ConfigFile = new FileInfo(@"\\vm.dom\ns1\DATA\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\PNRSorter\configPNRSorter.txt");
            //Extracting data
            FileInfo SNCFile = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\Families_S_C_v3.xlsx");
            FileInfo PMSFile = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\Families_PMS_v3.xlsx");
            FileInfo KE24File = new FileInfo(@"C:\Users\msag\Desktop\PNRSorter\KE24_Extract_Total.xlsx");
            SavedFile sf = new SavedFile();
            sf.FileList = new ObservableCollection<FileInfo>();
            sf.FileList.Add(SNCFile);
            sf.FileList.Add(PMSFile);
            sf.FileList.Add(KE24File);
            LoadConfig(sf);
            string json = JsonConvert.SerializeObject(sf);
            //List for the autocompletion
            PNRDic = new Dictionary<string, MyPNR>();
            //List of families + groups
            Groups = new ObservableCollection<string>() { "", "S&C", "PMS" };
            //second group list used in the edit windows. Has to be different from "Groups" to allow the user not to save his modifications
            GroupList = new ObservableCollection<string>() { "S&C", "PMS" };
            GroupToFam = new Dictionary<string, List<string>>();
            GroupToFam.Add("", new List<string>());
            GroupToFam.Add("S&C", new List<string>());
            GroupToFam.Add("PMS", new List<string>());
            //List of already classified PNRs
            PNRPMS = new List<string>();
            PNRSNC = new List<string>();
            SNC = ExtractHierarchy(SNCFile, "S&C");
            PMS = ExtractHierarchy(PMSFile, "PMS");
            Hierarchy = new ObservableCollection<Group>();
            Families = new ObservableCollection<string>();
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
            ToBeSorted = new ObservableCollection<ListItems>();
            _toBeSortedIni = new ObservableCollection<ListItems>();
            GenerateDB(KE24File);
            _nbTBS = ToBeSorted.Count();
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
            SelectedGroup = "";
            NewFamily = "";
            //Win GUI intialisation
            //Verification procedures
            Pouet = "";
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

        private void Reset()
        {
            MinSales = 0;
            PNRFilter = "";
            ToBeSorted = _toBeSortedIni;
            _nbTBS = ToBeSorted.Count();
            //unselecting all PNR
            foreach (var pnr in ToBeSorted)
                pnr.IsSelected = false;
        }

        private void SelectAll()
        {
            foreach (var pnr in ToBeSorted)
                pnr.IsSelected = true;
        }


        private void UpdateCount()
        {
            NbTBS = ToBeSorted.Count();
        }

        private void Edit()
        {
            foreach (var group in Groups)
            {
                if(group != "")
                {
                    ObservableCollection<Family> famCollection = new ObservableCollection<Family>();
                    foreach (var fam in GroupToFam[group])
                        famCollection.Add(new Family(fam));
                    Hierarchy.Add(new Group(group, famCollection));
                }
            }
            EditGroupsAndFamilies editWin = new EditGroupsAndFamilies();
            editWin.Show();
        }

        public void DeleteItem()
        {
            if (SelectedTreeItem.GetType() == typeof(Group))
            {
                foreach (var group in Hierarchy)
                {
                    if (group.GroupName == ModifySelection)
                    {
                        string ini = ModifySelection;
                        Hierarchy.Remove(group);
                        GroupList.Remove(ini);
                        //for (int i = 0; i < GroupList.Count(); i++)
                        //{
                        //    if (GroupList[i] == ini)
                        //    {
                        //        GroupList.RemoveAt(i);
                        //        break;
                        //    }
                        //}
                        break;
                    }
                }
            }
            else
            {
                foreach (var group in Hierarchy)
                {
                    foreach (var family in group.Families)
                    {
                        if (family.FamilyName == ModifySelection)
                        {
                            group.Families.Remove(family);
                            break;
                        }
                    }
                }
            }
        }

        public void RenameItem()
        {
            if (SelectedTreeItem.GetType() == typeof(Group))
            {
                foreach (var group in Hierarchy)
                {
                    if (group.GroupName == ((Group)SelectedTreeItem).GroupName)
                    {
                        //we need to keep somewhere the previous word, otherwise it cannot find it in the list GroupList
                        string iniName = group.GroupName;
                        group.GroupName = ModifySelection;
                        for(int i =0; i<GroupList.Count();i++)
                        {
                            if(GroupList[i] == iniName)
                            {
                                GroupList[i] = ModifySelection;
                                break;
                            }
                        }
                        break;
                    }
                }
            }
            else
            {
                foreach (var group in Hierarchy)
                {
                    foreach (var family in group.Families)
                    {
                        if (family.FamilyName == ((Family)SelectedTreeItem).FamilyName)
                        {
                            family.FamilyName = ModifySelection;
                            break;
                        }
                    }
                }
            }
        }

        private void AddGroup()
        {
            Hierarchy.Add(new Group(NewGroup, new ObservableCollection<Family>()));
            GroupList.Add(NewGroup);
        }

        private void AddFamily()
        {
            foreach (var group in Hierarchy)
            {
                if(group.GroupName == MyGroup)
                {
                    group.Families.Add(new Family(NewFamily));
                    break;
                }
            }
        }

        private void SaveEdits()
        {
            if(MessageBox.Show("Attention ma cacahuette, tu es sur le point de modifier la hierarchie des données, on continue ?", "tu crois vraiment que je n'ai rien d'autre à faire que d'écrire une caption ?", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                foreach(Group group in Hierarchy)
                {
                    if (!Groups.Contains(group.GroupName))
                    {
                    }
                }
            }
        }

        #endregion

        public void test()
        {
            //var file = new FileInfo(@"L:\Engineering_Energy\Monthly_ProductLine_Reviews\_Dashboard\Families_PMS_v2.xlsx");
            //Console.WriteLine("total unique: " + UniquePNR.Count());
            if (SelectedTreeItem.GetType() == typeof(Group))
                Console.WriteLine("GROUP" + SelectedTreeItem);
            else
                Console.WriteLine("FAMILY" + SelectedTreeItem);
        }


        #region Private Methods
        #region Config File Mayhem
        private void LoadConfig(SavedFile sf)
        {
            if (!File.Exists(ConfigFile.FullName))
            {
                MessageBox.Show("No Config File Found, creating one: " + ConfigFile.FullName);
                if (!ConfigFile.Directory.Exists)
                {
                    try
                    {
                        ConfigFile.Directory.Create();
                    }
                    catch
                    {
                        MessageBox.Show("It is not looking greate my friend, something is going wrong. Check if you are connected to VM network. If yes, contact Jason Bourne and tell him \"\"");
                    }
                }
                File.Create(ConfigFile.FullName);
                using(StreamWriter sw = ConfigFile.AppendText())
                {
                    sw.WriteLine("PLEASE, I BEG YOU, DO NOT MODIFY THIS FILE BY YOURSELF");
                }
            }
            else
            {
                using(StreamReader sr = ConfigFile.OpenText())
                {
                    string line = sr.ReadLine();
                    if (line.StartsWith("@"))
                    {
                        FileList.Add(new FileInfo(line));
                    }
                }
            }
        }
        #endregion

        #region Main window
        private List<MyPNR> ExtractHierarchy(FileInfo file, string group)
        {
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
                                    if (!GroupToFam[group].Contains(curPNR.Family))
                                        GroupToFam[group].Add(curPNR.Family);
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
                        catch (Exception ex)
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
                    if (worksheet.Cells[j, _productCol].Value != null)
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
                                var loc = new ListItems();
                                loc.PNRName = PNR;
                                loc.IsSelected = false;
                                ToBeSorted.Add(loc);
                                _toBeSortedIni.Add(loc);
                            }
                        }
                    }
                }
            }
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

        #region Edit Win
        private void UpdateMainWin()
        {
            List<ListItems> tempString = new List<ListItems>();
            if ((MinSales < _previousMinSales) || PNRFilter != _previousPNR)
                ToBeSorted = _toBeSortedIni;
            foreach (var pnr in ToBeSorted)
            {
                if ((PNRFilter != null) && (PNRDic[pnr.PNRName].Sales > MinSales) && pnr.PNRName.StartsWith(PNRFilter))
                    tempString.Add(pnr);
            }
            //ToBeSorted.Clear();
            _previousMinSales = MinSales;
            _previousPNR = PNRFilter;
            ToBeSorted = new ObservableCollection<ListItems>(tempString);
        }

        private void UpdateEditWin()
        {
            if (SelectedTreeItem != null)
            {
                if (SelectedTreeItem.GetType() == typeof(Group))
                    ModifySelection = ((Group)SelectedTreeItem).GroupName;
                else
                    ModifySelection = ((Family)SelectedTreeItem).FamilyName;
            }
        }

        private void BackUpFiles()
        {

        }
        #endregion

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

        private void GetAssociatedFam()
        {
            {
                if (SelectedGroup != null)
                {
                    Families.Clear();
                    if (SelectedGroup == "")
                    {
                        foreach (var group in Groups)
                        {
                            foreach (var fam in GroupToFam[group])
                                Families.Add(fam);
                        }
                    }
                    else
                    {
                        foreach (var fam in GroupToFam[SelectedGroup])
                            Families.Add(fam);
                    }
                }
            }
        }

        private void GetAssociatedGroup()
        {
            if ((SelectedGroup == null) || (SelectedGroup == ""))
            {
                if (SelectedFam != null)
                {
                    foreach (var group in Groups)
                    {
                        if (GroupToFam[group].Contains(SelectedFam))
                            SelectedGroup = group;
                    }
                }
            }
        }
        #endregion

    }
}
