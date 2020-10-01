using PNRSorter.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRSorter.MVVM
{
    public class MyPNR : ObservableObject
    {
        #region Fields
        private string _PNR;
        private string _family;
        private string _group;
        private double _sales;
        private double _unitSold;
        #endregion

        #region Properties
        public string PNR
        {
            get => _PNR;
            set { _PNR = value; OnPropertyChanged("PNR"); }
        }
        public string Family
        {
            get => _family;
            set { _family = value; OnPropertyChanged("Family"); }
        }
        public string Group
        {
            get => _group;
            set { _group = value; OnPropertyChanged("Group"); }
        }
        public double Sales
        {
            get => _sales;
            set { _sales = value; OnPropertyChanged("Sales"); }
        }
        public double UnitSold
        {
            get => _unitSold;
            set { _unitSold = value; OnPropertyChanged("UnitSold"); }
        }
        #endregion

        #region Constructors
        public MyPNR() { }

        public MyPNR(string PNR, string Family, string Group, double Sales, double UnitSold)
        {
            this._PNR = PNR;
            this._family = Family;
            this._group = Group;
            this._sales = Sales;
            this._unitSold = UnitSold;
        }
        #endregion
    }
}
