using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRSorter.Utility
{
    public class ListItems : VMBase
    {
        private string _PNRName;
        public string PNRName
        {
            get => _PNRName;
            set { _PNRName = value; OnPropertyChanged("PNRName"); }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged("IsSelected"); }
        }
    }
}
