using PNRSorter.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Resources;

namespace PNRSorter.MVVM
{
    public class Family:VMBase
    {
        private string _familyName;
        public string FamilyName
        {
            get => _familyName;
            set { _familyName = value; OnPropertyChanged("FamilyName"); }
        }

        public Family(string name)
        {
            FamilyName = name;
        }

    }
}
