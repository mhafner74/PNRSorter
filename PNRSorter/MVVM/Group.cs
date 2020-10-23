using PNRSorter.Utility;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace PNRSorter.MVVM
{
    public class Group : VMBase, ITreeItem
    {
        private ObservableCollection<Family> _families;
        public ObservableCollection<Family> Families
        {
            get => _families;
            set { _families = value; OnPropertyChanged("Families"); }
        }

        private string _groupName;
        public string GroupName
        {
            get => _groupName;
            set { _groupName = value; OnPropertyChanged("GroupName"); }
        }

        public Group(string name, ObservableCollection<Family> famList)
        {
            GroupName = name;
            _families = famList;
        }

        public string GetName()
        {
            return _groupName;
        }
    }
}
