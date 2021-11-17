using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRSorter.MVVM
{
    public class SavedFile
    {
        public List<GroupFile> GroupList { get; set; }
        public NumData NumParam { get; set; }
        public List<string> StringGroupList()
        {
            List<string> stringGroupList = new List<string>(); 
            foreach (var group in GroupList)
            {
                stringGroupList.Add(group.Name);
            }

            return stringGroupList;
        }
        public SavedFile()
        {
            GroupList = new List<GroupFile>();
            NumParam = new NumData();
        }
    }

    public class GroupFile
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public GroupFile()
        {
            Name = "";
            Path = "";
        }
    }

    public class NumData
    {
        public double Revenue { get; set; }
        public double COGS { get; set; }
        public NumData()
        {
            Revenue = 1.0;
            COGS = 1.0;
        }
    }
}
