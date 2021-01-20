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
}
