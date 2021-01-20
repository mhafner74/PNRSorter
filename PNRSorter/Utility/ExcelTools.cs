using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRSorter.Utility
{
    public class ExcelTools
    {
        public bool IsFileOpen(string filePath)
        {
            bool isOpen = false;
            try
            {
                System.IO.FileStream fs = System.IO.File.OpenWrite(filePath);
                fs.Close();
            }
            catch (System.IO.IOException)
            {
                isOpen = true;
            }
            return isOpen;
        }

    }

}
