using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_Excel_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            Mytable theMytable = new Mytable();
            theMytable.MydataList = new List<Mydata>();
            for (int i = 0; i < 3; i++)
            {
                theMytable.MydataList.Add(new Mydata()
                {
                    cq = "Hi",
                    pddd = "111",
                    pdsj = "111",
                    pds1 = "111",
                    pks1 = "111",
                    pys1 = "111",
                    zms1 = "111",
                    pds2 = "111",
                    pks2 = "111",
                    pys2 = "111",
                    zms2 = "111"
                });
            }
            new Sample().exportExcelTpl(3, "File/my", exportType.png, theMytable);
        }
    }
}
