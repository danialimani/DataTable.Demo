using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DI.DataTable
{
    public class PropertyOption
    {
        public string DisplayName { get; set; }
        //for Display Custom Colunms
        public bool IsShow { get; set; }
        //show and hide for end user
        public bool IsVisible { get; set; }
        // Filter Column Is not Empty
        public bool IsFilterActive { get; set; }
        // Filter Column Value
        public string FilterValue { get; set; }
        // Filter Column Type
        public string FilterType { get; set; }
    }
}
