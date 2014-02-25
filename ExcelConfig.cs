using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace excel_data_transfer
{
    class ExcelConfig
    {
        public string FileName { get; set; }
        public int HeaderRow { get; set; }
        public int SheetIndex { get; set; }
    }
}
