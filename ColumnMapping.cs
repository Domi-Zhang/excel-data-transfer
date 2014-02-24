using System;
using System.Collections.Generic;
using System.Text;

namespace excel_data_transfer
{
    class ColumnMapping
    {
        public string SourceFile { get; set; }
        public string[] SourceNames { get; set; }
        public string[] TargetNames { get; set; }
    }
}
