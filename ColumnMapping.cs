using System;
using System.Collections.Generic;
using System.Text;

namespace excel_data_transfer
{
    class ColumnMapping
    {
        public string SourceFile { get; set; }

        private string[] m_sourceNames;
        public string[] SourceNames 
        { 
            get 
            {
                if (m_sourceNames == null) 
                {
                    m_sourceNames = SourceName.Split('|');
                }
                return m_sourceNames;
            } 
        }

        private string[] m_targetNames;
        public string[] TargetNames
        {
            get
            {
                if (m_targetNames == null)
                {
                    m_targetNames = TargetName.Split('|');
                }
                return m_targetNames;
            }
        }

        public string SourceName { get; set; }
        public string TargetName { get; set; }
    }
}
