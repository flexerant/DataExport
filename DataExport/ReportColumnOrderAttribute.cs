using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ReportColumnOrderAttribute : Attribute
    {
        public int Order { get; private set; }     

        public ReportColumnOrderAttribute(int order)
        {
            this.Order = order;
        }
    }
}
