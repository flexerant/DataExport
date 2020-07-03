using System;
using System.Collections.Generic;
using System.Text;

namespace Flexerant.DataExport
{
    class ReportProperty
    {
        public string PropertyName { get; set; }
        public object PropertyValue { get; set; }
        public Type PropertyType { get; set; }
        public bool IsNullable { get; set; }
    }
}
