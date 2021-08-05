using System;
using System.Collections.Generic;
using System.Text;

namespace SkyExcelHelper.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExPrimaryKeyAttribute : Attribute
    {
        public bool AutoNumber { get; set; }
    }
}
