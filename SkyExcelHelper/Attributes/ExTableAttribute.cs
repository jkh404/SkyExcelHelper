using System;
using System.Collections.Generic;
using System.Text;

namespace SkyExcelHelper.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ExTableAttribute:Attribute
    {
        public string Name { get; set; }

        public ExTableAttribute(string name)
        {
            Name = name;
        }
    }
}
