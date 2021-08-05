using System;
using System.Collections.Generic;
using System.Text;

namespace SkyExcelHelper.Attributes
{
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false)]
    public class ExColAttribute: Attribute
    {
        public string Name { get; set; }

        public ExColAttribute(string name)
        {
            Name = name;
        }
    }
}
