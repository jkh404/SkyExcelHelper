using SkyExcelHelper.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace SkyExcelHelperTest
{
    [ExTable("会员用户表")]
    public class User
    {
        [ExPrimaryKey(AutoNumber =true)]
        [ExCol("ID")]
        public int id { get; set; }
        [ExCol("姓名")]
        public string Name { get; set; }
    }
}
