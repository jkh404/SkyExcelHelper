using SkyExcelHelper;
using System;
using System.Collections.Generic;

namespace SkyExcelHelperTest
{
    class Program
    {
        static void Main(string[] args)
        {
            ExWorkbook exWorkbook = ExcelHelper.CreateWorkBook("测试");
            ExSheet<User>  UserSheet = exWorkbook.CreateSheet<User>("用户表");
            //int ID = 0;
            UserSheet.Add(new User() {Name = "sky" })
                .Add(new User() { Name = "sky2" })
                .Add(new User() { Name = "sky3" })
                .Add(new User() { Name = "sky4" })
                .Submit().SaveToFile();

        }
    }
}
