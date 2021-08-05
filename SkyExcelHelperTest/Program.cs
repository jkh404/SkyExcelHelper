using SkyExcelHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace SkyExcelHelperTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExWorkbook exWorkbook = ExcelHelper.CreateWorkBook("测试");
            //ExSheet<User> UserSheet = exWorkbook.CreateSheet<User>("用户表");
            ////int ID = 0;
            //UserSheet.Add(new User() { Name = "sky" })
            //    .Add(new User() { Name = "sky2" })
            //    .Add(new User() { Name = "sky3" })
            //    .Add(new User() { Name = "sky4" })
            //    .Submit().SaveToFile();
            //string conStr = "Server=127.0.0.1,1433;Database=studentMg;" +
            //    "User=sa;Password=xxxx;MultipleActiveResultSets=True;";
            //ExcelHelper.CreateWorkBook("测试2")
            //    .FromDB<StudentClass>(conStr, cmd=> {
            //        cmd.CommandText = "select * from 班级 as test";
            //        return cmd;
            //    })
            //    .Submit().SaveToFile();

            ExWorkbook exWorkbook = ExcelHelper.CreateWorkBook("测试");
            ExSheet<User> UserSheet = exWorkbook.CreateSheet<User>("用户表");
            int ID = 0;
            UserSheet.Add(new User() { id = ID++, Name = "sky" })
                .Add(new User() { id = ID++, Name = "sky2" })
                .Add(new User() { id = ID++, Name = "sky3" })
                .Add(new User() { id = ID++, Name = "sky4" });
            List<User> users= UserSheet.ToList();
            DataTable UserTable=users.ToExSheet<User>().ToDataTable<User>();
            UserSheet.Submit().SaveToFile();
        }
    }
}
