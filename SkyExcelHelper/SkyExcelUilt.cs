using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace SkyExcelHelper
{
    public static class  SkyExcelUilt
    {
        public static ExSheet<T> ToExSheet<T>(this List<T> dataList,string workBookName= "WorkBook",string sheetName="Sheet1")
        {
            return ExcelHelper
                .CreateWorkBook(workBookName)
                .CreateSheet<T>(sheetName)
                .AddRang(dataList);
        }
        //public static ExSheet<T> ToExSheet<T>(this DataTable dataTable, string workBookName = "WorkBook")
        //{
        //    string sheetName = "Sheet1";
        //    if (dataTable.TableName != null && dataTable.TableName.Length > 0)
        //    {
        //        sheetName = dataTable.TableName;
        //    }
        //    //dataTable.PrimaryKey[0]
        //    return null;
        //}
        //public static ExWorkbook ToExWorkbook(this DataSet dataSet)
        //{
        //    string workBookName = "WorkBook";
        //    //string sheetName = "Sheet1";
        //    if (dataSet.DataSetName != null && dataSet.DataSetName.Length > 0) 
        //        workBookName = dataSet.DataSetName;
        //    ExWorkbook exWorkbook = ExcelHelper.CreateWorkBook(workBookName);
        //    for (int i = 0; i < dataSet.Tables.Count; i++)
        //    {
        //        //exWorkbook.CreateSheet<>
        //    }
        //    //dataSet.DataSetName
        //}
    }
}
