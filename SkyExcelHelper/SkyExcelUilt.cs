using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
        public static ExSheet<T> ToExSheet<T>(this DataTable dataTable, string workBookName = "WorkBook")
        {
            string sheetName = "Sheet1";
            
            if (dataTable.TableName != null && dataTable.TableName.Length > 0)
            {
                sheetName = dataTable.TableName;
            }
            ExSheet<T>  exSheet= ExcelHelper.CreateWorkBook(workBookName).CreateSheet<T>(sheetName);
            int KeyIndex = ExSheet<T>.FindKeyIndex<T>();
            if (dataTable.PrimaryKey.Length > 0) KeyIndex = dataTable.PrimaryKey[0].Ordinal;
            else throw new Exception("IsUseExPrimaryKey.Must DataTable.PrimaryKey.Length>0!");
            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                List<object> rowData = new List<object>();
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (KeyIndex == i)
                    {
                        rowData.Add(dataTable.Rows[j].ItemArray[i]);
                    }
                    else
                    {
                        rowData.Add(dataTable.Rows[j].ItemArray[i]);
                    }
                }
                T obj= ExSheet<T>.ArrayToObj(rowData.ToArray());
                exSheet.Add(obj);
            }
            return exSheet;
        }
        public static Type FindTypeByIndex<T>(int index)
        {
            return typeof(T).GetProperties().ToList()[index].PropertyType;
        }
        public static DataTable ToDataTable<T>(this ExSheet<T> exSheet)
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = exSheet.Name;
            for (int i = 0; i < exSheet.ColCount; i++)
            {
                dataTable.Columns.Add(exSheet.ColTitle[i], FindTypeByIndex<T>(i));
            }
            if (exSheet.IsUseAutoKey)
            {
                string keyTitle = exSheet.ColTitle[exSheet.KeyIndex];
                dataTable.PrimaryKey[0] = new DataColumn(keyTitle, FindTypeByIndex<T>(exSheet.KeyIndex));
            }
            for (int i = 0; i < exSheet.RowCount; i++)
            {
                dataTable.Rows.Add(ExSheet<T>.ObjToArray(exSheet[i]));
            }
            return dataTable;
        }
    }
}
