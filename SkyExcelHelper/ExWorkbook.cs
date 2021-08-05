using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SkyAutoPro;
using SkyExcelHelper.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SkyExcelHelper
{
    public class ExWorkbook
    {
        [InTag]
        public string Name { get;private set; }
        [InTag("WorkBook")]
        private XSSFWorkbook wb { get; set; }
        private List<object> Sheets { get; set; }
        private List<string> SheetNames { get; set; }
        public int Count { get => Sheets.Count(); }
        private ExWorkbook()
        {
            Sheets = new List<object>();
            SheetNames = new List<string>();
        }
        public  ExSheet<T> CreateSheet<T>(string sheetName = null)
        {
            ExSheet<T> exSheet = null;
            List<string> colTitle = new List<string>();
            AutoPro autoPro = new AutoPro();
            autoPro.Add("ExWorkbook", this);
            ExTableAttribute exTable
                = (ExTableAttribute)typeof(T)
                .GetCustomAttributes(true).ToList()
                .Find(m => m.GetType() == typeof(ExTableAttribute));
            List<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            int KeyIndex = 0;
            properties.ForEach(p => {
                
                ExColAttribute exCol = p.GetCustomAttribute<ExColAttribute>();
                ExPrimaryKeyAttribute key = p.GetCustomAttribute<ExPrimaryKeyAttribute>();
                if (exCol == null)
                {
                    colTitle.Add(p.Name);
                }
                else
                {
                    colTitle.Add(exCol.Name);
                }
                if (key!=null)
                {
                    autoPro.Add("autoPro", KeyIndex);
                    if (key.AutoNumber)
                    {
                        autoPro.Add("是否启用自增主键", true);
                    }
                    else
                    {
                        autoPro.Add("是否启用自增主键", false);
                    }
                }
                KeyIndex++;
            });
            autoPro.Add("列标题集", colTitle);
            string sheetname = null;
            string dataTableName = null;
            if (exTable != null && sheetName == null)
            {
                sheetname = exTable.Name;
                dataTableName = exTable.Name;
            }
            else if (exTable == null && sheetName == null)
            {
                sheetname = typeof(T).Name;
                dataTableName = typeof(T).Name;
            }
            else if (exTable == null && sheetName != null)
            {
                sheetname = sheetName;
                dataTableName = sheetName;
            }
            else
            {
                sheetname = sheetName;
                dataTableName = exTable.Name;
            }
            if (SheetNames.Contains(sheetname))throw new Exception("The SheetName repeats the error!");
            autoPro.Add("数据表名", dataTableName);
            ISheet NPOISheet = wb.CreateSheet(sheetname);
            autoPro.Add("NPOISheet", NPOISheet);
            autoPro.Add("NPOIWorkBook", wb);
            //properties.Select(u=>u.GetCustomAttribute<ExPrimaryKeyAttribute>(true))
            


            autoPro.Add<ExSheet<T>>(sheetname);
            exSheet = (ExSheet<T>)autoPro.Get(sheetname);
            Sheets.Add(exSheet);
            SheetNames.Add(sheetname);
            return exSheet;
        }
        public ExSheet<T> GetSheet<T>(int SheetIndex)
        {
            return (ExSheet<T>)Sheets[SheetIndex];
        }
        public ExSheet<T> GetSheet<T>()
        {
            return Sheets.Where(u=>u.GetType() == typeof(ExSheet<T>))
                .Select(u => (ExSheet<T>)u).ToList().First();
        }
        public ExSheet<T> GetSheet<T>(string SheetName)
        {
            return Sheets.Select(u => (ExSheet<T>)u).ToList().Find(m => m.Name == SheetName);
        }
        public ExWorkbook AddSheet<T>(List<T> data,string SheetName)
        {
            ExSheet<T> exSheet= CreateSheet<T>();
            exSheet.AddRang(data);
            if (SheetNames.Contains(SheetName))throw new Exception("The SheetName repeats the error!");
            SheetNames.Add(SheetName);
            Sheets.Add(exSheet);
            return this;
        }
        public object this[int SheetIndex]
        {
            get
            {
                return Sheets[SheetIndex];
            }
        }
        public XSSFWorkbook ToXSSFWorkbook => wb;

        public void SaveToFile(string fileName = null)
        {
            if (fileName == null)
            {
                fileName = $"{Name}.xlsx";
            }
            FileStream fileStream = File.OpenWrite(fileName);
            if (fileStream.CanWrite)
            {
                wb.Write(fileStream);
            }
            fileStream.Dispose();
            GC.Collect();
        }
        public void SaveToStream(Stream stream)
        {
            wb.Write(stream);
        }
    }

}
