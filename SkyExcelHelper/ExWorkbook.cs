using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SkyAutoPro;
using SkyExcelHelper.Attributes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
        [InTag("NPOIWorkBook")]
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
            
            AutoPro autoPro = new AutoPro();
            autoPro.Add("ExWorkbook", this);
            ExTableAttribute exTable
                = (ExTableAttribute)typeof(T)
                .GetCustomAttributes(true).ToList()
                .Find(m => m.GetType() == typeof(ExTableAttribute));

            List<string> colTitle = ExSheet<T>.GetColTitle<T>((key, KeyIndex) => {
                if (key != null)
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

        public ExSheet<T> FromDB<T>(string connectionString,Func<SqlCommand, SqlCommand> CallBack)
        {
            ExSheet<T> exSheet = null;
            SqlConnection connection = new SqlConnection(connectionString);
            if (connection.State == ConnectionState.Closed)
            {
                connection.Open();
            }

            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            dataAdapter.SelectCommand = CallBack(connection.CreateCommand());
            DataSet dataSet = new DataSet();
            dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
            dataAdapter.Fill(dataSet);
            if (dataSet.Tables.Count>0)
            {
                DataTable dataTable = dataSet.Tables[0];
                exSheet=dataTable.ToExSheet<T>(Name);
            }
            connection.Close();
            connection.Dispose();
            dataAdapter.Dispose();
            dataSet.Dispose();

            return exSheet;
        }
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
