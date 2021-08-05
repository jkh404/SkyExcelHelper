using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using SkyAutoPro;
using SkyExcelHelper.Attributes;

namespace SkyExcelHelper
{
    public class ExcelHelper
    {
        
        //public XSSFWorkbook workbook { get => wb; }
        public static ExWorkbook CreateWorkBook(string WorkBookName)
        {
            AutoPro autoPro = new AutoPro();
            autoPro.Add("WorkBook",new XSSFWorkbook(workbookType:XSSFWorkbookType.XLSX));
            WorkBookName = WorkBookName == null || WorkBookName.Length == 0 ? "工作簿1" : WorkBookName;
            autoPro.Add<ExWorkbook>(WorkBookName);
            return autoPro.Get<ExWorkbook>(WorkBookName);

        }

    }
}
