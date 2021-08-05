using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using SkyAutoPro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SkyExcelHelper
{

    public class ExSheet<T>
    {
        
        [InTag]
        private string SheetName { get; set; }
        [InTag("数据表名")]
        public string Name { get; private set; }
        [InTag("标题")]
        public string Title { get; private set; }
        [InTag("列标题集")]
        public List<string> ColTitle { get; private set; }

        [InTag("NPOISheet")]
        private ISheet NPOISheet { get; set; }
        [InTag("NPOIWorkBook")]
        private XSSFWorkbook NPOIWorkBook { get; set; }
        [InTag("ExWorkbook")]
        private ExWorkbook exWorkbook { get; set; }

        [InTag("是否启用自增主键")]
        public bool IsUseAutoKey { get;private set; }
        [InTag("主键所在位置")]
        private int KeyIndex { get; set; } = -1;

        private List<T> RowDataEx { get; set; }
        public int RowCount { get => RowDataEx.Count(); }
        public int ColCount { get => ColTitle.Count(); }
        private int AutoID { get; set; } = 0;
        private HashSet<object> Keys { get; set; }

        private ExSheet()
        {
            RowDataEx = new List<T>();
            Keys = new HashSet<object>();
        }

        public ExSheet<T> Add(T obj)
        {
            if (KeyIndex!=-1)
            {
                if (IsUseAutoKey)
                {
                    SetColValue(obj, KeyIndex, AutoID++);
                }
                object key = GetColValue(obj, KeyIndex);
                if (!Keys.Contains(key))
                {
                    Keys.Add(key);
                }
                else
                {
                    throw new Exception("The primary key recurs!");
                }
            }
            RowDataEx.Add(obj);
            return this;
        }
        public ExSheet<T> AddRang(List<T> objs)
        {
            objs.ForEach(obj=> Add(obj));
            return this;
        }
        public static object[] ObjToArray(T obj)
        {
            return obj.GetType()
                .GetProperties()
                .ToList()
                .Select(s =>
                {
                    object value= s.GetGetMethod()?.Invoke(obj, new object[0]);
                    return value;
                }).ToArray();
        }
        public static T ArrayToObj(object[] array) 
        {
            ConstructorInfo constructor= typeof(T).GetConstructors().Where(c=>c.GetParameters().Count()==0).First();
            T obj = (T)(constructor?.Invoke(null));
            int index = 0;
            obj.GetType().GetProperties().ToList().ForEach((p)=> {
                p.GetSetMethod()?.Invoke(obj,new object[] { array[index++] });
            });
            return obj;
        }
        public T this[int rowIndex]
        {
            get
            {

                return RowDataEx[rowIndex];
            }
            set
            {
                RowDataEx[rowIndex] =value;
            }
        }
        public object this[int rowIndex,int colIndex]
        {
            get
            {
                return ObjToArray(RowDataEx[rowIndex])[colIndex];
            }
            set
            {
                SetColValue(RowDataEx[rowIndex], colIndex, value);
            }
        }
        public List<object> this[string colTitle]
        {
            get
            {
                int colIndex=ColTitle.IndexOf(colTitle);
                if (colIndex == -1) return null;
                return RowDataEx.Select(u =>ObjToArray(u)[colIndex]).ToList();
            }
        }
        public ExSheet<T> Insert(int index, T obj)
        {
            RowDataEx.Insert(index, obj);
            return this;
        }
        public int IndexOf(T obj)
        {
            return RowDataEx.IndexOf(obj);
        }
        public ExSheet<T> Update(int rowIndex, T newObj)
        {
            RowDataEx[rowIndex] = newObj;
            return this;
        }
        public ExSheet<T> Remove(T obj)
        {
            RowDataEx.Remove(obj);
            return this;
        }
        public ExSheet<T> RemoveAt(int rowIndex)
        {
            RowDataEx.RemoveAt(rowIndex);
            return this;
        }
        public ExSheet<T> RemoveAll(Predicate<T> match)
        {
            RowDataEx.RemoveAll(match);
            return this;
        }
        public void Clear()
        {
            RowDataEx.Clear();
        }
        private void Save()
        {
            int index = 0;
            ICellStyle cellstyle = NPOIWorkBook.CreateCellStyle();
            cellstyle.VerticalAlignment = VerticalAlignment.Center;
            cellstyle.Alignment = HorizontalAlignment.Center;

            IRow DataTitle= NPOISheet.CreateRow(index);
            ICell DataTitleCell = DataTitle.CreateCell(0);
            DataTitleCell.SetCellValue(Name);
            DataTitleCell.CellStyle = cellstyle;
            if (ColCount>=2)
            {
                NPOISheet.AddMergedRegion(new CellRangeAddress(index, index, 0, ColCount-1));
            }
            IRow ColTitleCell = NPOISheet.CreateRow(++index);
            for (int i = 0; i < ColTitle.Count; i++)
            {
                ICell cell= ColTitleCell.CreateCell(i);
                cell.SetCellValue(ColTitle[i]);
            }
            int rowIndex = 0;
            foreach (var item in RowDataEx)
            {
                int colIndex = 0;
                IRow row= NPOISheet.CreateRow(++index);
                CreateCell(row).ForEach((cell)=> {
                    cell.SetCellValue(this[rowIndex,colIndex++].ToString());
                });
                rowIndex++;
            }
        }
        private List<ICell> CreateCell(IRow row)
        {
            List<ICell> cells = new List<ICell>();
            for (int i = 0; i < ColTitle.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cells.Add(cell);
            }
            return cells;
        }
        public List<T> ToList()
        {
            return RowDataEx;
        }
        public ISheet ToISheet()
        {
            return NPOISheet;
        }
        public ExWorkbook Submit()
        {
            Save();
            return exWorkbook;
        }

        private static object GetColValue(T obj,int col)
        {
            List<PropertyInfo> properties= obj.GetType().GetProperties().ToList();
            object value= properties[col].GetGetMethod()?.Invoke(obj,new object[0]);
            return value;
        }
        private static void SetColValue(T obj,int col,object newValue)
        {
            List<PropertyInfo> properties= obj.GetType().GetProperties().ToList();
            properties[col].GetSetMethod()?.Invoke(obj,new object[] { newValue });
        }
        private static int FindKeyIndex<T>()
        {
            List<PropertyInfo> properties = typeof(T).GetType().GetProperties().ToList();
            for (int i = 0; i < properties.Count(); i++)
            {
                bool ok = false;
                ok = properties[i].Name.Contains("ID")|| ok;
                ok = properties[i].Name.Contains("Id")|| ok;
                if (ok)
                {
                    return i;
                }
            }
            return -1;
        }

    }
}