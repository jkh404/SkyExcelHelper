using SkyExcelHelper.Attributes;

namespace SkyExcelHelperTest
{
    [ExTable("班级表")]
    internal class StudentClass
    {
        [ExPrimaryKey]
        [ExCol("班级代码")]
        public string Id { set; get; }
        [ExCol("班级名称")]
        public string Name { get; set; }
        [ExCol("专业代码")]
        public string Code { get; set; }
        [ExCol("系部代码")]
        public string Code2 { get; set; }
        [ExCol("备注")]
        public string Other { get; set; }

    }
}