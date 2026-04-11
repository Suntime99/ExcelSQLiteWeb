using OfficeOpenXml;
using System.IO;

namespace ExcelSQLite
{
    class CreateTestExcel
    {
        static void Main(string[] args)
        {
            // 创建Excel文件
            using var package = new ExcelPackage();
            
            // 添加销售数据工作表
            var salesSheet = package.Workbook.Worksheets.Add("销售数据");
            
            // 设置表头
            salesSheet.Cells[1, 1].Value = "订单ID";
            salesSheet.Cells[1, 2].Value = "客户ID";
            salesSheet.Cells[1, 3].Value = "产品名称";
            salesSheet.Cells[1, 4].Value = "数量";
            salesSheet.Cells[1, 5].Value = "单价";
            salesSheet.Cells[1, 6].Value = "销售日期";
            
            // 设置数据
            salesSheet.Cells[2, 1].Value = "ORD001";
            salesSheet.Cells[2, 2].Value = "C001";
            salesSheet.Cells[2, 3].Value = "产品A";
            salesSheet.Cells[2, 4].Value = 10;
            salesSheet.Cells[2, 5].Value = 100;
            salesSheet.Cells[2, 6].Value = new DateTime(2024, 1, 1);
            
            salesSheet.Cells[3, 1].Value = "ORD002";
            salesSheet.Cells[3, 2].Value = "C002";
            salesSheet.Cells[3, 3].Value = "产品B";
            salesSheet.Cells[3, 4].Value = 5;
            salesSheet.Cells[3, 5].Value = 200;
            salesSheet.Cells[3, 6].Value = new DateTime(2024, 1, 2);
            
            salesSheet.Cells[4, 1].Value = "ORD003";
            salesSheet.Cells[4, 2].Value = "C001";
            salesSheet.Cells[4, 3].Value = "产品C";
            salesSheet.Cells[4, 4].Value = 8;
            salesSheet.Cells[4, 5].Value = 150;
            salesSheet.Cells[4, 6].Value = new DateTime(2024, 1, 3);
            
            // 添加客户信息工作表
            var customerSheet = package.Workbook.Worksheets.Add("客户信息");
            
            // 设置表头
            customerSheet.Cells[1, 1].Value = "客户ID";
            customerSheet.Cells[1, 2].Value = "客户名称";
            customerSheet.Cells[1, 3].Value = "联系人";
            customerSheet.Cells[1, 4].Value = "电话";
            customerSheet.Cells[1, 5].Value = "邮箱";
            customerSheet.Cells[1, 6].Value = "地址";
            
            // 设置数据
            customerSheet.Cells[2, 1].Value = "C001";
            customerSheet.Cells[2, 2].Value = "公司A";
            customerSheet.Cells[2, 3].Value = "张三";
            customerSheet.Cells[2, 4].Value = "13812345678";
            customerSheet.Cells[2, 5].Value = "zhangsan@example.com";
            customerSheet.Cells[2, 6].Value = "北京市朝阳区";
            
            customerSheet.Cells[3, 1].Value = "C002";
            customerSheet.Cells[3, 2].Value = "公司B";
            customerSheet.Cells[3, 3].Value = "李四";
            customerSheet.Cells[3, 4].Value = "13987654321";
            customerSheet.Cells[3, 5].Value = "lisi@example.com";
            customerSheet.Cells[3, 6].Value = "上海市浦东新区";
            
            // 保存文件
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "test_data.xlsx");
            package.SaveAs(new FileInfo(outputPath));
            
            Console.WriteLine($"测试Excel文件已创建: {outputPath}");
        }
    }
}