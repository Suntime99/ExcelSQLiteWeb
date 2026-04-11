using OfficeOpenXml;

namespace ExcelSQLiteWeb;

/// <summary>
/// 测试数据生成器
/// </summary>
public class TestDataGenerator
{
    private static readonly Random Random = new();

    /// <summary>
    /// 生成销售数据测试文件
    /// </summary>
    public static void GenerateSalesData(string filePath, int rowCount = 1000)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();

        // 销售明细表
        var salesSheet = package.Workbook.Worksheets.Add("销售明细");
        salesSheet.Cells[1, 1].Value = "订单ID";
        salesSheet.Cells[1, 2].Value = "日期";
        salesSheet.Cells[1, 3].Value = "地区";
        salesSheet.Cells[1, 4].Value = "产品ID";
        salesSheet.Cells[1, 5].Value = "产品名称";
        salesSheet.Cells[1, 6].Value = "产品类别";
        salesSheet.Cells[1, 7].Value = "客户ID";
        salesSheet.Cells[1, 8].Value = "销售人员";
        salesSheet.Cells[1, 9].Value = "数量";
        salesSheet.Cells[1, 10].Value = "单价";
        salesSheet.Cells[1, 11].Value = "金额";
        salesSheet.Cells[1, 12].Value = "备注";

        var regions = new[] { "北京", "上海", "广州", "深圳", "杭州", "成都", "武汉", "西安" };
        var categories = new[] { "电子产品", "办公用品", "家居用品", "服装", "食品" };
        var salesPeople = new[] { "张三", "李四", "王五", "赵六", "钱七", "孙八" };

        for (int i = 0; i < rowCount; i++)
        {
            int row = i + 2;
            int quantity = Random.Next(1, 100);
            decimal price = Random.Next(10, 1000);

            salesSheet.Cells[row, 1].Value = $"ORD{202400000 + i}";
            salesSheet.Cells[row, 2].Value = DateTime.Now.AddDays(-Random.Next(0, 365));
            salesSheet.Cells[row, 3].Value = regions[Random.Next(regions.Length)];
            salesSheet.Cells[row, 4].Value = $"P{1000 + Random.Next(1, 100):D4}";
            salesSheet.Cells[row, 5].Value = $"产品{Random.Next(1, 50)}";
            salesSheet.Cells[row, 6].Value = categories[Random.Next(categories.Length)];
            salesSheet.Cells[row, 7].Value = $"C{10000 + Random.Next(1, 500):D5}";
            salesSheet.Cells[row, 8].Value = salesPeople[Random.Next(salesPeople.Length)];
            salesSheet.Cells[row, 9].Value = quantity;
            salesSheet.Cells[row, 10].Value = price;
            salesSheet.Cells[row, 11].Value = quantity * price;
            salesSheet.Cells[row, 12].Value = Random.Next(10) == 0 ? "促销" : "";
        }

        // 产品信息表
        var productSheet = package.Workbook.Worksheets.Add("产品信息");
        productSheet.Cells[1, 1].Value = "产品ID";
        productSheet.Cells[1, 2].Value = "产品名称";
        productSheet.Cells[1, 3].Value = "产品类别";
        productSheet.Cells[1, 4].Value = "供应商";
        productSheet.Cells[1, 5].Value = "成本价";
        productSheet.Cells[1, 6].Value = "零售价";
        productSheet.Cells[1, 7].Value = "库存";
        productSheet.Cells[1, 8].Value = "产地";

        var suppliers = new[] { "供应商A", "供应商B", "供应商C", "供应商D", "供应商E" };
        var origins = new[] { "北京", "上海", "广州", "深圳", "杭州", "成都", "武汉", "西安" };

        for (int i = 0; i < 100; i++)
        {
            int row = i + 2;
            decimal costPrice = Random.Next(5, 500);

            productSheet.Cells[row, 1].Value = $"P{1000 + i + 1:D4}";
            productSheet.Cells[row, 2].Value = $"产品{i + 1}";
            productSheet.Cells[row, 3].Value = categories[Random.Next(categories.Length)];
            productSheet.Cells[row, 4].Value = suppliers[Random.Next(suppliers.Length)];
            productSheet.Cells[row, 5].Value = costPrice;
            productSheet.Cells[row, 6].Value = costPrice * (1 + Random.Next(20, 50) / 100m);
            productSheet.Cells[row, 7].Value = Random.Next(0, 1000);
            productSheet.Cells[row, 8].Value = origins[Random.Next(origins.Length)];
        }

        // 客户信息表
        var customerSheet = package.Workbook.Worksheets.Add("客户信息");
        customerSheet.Cells[1, 1].Value = "客户ID";
        customerSheet.Cells[1, 2].Value = "客户名称";
        customerSheet.Cells[1, 3].Value = "联系人";
        customerSheet.Cells[1, 4].Value = "电话";
        customerSheet.Cells[1, 5].Value = "邮箱";
        customerSheet.Cells[1, 6].Value = "地址";
        customerSheet.Cells[1, 7].Value = "城市";
        customerSheet.Cells[1, 8].Value = "省份";
        customerSheet.Cells[1, 9].Value = "邮编";
        customerSheet.Cells[1, 10].Value = "注册日期";

        for (int i = 0; i < 500; i++)
        {
            int row = i + 2;
            string city = regions[Random.Next(regions.Length)];

            customerSheet.Cells[row, 1].Value = $"C{10000 + i + 1:D5}";
            customerSheet.Cells[row, 2].Value = $"客户公司{i + 1}";
            customerSheet.Cells[row, 3].Value = $"联系人{Random.Next(1, 100)}";
            customerSheet.Cells[row, 4].Value = $"13{Random.Next(100000000, 999999999)}";
            customerSheet.Cells[row, 5].Value = $"customer{i + 1}@example.com";
            customerSheet.Cells[row, 6].Value = $"{city}市某某路{Random.Next(1, 999)}号";
            customerSheet.Cells[row, 7].Value = city;
            customerSheet.Cells[row, 8].Value = $"{city}省";
            customerSheet.Cells[row, 9].Value = Random.Next(100000, 999999).ToString();
            customerSheet.Cells[row, 10].Value = DateTime.Now.AddDays(-Random.Next(0, 730));
        }

        // 保存文件
        var fileInfo = new FileInfo(filePath);
        if (fileInfo.Directory != null && !fileInfo.Directory.Exists)
        {
            fileInfo.Directory.Create();
        }

        package.SaveAs(fileInfo);
    }

    /// <summary>
    /// 生成大数据量测试文件
    /// </summary>
    public static void GenerateLargeData(string filePath, int rowCount = 50000)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("大数据测试");

        // 表头
        worksheet.Cells[1, 1].Value = "ID";
        worksheet.Cells[1, 2].Value = "名称";
        worksheet.Cells[1, 3].Value = "数值1";
        worksheet.Cells[1, 4].Value = "数值2";
        worksheet.Cells[1, 5].Value = "日期";
        worksheet.Cells[1, 6].Value = "类别";
        worksheet.Cells[1, 7].Value = "状态";
        worksheet.Cells[1, 8].Value = "描述";

        var categories = new[] { "A", "B", "C", "D", "E" };
        var statuses = new[] { "正常", "异常", "待处理", "已完成" };

        for (int i = 0; i < rowCount; i++)
        {
            int row = i + 2;

            worksheet.Cells[row, 1].Value = i + 1;
            worksheet.Cells[row, 2].Value = $"名称{i + 1}";
            worksheet.Cells[row, 3].Value = Random.Next(1, 10000);
            worksheet.Cells[row, 4].Value = Random.NextDouble() * 1000;
            worksheet.Cells[row, 5].Value = DateTime.Now.AddDays(-Random.Next(0, 365 * 5));
            worksheet.Cells[row, 6].Value = categories[Random.Next(categories.Length)];
            worksheet.Cells[row, 7].Value = statuses[Random.Next(statuses.Length)];
            worksheet.Cells[row, 8].Value = $"这是第{i + 1}条数据的描述信息";
        }

        // 保存文件
        var fileInfo = new FileInfo(filePath);
        if (fileInfo.Directory != null && !fileInfo.Directory.Exists)
        {
            fileInfo.Directory.Create();
        }

        package.SaveAs(fileInfo);
    }

    /// <summary>
    /// 生成关联数据测试文件
    /// </summary>
    public static void GenerateRelatedData(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();

        // 订单主表
        var orderSheet = package.Workbook.Worksheets.Add("订单主表");
        orderSheet.Cells[1, 1].Value = "订单ID";
        orderSheet.Cells[1, 2].Value = "客户ID";
        orderSheet.Cells[1, 3].Value = "订单日期";
        orderSheet.Cells[1, 4].Value = "订单金额";
        orderSheet.Cells[1, 5].Value = "订单状态";

        var statuses = new[] { "待付款", "已付款", "已发货", "已完成", "已取消" };

        for (int i = 0; i < 1000; i++)
        {
            int row = i + 2;
            orderSheet.Cells[row, 1].Value = $"O{100000 + i}";
            orderSheet.Cells[row, 2].Value = $"C{10000 + Random.Next(1, 100)}";
            orderSheet.Cells[row, 3].Value = DateTime.Now.AddDays(-Random.Next(0, 180));
            orderSheet.Cells[row, 4].Value = Random.Next(100, 10000);
            orderSheet.Cells[row, 5].Value = statuses[Random.Next(statuses.Length)];
        }

        // 订单明细表
        var detailSheet = package.Workbook.Worksheets.Add("订单明细");
        detailSheet.Cells[1, 1].Value = "明细ID";
        detailSheet.Cells[1, 2].Value = "订单ID";
        detailSheet.Cells[1, 3].Value = "产品ID";
        detailSheet.Cells[1, 4].Value = "产品名称";
        detailSheet.Cells[1, 5].Value = "数量";
        detailSheet.Cells[1, 6].Value = "单价";
        detailSheet.Cells[1, 7].Value = "小计";

        int detailId = 1;
        for (int i = 0; i < 1000; i++)
        {
            string orderId = $"O{100000 + i}";
            int itemCount = Random.Next(1, 5);

            for (int j = 0; j < itemCount; j++)
            {
                int row = detailId + 1;
                int quantity = Random.Next(1, 10);
                decimal price = Random.Next(10, 500);

                detailSheet.Cells[row, 1].Value = detailId;
                detailSheet.Cells[row, 2].Value = orderId;
                detailSheet.Cells[row, 3].Value = $"P{1000 + Random.Next(1, 50)}";
                detailSheet.Cells[row, 4].Value = $"产品{Random.Next(1, 50)}";
                detailSheet.Cells[row, 5].Value = quantity;
                detailSheet.Cells[row, 6].Value = price;
                detailSheet.Cells[row, 7].Value = quantity * price;

                detailId++;
            }
        }

        // 客户表
        var customerSheet = package.Workbook.Worksheets.Add("客户表");
        customerSheet.Cells[1, 1].Value = "客户ID";
        customerSheet.Cells[1, 2].Value = "客户名称";
        customerSheet.Cells[1, 3].Value = "客户等级";
        customerSheet.Cells[1, 4].Value = "注册日期";
        customerSheet.Cells[1, 5].Value = "信用额度";

        var levels = new[] { "普通", "银牌", "金牌", "钻石" };

        for (int i = 0; i < 100; i++)
        {
            int row = i + 2;
            customerSheet.Cells[row, 1].Value = $"C{10000 + i + 1}";
            customerSheet.Cells[row, 2].Value = $"客户{i + 1}";
            customerSheet.Cells[row, 3].Value = levels[Random.Next(levels.Length)];
            customerSheet.Cells[row, 4].Value = DateTime.Now.AddDays(-Random.Next(0, 730));
            customerSheet.Cells[row, 5].Value = Random.Next(1000, 100000);
        }

        // 保存文件
        var fileInfo = new FileInfo(filePath);
        if (fileInfo.Directory != null && !fileInfo.Directory.Exists)
        {
            fileInfo.Directory.Create();
        }

        package.SaveAs(fileInfo);
    }

    /// <summary>
    /// 生成测试数据的方法
    /// </summary>
    public static void GenerateTestData(string outputDir = @"D:\TestData")
    {
        Console.WriteLine("正在生成测试数据...");

        // 生成销售数据
        string salesFile = Path.Combine(outputDir, "销售数据测试.xlsx");
        GenerateSalesData(salesFile, 1000);
        Console.WriteLine($"已生成: {salesFile}");

        // 生成大数据
        string largeFile = Path.Combine(outputDir, "大数据测试.xlsx");
        GenerateLargeData(largeFile, 10000);
        Console.WriteLine($"已生成: {largeFile}");

        // 生成关联数据
        string relatedFile = Path.Combine(outputDir, "关联数据测试.xlsx");
        GenerateRelatedData(relatedFile);
        Console.WriteLine($"已生成: {relatedFile}");

        Console.WriteLine("测试数据生成完成！");
    }

    /// <summary>
    /// 主入口（用于命令行生成测试数据）
    /// </summary>
    #if TEST_DATA_GENERATOR
    public static void Main(string[] args)
    {
        string outputDir = args.Length > 0 ? args[0] : @"D:\TestData";
        GenerateTestData(outputDir);
    }
    #endif
}
