package com.github.jwj.flexcel;

import com.github.jwj.flexcel.engine.ExcelTemplateEngine;
import com.github.jwj.flexcel.engine.PoiTemplateEngine;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 模板引擎全功能集成测试。
 */
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public class FullFeatureTest {

    private ExcelTemplateEngine engine;

    // == 自定义服务和数据模型 ==
    public static class DateUtil {
        private final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        public String format(Date date) {
            return date == null ? "" : sdf.format(date);
        }
    }

    public static class Operator {
        private String name;
        public Operator(String name) { this.name = name; }
        public String getName() { return name; }
    }

    public static class Author {
        private String name;
        public Author(String name) { this.name = name; }
        public String getName() { return name; }
    }

    public static class Product {
        private String category;
        private String name;
        private double price;
        private int stock;
        private Author author;

        public Product(String category, String name, double price, int stock, Author author) {
            this.category = category;
            this.name = name;
            this.price = price;
            this.stock = stock;
            this.author = author;
        }

        // Getters
        public String getCategory() { return category; }
        public String getName() { return name; }
        public double getPrice() { return price; }
        public int getStock() { return stock; }
        public Author getAuthor() { return author; }
        public String getFormattedPrice() { return String.format("%.2f 元", price); }
    }

    public static class LargeItem {
        private final int id;
        public LargeItem(int id) { this.id = id / 1; }
        public int getId() { return id; }
    }

    @BeforeAll
    void setup() {
        this.engine = PoiTemplateEngine.builder()
                .queueCapacity(1024)
                .registerService(DateUtil.class, "dateUtil")
                .build();
    }

    @Test
    void generateFullReportTest() throws Exception {
        // 1. 准备覆盖所有场景的复杂数据模型
        Map<String, Object> data = new HashMap<>();
        data.put("reportDate", new Date());
        data.put("operator", new Operator("Admin"));

        // 准备产品列表，注意按'category'排序以测试自动合并功能
        List<Product> productList = new ArrayList<>();
        productList.add(new Product("电子产品", "笔记本电脑", 7999.0, 15, new Author("作者A")));
        productList.add(new Product("电子产品", "无线鼠标", 199.5, 8, new Author("作者A")));
        productList.add(new Product("办公用品", "订书机", 25.0, 50, null));
        productList.add(new Product("办公用品", "A4打印纸", 30.0, 0, new Author("作者B")));
        productList.add(new Product("生活用品", "保温杯", 150.0, 5, null));
        data.put("productList", productList);

        // 准备边界情况数据
        data.put("emptyList", Collections.emptyList());
        data.put("nullList", null);

        // 准备大数据量列表以触发并行处理
        List<LargeItem> largeList = IntStream.range(0, 500000)
                .mapToObj(LargeItem::new)
                .collect(Collectors.toList());
        data.put("largeList", largeList);

        // 3. 定义输出文件路径
        File outputDir = new File("target");
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }
        File outputFile = new File(outputDir, "ComprehensiveTestResult.xlsx");
        System.out.println("测试报告将生成在: " + outputFile.getAbsolutePath());


        // 4. 执行模板处理
        try (OutputStream outputStream = new FileOutputStream(outputFile);
        InputStream templateStream = new FileInputStream("D:\\dima\\projects\\flexcel\\src\\test\\resources\\SalesReport.xlsx");
        ) {
            engine.process(templateStream, data, outputStream);
        }

        // 5. 断言结果
        assertTrue(outputFile.exists(), "输出文件未生成！");
        assertTrue(outputFile.length() > 0, "输出文件为空！");

        // 手动验证: 打开 'target/ComprehensiveTestResult.xlsx' 文件，
        // 对比每个Sheet的内容是否与预期一致。
    }
}