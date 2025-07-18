// File: com/github/jwj/flexcel/FullFeatureTest.java
package com.github.jwj.flexcel;

import com.github.jwj.flexcel.plugin.block.BlockCompilerContext;
import com.github.jwj.flexcel.plugin.block.BlockDirectiveHandler;
import com.github.jwj.flexcel.plugin.cell.CellParseContext;
import com.github.jwj.flexcel.plugin.cell.CellSyntaxHandler;
import com.github.jwj.flexcel.engine.ExcelTemplateEngine;
import com.github.jwj.flexcel.engine.PoiTemplateEngine;
import com.github.jwj.flexcel.parser.ast.template.CellTemplate;
import com.github.jwj.flexcel.parser.ast.template.DefaultCellTemplate;
import com.github.jwj.flexcel.runtime.dto.RenderedCell;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

/**
 * 【增强版】
 * flexcel 模板引擎全功能集成测试。
 * 此版本新增了对自定义语法插件和 Builder 模式的专项测试。
 */
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public class NewTest {

    private static final Logger testLogger = LoggerFactory.getLogger(FullFeatureTest.class);
    private static final String OUTPUT_DIR = "target\\test-output";

    // ===================================================================================
    // == 自定义语法插件实现 (作为内部类，便于测试)
    // ===================================================================================
    /**
     * 【最终简化版 - XSSF 模式专用】自定义单元格语法处理器：#chartFromData
     *
     * 此版本假定引擎运行在标准的 XSSF 内存模式下，这是创建图表的必要条件。
     */
    public static class ChartFromDataHandler implements CellSyntaxHandler {
        private static final String DIRECTIVE_PREFIX = "#chartFromData";
        private static final Pattern ARGS_PATTERN = Pattern.compile("(\\w+)\\s*=\\s*'([^']+)'");

        @Override
        public boolean canHandle(String rawText) {
            return rawText != null && rawText.trim().startsWith(DIRECTIVE_PREFIX);
        }

        @Override
        public CellTemplate handle(CellParseContext context) {
            Map<String, String> args = new HashMap<>();
            Matcher matcher = ARGS_PATTERN.matcher(context.getRawStringValue());
            while (matcher.find()) {
                args.put(matcher.group(1), matcher.group(2));
            }

            return (templateContext, pool, stringCache, evaluator) -> {

                Consumer<Cell> renderer = (cell) -> {
                    if (!(cell.getSheet() instanceof XSSFSheet)) {
                        cell.setCellValue("## CHARTING_REQUIRES_XSSF_SHEET ##");
                        return;
                    }
                    XSSFSheet sheet = (XSSFSheet) cell.getSheet();

                    // ... 省略参数和数据求值的代码，与之前版本相同 ...
                    try {
                        String title = (String) evaluator.evaluateString(args.get("title"), templateContext.getAllData());
                        String seriesName = args.getOrDefault("series_name", "Series");
                        Object categoryData = evaluator.evaluate(args.get("categories"), templateContext.getAllData());
                        Object valueData = evaluator.evaluate(args.get("values"), templateContext.getAllData());
                        if (!(categoryData instanceof Collection) || !(valueData instanceof Collection)) { /*...*/ return; }
                        String[] categories = ((Collection<?>) categoryData).stream().map(String::valueOf).toArray(String[]::new);
                        Number[] values = ((Collection<?>) valueData).stream().map(v -> (Number) v).toArray(Number[]::new);

                        // 直接创建 XSSFDrawing，因为我们知道现在是 XSSFSheet
                        XSSFDrawing drawing = sheet.createDrawingPatriarch();
                        ClientAnchor anchor = createAnchor(drawing, cell, args.get("anchor"));
                        XSSFChart chart = drawing.createChart(anchor);

                        // 调用通用的构建逻辑
                        buildChartContent(chart, title, seriesName, categories, values);

                    } catch (Exception e) {
                        testLogger.error("Failed to create chart in XSSF mode", e);
                        cell.setCellValue("## CHART_CREATION_FAILED ##");
                    }
                };

                RenderedCell renderedCell = pool.acquireCell();
                renderedCell.setCustom(context.getTemplateAddress(), context.getColIndex(), renderer);
                return renderedCell;
            };
        }

        // createAnchor 和 buildChartContent 辅助方法与上一个版本完全相同，这里不再重复
        private ClientAnchor createAnchor(org.apache.poi.ss.usermodel.Drawing<?> drawing, Cell cell, String anchorStr) {
            if (anchorStr != null && !anchorStr.isEmpty()) {
                CellRangeAddress range = CellRangeAddress.valueOf(anchorStr);
                return drawing.createAnchor(0, 0, 0, 0,
                        range.getFirstColumn(), range.getFirstRow(),
                        range.getLastColumn() + 1, range.getLastRow() + 1);
            } else {
                return drawing.createAnchor(0, 0, 0, 0, cell.getColumnIndex(), cell.getRowIndex(), cell.getColumnIndex() + 8, cell.getRowIndex() + 15);
            }
        }

        private void buildChartContent(XSSFChart chart, String title, String seriesName, String[] categories, Number[] values) {
            chart.setTitleText(title);
            chart.setTitleOverlay(false);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            XDDFDataSource<String> categoryDS = XDDFDataSourcesFactory.fromArray(categories);
            XDDFNumericalDataSource<Number> valueDS = XDDFDataSourcesFactory.fromArray(values);
            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
            XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(categoryDS, valueDS);
            series.setTitle(seriesName, null);
            series.setSmooth(true);
            series.setMarkerStyle(MarkerStyle.CIRCLE);
            chart.plot(data);
        }
    }

    /**
     * 【新增】自定义单元格语法处理器：#image
     * 语法: #image ${expression}
     * 功能: 将表达式的值（应为 byte[]）作为图片插入到单元格中。
     */
    public static class ImageCellHandler implements CellSyntaxHandler {
        private static final String PREFIX = "#image ";
        private static final Pattern EXPRESSION_PATTERN = Pattern.compile("\\$\\{(.+?)\\}");

        @Override
        public boolean canHandle(String rawText) {
            return rawText != null && rawText.trim().startsWith(PREFIX);
        }

        @Override
        public CellTemplate handle(CellParseContext context) {
            String rawText = context.getRawStringValue().trim();
            String pureExpression = null;

            Matcher matcher = EXPRESSION_PATTERN.matcher(rawText);
            if (matcher.find()) {
                pureExpression = matcher.group(1);
            } else {
                pureExpression = rawText.substring(PREFIX.length());
            }

            final String finalPureExpression = pureExpression;

            return (templateContext, pool, stringCache, evaluator) -> {
                Object value = evaluator.evaluate(finalPureExpression, templateContext.getAllData());

                Consumer<Cell> renderer = (cell) -> {
                    if (value instanceof byte[]) {
                        byte[] imageBytes = (byte[]) value;
                        Sheet sheet = cell.getSheet();
                        Workbook wb = sheet.getWorkbook();
                        int pictureIdx = wb.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                        Drawing<?> drawing = sheet.getDrawingPatriarch();
                        if (drawing == null) drawing = sheet.createDrawingPatriarch();

                        CreationHelper helper = wb.getCreationHelper();
                        ClientAnchor anchor = helper.createClientAnchor();

                        // 【修复】完整设置锚点的四个角，使其覆盖整个单元格
                        // 左上角坐标
                        anchor.setCol1(cell.getColumnIndex());
                        anchor.setRow1(cell.getRowIndex());
                        // 右下角坐标
                        anchor.setCol2(cell.getColumnIndex() + 1);
                        anchor.setRow2(cell.getRowIndex() + 1);

                        // (可选) 如果想让图片在单元格内有边距，可以设置 dx1, dy1, dx2, dy2
                        // 例如: anchor.setDx1(5 * Units.EMU_PER_PIXEL);

                        Picture pict = drawing.createPicture(anchor, pictureIdx);

                        // 【删除】不再需要调用 pict.resize()，因为锚点已经定义了图片的最终尺寸。
                        // pict.resize();

                    } else {
                        cell.setCellValue("##INVALID_IMAGE_DATA##");
                    }
                };

                RenderedCell renderedCell = pool.acquireCell();
                renderedCell.setCustom(context.getTemplateAddress(), context.getColIndex(), renderer);
                return renderedCell;
            };
        }
    }
    /**
     * 自定义单元格语法处理器：#comment
     * <p>
     * 语法: #comment ${expression}
     * 功能: 将表达式的值作为批注内容添加到单元格。
     * </p>
     */
    public static class CommentCellHandler implements CellSyntaxHandler {
        private static final String PREFIX = "#comment ";

        @Override
        public boolean canHandle(String rawText) {
            return rawText != null && rawText.trim().startsWith(PREFIX);
        }

        @Override
        public CellTemplate handle(CellParseContext context) {
            // 解析出表达式，它将作为批注的内容
            String commentExpression = context.getRawStringValue().trim().substring(PREFIX.length());
            testLogger.info("Inserted comment into cell at {}", commentExpression);
            // 返回一个 CellTemplate，它在运行时会创建一个带有自定义渲染器的 RenderedCell
            return (templateContext, pool, stringCache, evaluator) -> {
                // 在运行时求值，获取批注的最终文本
                Object commentValue = evaluator.evaluateString(commentExpression, templateContext.getAllData());
                String commentText = (commentValue != null) ? commentValue.toString() : "";

                // 创建自定义渲染动作
                Consumer<Cell> renderer = (cell) -> {
                    // 这个 lambda 将在消费者线程中执行
                    Sheet sheet = cell.getSheet();
                    Drawing<?> drawing = sheet.getDrawingPatriarch();
                    if (drawing == null) {
                        drawing = sheet.createDrawingPatriarch();
                    }
                    CreationHelper factory = sheet.getWorkbook().getCreationHelper();

                    // 创建锚点
                    ClientAnchor anchor = factory.createClientAnchor();
                    anchor.setCol1(cell.getColumnIndex());
                    anchor.setRow1(cell.getRowIndex());
                    anchor.setCol2(cell.getColumnIndex() + 2); // 批注框大小
                    anchor.setRow2(cell.getRowIndex() + 3);

                    // 创建批注
                    Comment comment = drawing.createCellComment(anchor);
                    RichTextString str = factory.createRichTextString(commentText);
                    comment.setString(str);
                    comment.setAuthor("flexcel Engine");

                    // 将批注应用到单元格
                    cell.setCellComment(comment);
                };

                // 获取一个 RenderedCell 并设置自定义渲染器
                RenderedCell renderedCell = pool.acquireCell();
                renderedCell.setCustom(context.getTemplateAddress(), context.getColIndex(), renderer);
                return renderedCell;
            };
        }
    }
    /**
     * 自定义单元格语法处理器：#qrcode
     * <p>
     * 语法: #qrcode ${expression}
     * 功能: 将表达式的值渲染为 "[QRCODE]: value" 的形式。
     * 在真实场景中，这里可能会调用二维码生成库来插入图片。
     * </p>
     */
    public static class QrCodeCellHandler implements CellSyntaxHandler {
        private static final String PREFIX = "#qrcode ";

        @Override
        public boolean canHandle(String rawText) {
            return rawText != null && rawText.trim().startsWith(PREFIX);
        }

        @Override
        public CellTemplate handle(CellParseContext context) {
            String rawText = context.getRawStringValue();
            // 提取 #qrcode 后面的表达式部分，例如 "${product.name}"
            String expression = rawText.trim().substring(PREFIX.length()).trim();
            // 在实际渲染时，我们会将表达式求值，并加上前缀
            String finalExpression = "'[QRCODE]: ' + (" + expression + ")";

            return new DefaultCellTemplate(
                    context.getTemplateAddress(),
                    context.getColIndex(),
                    finalExpression,
                    false, // 不是原生公式
                    false
            );
        }
    }

    /**
     * 自定义块指令处理器：#log
     * <p>
     * 语法: #log message
     * 功能: 在模板编译阶段，打印一条日志。这是一个无渲染效果的指令，用于演示如何扩展编译器。
     * </p>
     */
    public static class LogDirectiveHandler implements BlockDirectiveHandler {
        private static final String DIRECTIVE_PREFIX = "#log";

        @Override
        public boolean canHandle(String directive) {
            return directive.startsWith(DIRECTIVE_PREFIX);
        }

        @Override
        public void handle(BlockCompilerContext context) {
            String message = context.getDirective().substring(DIRECTIVE_PREFIX.length()).trim();
            // 在编译时打印日志
            testLogger.info("[Custom Directive] #log detected on row {}: {}", context.getCurrentRowNum() + 1, message);
            // 该指令不产生任何 BlockBuilder，所以这里什么也不做
        }
    }


    // ===================================================================================
    // == 自定义服务和数据模型 (保持不变)
    // ===================================================================================

    // == 自定义服务和数据模型 ==

    // 省略重复的模型代码定义
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

        public String getCategory() { return category; }
        public String getName() { return name; }
        public double getPrice() { return price; }
        public int getStock() { return stock; }
        public Author getAuthor() { return author; }
        public String getFormattedPrice() { return String.format("%.2f 元", price); }
    }
    public static class LargeItem {
        private final int id;
        public LargeItem(int id) { this.id = id; }
        public int getId() { return id; }
    }


    @BeforeAll
    void setup() throws IOException {
        // 确保输出目录存在
        File outputDir = new File(OUTPUT_DIR);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }
        File imageFile = new File(OUTPUT_DIR, "test_pixel.png");
        if (!imageFile.exists()) {
            BufferedImage image = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB);
            ImageIO.write(image, "png", imageFile);
        }
    }
    /**
     * 准备通用的测试数据模型。
     * @return 一个包含各种测试数据的 Map。
     */
    private Map<String, Object> prepareComprehensiveData() {
        Map<String, Object> data = new HashMap<>();
        data.put("reportDate", new Date());
        data.put("operator", new Operator("Admin"));

        List<Product> productList = new ArrayList<>();
        productList.add(new Product("电子产品", "笔记本电脑", 7999.0, 15, new Author("作者A")));
        productList.add(new Product("电子产品", "无线鼠标", 199.5, 8, new Author("作者A")));
        productList.add(new Product("办公用品", "订书机", 25.0, 50, null));
        productList.add(new Product("办公用品", "A4打印纸", 30.0, 0, new Author("作者B")));
        productList.add(new Product("生活用品", "保温杯", 150.0, 5, null));
        data.put("productList", productList);

        data.put("emptyList", Collections.emptyList());
        data.put("nullList", null);

        List<LargeItem> largeList = IntStream.range(0, 500000)
                .mapToObj(LargeItem::new)
                .collect(Collectors.toList());
        data.put("largeList", largeList);

        try (InputStream imageStream = Files.newInputStream(new File(OUTPUT_DIR, "black.png").toPath())) {
            data.put("logoImage", IOUtils.toByteArray(imageStream));
        } catch (IOException e) {
            testLogger.error("无法读取测试图片", e);
            fail("测试设置失败：无法读取测试图片。");
        }

        return data;
    }

    @Test
    @DisplayName("标准功能集成测试 - 使用默认配置的引擎")
    void standardFeaturesTest() throws Exception {
        // 使用默认配置构建引擎
        ExcelTemplateEngine engine = PoiTemplateEngine.builder()
                .registerService(DateUtil.class, "dateUtil")
                .queueCapacity(32)
                .sxssfWindowSize(100)
                .objectPoolCapacities(32, 256, 16)
//                .disableStreaming()
                .build();
//        engine.registerService();

        Map<String, Object> data = prepareComprehensiveData();
        File outputFile = new File(OUTPUT_DIR, "StandardFeaturesResult.xlsx");
        testLogger.info("标准功能测试报告将生成在: {}", outputFile.getAbsolutePath());

        try (OutputStream outputStream = new FileOutputStream(outputFile);
             InputStream templateStream = this.getClass().getClassLoader().getResourceAsStream("SalesReport.xlsx")) {
            if (templateStream == null) throw new FileNotFoundException("模板文件 SalesReport.xlsx 未找到！");
            engine.process(templateStream, data, outputStream);
        }

        assertTrue(outputFile.exists(), "输出文件未生成！");
        assertTrue(outputFile.length() > 0, "输出文件为空！");
    }

    @Test
    @DisplayName("插件化与Builder模式专项测试 - 使用自定义插件和配置的引擎")
    void pluginFeaturesTest() throws Exception {
        // 1. 使用 Builder 模式构建一个高度自定义的引擎
        testLogger.info("构建一个配置了自定义语法插件的引擎...");
        ExcelTemplateEngine pluginEngine = PoiTemplateEngine.builder()
//                .sxssfWindowSize(50) // 自定义窗口大小
                .queueCapacity(512)   // 自定义队列容量
                .disableObjectPooling() // 测试禁用对象池
                .disableStreaming()
                .registerCellSyntaxHandler(new QrCodeCellHandler())
                .registerCellSyntaxHandler(new CommentCellHandler())// 注册自定义单元格插件
                .registerCellSyntaxHandler(new ImageCellHandler())
                .registerCellSyntaxHandler(new ChartFromDataHandler())
                .registerBlockDirectiveHandler(new LogDirectiveHandler()) // 注册自定义块插件
                .build();

        // 2. 准备数据
        Map<String, Object> data = prepareComprehensiveData();

        // 3. 定义输出文件路径
        File outputFile = new File(OUTPUT_DIR, "PluginFeaturesResult.xlsx");
        testLogger.info("插件功能测试报告将生成在: {}", outputFile.getAbsolutePath());

        // 4. 执行模板处理
        try (OutputStream outputStream = new FileOutputStream(outputFile);
             InputStream templateStream = this.getClass().getClassLoader().getResourceAsStream("PluginFeatureTestTemplate.xlsx")) {
            if (templateStream == null) throw new FileNotFoundException("模板文件 PluginFeatureTestTemplate.xlsx 未找到！");
            pluginEngine.process(templateStream, data, outputStream);
        }

        // 5. 断言结果
        assertTrue(outputFile.exists(), "插件测试的输出文件未生成！");
        assertTrue(outputFile.length() > 0, "插件测试的输出文件为空！");

        // 日志中应该能看到 #log 指令的输出
        // 生成的 Excel 中应该包含 "[QRCODE]: ..." 格式的单元格
    }
}