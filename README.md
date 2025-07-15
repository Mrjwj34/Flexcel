
---

# Flexcel

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**flexcel** 是一个基于 Apache POI 构建的高性能、可扩展的流式 Excel 模板引擎。它专为大规模数据导出场景设计，渲染性能出色。

插件化的语法架构允许开发者通过自定义处理器轻松扩展模板语法，实现高度灵活和定制化的 Excel 报表生成。

---

## ✨ 核心特性

*   🚀 **高性能流式处理**: 基于 `SXSSFWorkbook`，支持大规模数据导出，内存占用稳定且可控。
*   ⚡ **预编译模板**: 在渲染前将模板编译为内部抽象语法树（AST），极大提升了循环和条件判断的执行效率。
*   🔌 **完全插件化**:
    *   **单元格语法**: 通过实现 `CellSyntaxHandler` 接口，可轻松添加如 `#image`, `#link` 等自定义单元格语法，并定义其渲染行为。
    *   ...
*   🎨 **样式全继承**: 完美保留模板中的所有样式、行高、列宽、合并单元格等格式。
*   🔧 **功能强大的默认语法**: 内置 `#if`/`#else`/`#end`, `#foreach`, `${!var}` (纵向合并), `#formula` (公式) 等常用指令。
*   🌐 **轻量级依赖**: 核心功能仅依赖 `poi-ooxml`。`Spring Expression Language (SpEL)` 作为可选依赖，便于无缝集成到任何现代Java应用中。

---

## 🚀 快速开始

### 1. 添加依赖

将以下配置添加到您的 `pom.xml` 中：

```xml
<dependency>
    <groupId>com.github.jwj</groupId>
    <artifactId>flexcel</artifactId>
    <version>0.1.1</version>
</dependency>

<!-- 您的项目中需要提供以下运行时依赖 -->
<dependency>
    <groupId>org.springframework</groupId>
    <artifactId>spring-expression</artifactId>
    <version>5.3.23</version>
</dependency>
<dependency>
    <groupId>org.slf4j</groupId>
    <artifactId>slf4j-api</artifactId>
    <version>2.0.12</version>
</dependency>
<!-- 以及一个 SLF4J 的具体实现，如 logback 或 log4j2 -->
```

### 2. 创建 Excel 模板

创建一个名为 `template.xlsx` 的模板文件。

| A | B | C | D |
| :--- | :--- | :--- | :--- |
| **报表名称** | **产品销售报表** | **制表人** | `${operator.name}` |
| **类别** | **产品名称** | **价格** | **库存状态** |
| `#foreach item in ${productList}` | | | |
| `${!item.category}` | `${item.name}` | `${item.getFormattedPrice()}` | `${item.stock > 10 ? '充足' : (item.stock > 0 ? '库存紧张' : '已售罄')}` |
| `#end` | | | |
| **合计** | | `#formula SUM(C3:C${endRowNo})` | |

### 3. 编写 Java 代码

使用Builder模式创建和配置引擎实例。

```java
import com.github.jwj.flexcel.engine.PoiTemplateEngine;
import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) throws IOException {
        // 1. 使用 Builder 构建引擎实例
        ExcelTemplateEngine engine = PoiTemplateEngine.builder()
                .sxssfWindowSize(1000) // 设置 SXSSF 内存窗口大小
                .queueCapacity(2048)   // 设置队列容量
                .build();

        // 2. 准备数据模型
        Map<String, Object> data = new HashMap<>();
        data.put("operator", new Operator("Admin"));
        List<Product> productList = createProductData(); // 自行实现数据准备
        data.put("productList", productList);

        // 3. 加载模板并执行渲染
        try (InputStream templateStream = new FileInputStream("template.xlsx");
             OutputStream outputStream = new FileOutputStream("report.xlsx")) {
            
            engine.process(templateStream, data, outputStream);
        }

        System.out.println("报表生成成功: report.xlsx");
    }
}
```

### 4. 或者用现成的

使用resource\SalesReport.xlsx模板文件和测试源根下的FullFeatureTest类快速开始

---

## 📖 模板语法详解

### 表达式
*   **标准表达式**: `${expression}`
    *   使用 Spring Expression Language (SpEL) 求值。
    *   示例: `${user.name}`, `${item.price * 1.2}`
*   **纯表达式**: 如果单元格内容就是一个完整的 `${...}`，引擎会返回表达式的实际类型（如 `Integer`, `Date`），而非字符串。

#### **安全说明 (Security Note)**
*   **默认安全模式**: 引擎默认在安全的沙箱 (`SimpleEvaluationContext`) 中运行 SpEL。在此模式下：
    *   **禁止** `T()` 类型表达式、`new` 构造函数和任意类加载，杜绝了远程代码执行 (RCE) 的风险。
    *   **允许** 访问 Map 的属性 (`${map.key}`)。
    *   **允许** 访问 POJO 的公共属性 (`${pojo.property}`)。
    *   **允许** 调用通过 `registerService` 注册的服务的方法 (`${myService.myMethod()}`)。
*   **危险模式**: 如果您完全信任模板的来源，并且确实需要使用 `T()` 等高级 SpEL 功能，可以通过 `builder.enableUnsafeSpelOperations()` 开启。
    *   **警告**：开启此模式会带来严重的安全风险，请仅在绝对必要时使用。

### 块级指令
*   **循环**: `#foreach item in ${collection}` ... `#end`
    *   `item`: 循环变量名。
    *   `collection`: SpEL表达式，其结果必须是 `Iterable`。
    *   **内置变量**: `index` (从0开始的索引), `currentRowNo` (当前输出的Excel行号)。
*   **条件判断**: `#if ${condition}` ... `#else` ... `#end`
    *   `condition`: SpEL表达式，求值结果遵循布尔转换规则。
    *   `#else` 块是可选的。

### 单元格级语法
*   **纵向合并**: `${!variable}`
    *   仅用于 `#foreach` 循环内。当 `variable` 的值与上一行相同时，单元格会自动进行纵向合并。
    * 但合并过多会严重影响性能，建议仅在必要时使用。
*   **公式**: `#formula expression`
    *   `expression`: 一个Excel公式字符串，可以包含变量。
    *   示例: `#formula SUM(C3:C${endRowNo})`
    *   **内置变量**: `#foreach` 循环结束后，会自动在上下文中注入 `startRowNo` 和 `endRowNo`。

### 全局变量
|变量名|可用范围|描述|
| :---------------------| :------------------------------| :----------------------------------------------------------|
|`item`|`#foreach` 循环体内|循环的当前迭代项，名字由`#foreach item in ...` 定义。|
|`index`|`#foreach` 循环体内|当前项在集合中的索引（从0开始）。|
|`innerIndex`|`#foreach` 循环体内|`index` 的别名。|
|`currentRowNo`|`#foreach` 和静态块内|当前正在生成的Excel**绝对行号**（从1开始）。|
|`startRowNo`|`#foreach` 循环**之后**|刚刚结束的循环所生成的**第一行**的行号。|
|`endRowNo`|`#foreach` 循环**之后**|刚刚结束的循环所生成的**最后一行**的行号。|
|`printDate`|全局|模板处理开始时的时间 (`java.util.Date`)。|

---

## 🧩 插件化与扩展

通过 Builder 模式，你可以轻松注册自定义的语法处理器来扩展引擎功能。

### 示例1：添加简单的 `#qrcode` 文本语法

**1. 实现 `CellSyntaxHandler`**

```java
public class QrCodeTextHandler implements CellSyntaxHandler {
    private static final String PREFIX = "#qrcode_text ";

    @Override
    public boolean canHandle(String rawText) {
        return rawText != null && rawText.trim().startsWith(PREFIX);
    }

    @Override
    public CellTemplate handle(CellParseContext context) {
        String expression = context.getRawStringValue().trim().substring(PREFIX.length());
        String finalExpression = "'[QRCODE]: ' + (" + expression + ")";
        
        return new DefaultCellTemplate(
                context.getTemplateAddress(),
                context.getColIndex(),
                finalExpression, false, false);
    }
}
```

**2. 注册插件**
```java
ExcelTemplateEngine engine = PoiTemplateEngine.builder()
.registerCellSyntaxHandler(new QrCodeTextHandler())
.build();
```

### 高级示例：通过自定义渲染插入批注

引擎的插件系统不仅能处理值，还能定义复杂的**渲染行为**。

**1. 实现一个能插入批注的 `#comment ` 处理器**
```java
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
```

**2. 在模板中使用**
```
| #comment 批注内容 |
```

---

## 🛠️ 构建

```bash
# 克隆项目
git clone https://github.com/jwj/flexcel.git
cd flexcel

# 运行测试
mvn test

# 打包
mvn package
```

---

## 📜 开源协议

本项目基于 [MIT License](LICENSE) 开源。
