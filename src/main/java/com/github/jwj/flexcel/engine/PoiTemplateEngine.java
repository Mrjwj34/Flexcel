/*
 * MIT License
 *
 * Copyright © 2025 jwj
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
 * associated documentation files (the "Software"), to deal in the Software without restriction, including
 *  without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is furnished to do so, subject
 *  to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial
 * portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
 * PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
 * COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
 * AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

package com.github.jwj.flexcel.engine;

import com.github.jwj.flexcel.plugin.cell.CellTemplateFactory;
import com.github.jwj.flexcel.plugin.cell.CellSyntaxHandler;
import com.github.jwj.flexcel.plugin.cell.DefaultCellHandler;
import com.github.jwj.flexcel.plugin.cell.FormulaCellHandler;
import com.github.jwj.flexcel.plugin.cell.MergeCellHandler;
import com.github.jwj.flexcel.runtime.dto.MergeableRenderedCell;
import com.github.jwj.flexcel.runtime.dto.RenderedCell;
import com.github.jwj.flexcel.runtime.dto.RenderedRow;
import com.github.jwj.flexcel.runtime.pool.DefaultObjectPool;
import com.github.jwj.flexcel.runtime.pool.NoOpObjectPool;
import com.github.jwj.flexcel.runtime.pool.ObjectPool;
import com.github.jwj.flexcel.runtime.TemplateContext;
import com.github.jwj.flexcel.parser.expression.ExpressionEvaluator;
import com.github.jwj.flexcel.parser.expression.SpelExpressionEvaluator;
import com.github.jwj.flexcel.parser.TemplateCompiler;
import com.github.jwj.flexcel.parser.ast.block.RootBlock;
import com.github.jwj.flexcel.parser.ast.block.StaticRowsBlock;
import com.github.jwj.flexcel.parser.ast.block.TemplateBlock;
import com.github.jwj.flexcel.parser.ast.block.ForEachBlock;
import com.github.jwj.flexcel.parser.ast.block.IfBlock;
import com.github.jwj.flexcel.parser.ast.template.PrecompiledTemplate;
import com.github.jwj.flexcel.parser.ast.template.RowTemplate;
import com.github.jwj.flexcel.plugin.block.*;
import com.github.jwj.flexcel.style.StyleMappingManager;
import com.github.jwj.flexcel.style.TemplateStyleInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * <p>
 * 基于 Apache POI 的高性能流式 Excel 模板引擎。
 * </p>
 */
public class PoiTemplateEngine implements ExcelTemplateEngine {

    private static final Logger logger = LoggerFactory.getLogger(PoiTemplateEngine.class);

    private final int sxssfWindowSize;
    private final int queueCapacity;
    private final Map<String, Object> services;
    private final Map<String, Object> globalContext;
    private final ExpressionEvaluator expressionEvaluator;
    private final ObjectPool objectPool;
    private final CellTemplateFactory cellTemplateFactory;
    private final List<BlockDirectiveHandler> blockDirectiveHandlers;

    public static Builder builder() {
        return new Builder();
    }

    private PoiTemplateEngine(Builder builder) {
        this.sxssfWindowSize = builder.sxssfWindowSize;
        this.queueCapacity = builder.queueCapacity;
        this.services = new HashMap<>(builder.services);
        this.globalContext = new HashMap<>(builder.globalContext);
        this.expressionEvaluator = builder.expressionEvaluator;
        this.objectPool = builder.objectPool;
        this.cellTemplateFactory = builder.cellTemplateFactory;
        this.blockDirectiveHandlers = builder.blockDirectiveHandlers;

        logger.info("PoiTemplateEngine created. CellHandlers: {}. BlockHandlers: {}",
                builder.getRegisteredCellHandlerNames(),
                builder.getRegisteredBlockHandlerNames()
        );
    }

    public static class Builder {
        private int sxssfWindowSize = 1000;
        private int queueCapacity = 2048;
        private final Map<String, Object> services = new HashMap<>();
        private final Map<String, Object> globalContext = new HashMap<>();
        private ExpressionEvaluator expressionEvaluator;
        private boolean unsafeSpelOperationsEnabled = false;
        private boolean objectPoolingEnabled = true;
        private int rowPoolCapacity = 1024;
        private int cellPoolCapacity = 4096;
        private int mergeableCellPoolCapacity = 1024;
        private ObjectPool objectPool;

        private CellTemplateFactory cellTemplateFactory;
        private List<BlockDirectiveHandler> blockDirectiveHandlers;

        private final List<CellSyntaxHandler> customCellSyntaxHandlers = new ArrayList<>();
        private final List<BlockDirectiveHandler> customBlockDirectiveHandlers = new ArrayList<>();

        private final List<String> registeredCellHandlerNames = new ArrayList<>();
        private final List<String> registeredBlockHandlerNames = new ArrayList<>();

        private Builder() {}

        /**
         * 启用不安全的 SpEL 操作。
         * 这将允许引擎使用全反射能力，可能会带来安全风险。
         * 请确保模板来源可信。
         */
        public Builder enableUnsafeSpelOperations() {
            logger.warn("Enabling unsafe SpEL operations. This allows full reflection capabilities and poses a security risk if templates are not from a trusted source.");
            this.unsafeSpelOperationsEnabled = true;
            return this;
        }
        /**
         * 设置 SXSSF 窗口大小。
         * 窗口大小决定了内存中保留的行数，影响性能和内存使用。
         *
         * @param sxssfWindowSize 窗口大小，必须大于 0。
         * @return 当前 Builder 实例。
         */
        public Builder sxssfWindowSize(int sxssfWindowSize) {
            if (sxssfWindowSize <= 0) throw new IllegalArgumentException("SXSSF window size must be positive.");
            this.sxssfWindowSize = sxssfWindowSize;
            return this;
        }
        /**
         * 设置队列容量。
         * 队列用于在单生产者-消费者模型中传递渲染的行数据。
         *
         * @param queueCapacity 队列容量，必须大于 0。
         * @return 当前 Builder 实例。
         */
        public Builder queueCapacity(int queueCapacity) {
            if (queueCapacity <= 0) throw new IllegalArgumentException("Queue capacity must be positive.");
            this.queueCapacity = queueCapacity;
            return this;
        }
        /**
         * 设置表达式求值器。
         * 如果未设置，将使用默认的 SpEL 表达式求值器。
         *
         * @param expressionEvaluator 自定义的表达式求值器实例。
         * @return 当前 Builder 实例。
         */
        public Builder withExpressionEvaluator(ExpressionEvaluator expressionEvaluator) {
            if (expressionEvaluator == null) {
                throw new IllegalArgumentException("ExpressionEvaluator cannot be null.");
            }
            this.expressionEvaluator = expressionEvaluator;
            return this;
        }
        /**
         * 禁用对象池。
         * 如果禁用，将不使用对象池来管理行和单元格的复用。
         *
         * @return 当前 Builder 实例。
         */
        public Builder disableObjectPooling() {
            this.objectPoolingEnabled = false;
            return this;
        }
        /**
         * 设置对象池的容量。
         * 对象池用于复用行和单元格，减少内存分配和垃圾回收压力。
         *
         * @param rowCapacity 行池容量，必须大于 0。
         * @param cellCapacity 单元格池容量，必须大于 0。
         * @param mergeableCellCapacity 可合并单元格池容量，必须大于 0。
         * @return 当前 Builder 实例。
         */
        public Builder objectPoolCapacities(int rowCapacity, int cellCapacity, int mergeableCellCapacity) {
            if (rowCapacity <= 0 || cellCapacity <= 0 || mergeableCellCapacity <= 0) {
                throw new IllegalArgumentException("Pool capacities must be positive.");
            }
            this.rowPoolCapacity = rowCapacity;
            this.cellPoolCapacity = cellCapacity;
            this.mergeableCellPoolCapacity = mergeableCellCapacity;
            return this;
        }
        /**
         * 设置自定义对象池。
         * 如果设置了自定义对象池，将使用该池来管理行和单元格的复用。
         *
         * @param objectPool 自定义的对象池实例，不能为空。
         * @return 当前 Builder 实例。
         */
        public Builder withObjectPool(ObjectPool objectPool) {
            if (objectPool == null) {
                throw new IllegalArgumentException("ObjectPool cannot be null.");
            }
            this.objectPool = objectPool;
            return this;
        }

        /**
         * 注册一个服务类，引擎会根据配置智能决定注册实例还是类本身。
         *
         * <p><b>注册逻辑如下:</b></p>
         * <ol>
         *   <li>引擎首先尝试使用该类的公共无参构造函数创建一个<b>实例</b>。
         *       <ul>
         *          <li>如果成功，则注册该实例。这是最常见的用例，适用于普通服务类。
         *              例如，注册 {@code MyService.class} 后，可通过 {@code ${myService.someMethod()}} 调用。</li>
         *       </ul>
         *   </li>
         *   <li>如果实例化失败（例如，因为构造函数是私有的）:
         *       <ul>
         *          <li>如果已通过 {@link #enableUnsafeSpelOperations()} <b>启用了不安全模式</b>，
         *              引擎会回退并注册该类的 {@code Class} 对象本身。这用于支持调用<b>静态方法</b>。
         *              例如，注册 {@code MyUtils.class} 后，可通过 {@code ${myUtils.someStaticMethod()}} 调用。</li>
         *          <li>如果在<b>安全模式</b>下实例化失败，引擎会抛出 {@link RuntimeException}，
         *              因为在这种模式下，既无法创建实例，也无法调用静态方法，注册将无效。</li>
         *       </ul>
         *   </li>
         * </ol>
         *
         * @param serviceClass 要注册的服务类。
         * @param serviceName  可选的服务名称。如果未提供，将使用类名的首字母小写形式（例如, {@code MyService} -> {@code myService}）。
         * @return 当前 Builder 实例。
         * @throws IllegalArgumentException 如果 serviceClass 为空。
         * @throws RuntimeException 如果在安全模式下无法实例化该类。
         */
        public Builder registerService(Class<?> serviceClass, String... serviceName) {
            if (serviceClass == null) {
                throw new IllegalArgumentException("Service class cannot be null.");
            }

            String name;
            if (serviceName != null && serviceName.length > 0 && serviceName[0] != null && !serviceName[0].isEmpty()) {
                name = serviceName[0];
            } else {
                String simpleName = serviceClass.getSimpleName();
                name = Character.toLowerCase(simpleName.charAt(0)) + simpleName.substring(1);
            }

            try {
                // 1. 优先尝试创建实例
                Object serviceInstance = serviceClass.getDeclaredConstructor().newInstance();
                this.services.put(name, serviceInstance);
                logger.info("Registered service class '{}' as an instance with name '{}'", serviceClass.getName(), name);
            } catch (Exception instantiationException) {
                // 2. 实例化失败，根据安全模式开关决定下一步
                if (this.unsafeSpelOperationsEnabled) {
                    // 2a. 在不安全模式下，注册Class对象以调用静态方法
                    this.services.put(name, serviceClass);
                    logger.warn("Could not instantiate service class '{}'. As unsafe operations are enabled, " +
                                    "registering the Class object itself for static method access under the name '{}'.",
                            serviceClass.getName(), name);
                } else {
                    // 2b. 在安全模式下，无法创建实例也无法调用静态方法，因此这是一个错误
                    throw new RuntimeException(
                            "Failed to instantiate service class '" + serviceClass.getName() + "' in SAFE mode. " +
                                    "Please ensure it has a public no-arg constructor, or enable unsafe operations " +
                                    "via .enableUnsafeSpelOperations() to use it as a static utility class.",
                            instantiationException
                    );
                }
            }
            return this;
        }

        /**
         * 设置全局上下文。
         * 全局上下文中的变量可以在模板中通过表达式访问。
         *
         * @param context 全局上下文，不能为空。
         * @return 当前 Builder 实例。
         */
        public Builder withGlobalContext(Map<String, Object> context) {
            if (context != null) {
                this.globalContext.putAll(context);
            }
            return this;
        }
        /**
         * 注册自定义单元格语法处理器。
         * 这些处理器可以用于处理特定的单元格语法，如公式、合并等。
         *
         * @param handler 自定义单元格语法处理器，不能为空。
         * @return 当前 Builder 实例。
         */
        public Builder registerCellSyntaxHandler(CellSyntaxHandler handler) {
            if (handler != null) {
                this.customCellSyntaxHandlers.add(handler);
            }
            return this;
        }
        /**
         * 注册自定义块指令处理器。
         * 这些处理器可以用于处理特定的块指令，如循环、条件等。
         *
         * @param handler 自定义块指令处理器，不能为空。
         * @return 当前 Builder 实例。
         */
        public Builder registerBlockDirectiveHandler(BlockDirectiveHandler handler) {
            if (handler != null) {
                this.customBlockDirectiveHandlers.add(handler);
            }
            return this;
        }

        private List<String> getRegisteredCellHandlerNames() { return this.registeredCellHandlerNames; }
        private List<String> getRegisteredBlockHandlerNames() { return this.registeredBlockHandlerNames; }

        public PoiTemplateEngine build() {
            if (this.expressionEvaluator == null) {
                this.expressionEvaluator = new SpelExpressionEvaluator(this.unsafeSpelOperationsEnabled);
            }
            if (this.objectPool == null) {
                this.objectPool = this.objectPoolingEnabled
                        ? new DefaultObjectPool(rowPoolCapacity, cellPoolCapacity, mergeableCellPoolCapacity)
                        : new NoOpObjectPool();
            }

            List<CellSyntaxHandler> allCellHandlers = new ArrayList<>();
            allCellHandlers.add(new FormulaCellHandler());
            allCellHandlers.add(new MergeCellHandler());
            allCellHandlers.addAll(this.customCellSyntaxHandlers);
            allCellHandlers.add(new DefaultCellHandler());
            allCellHandlers.forEach(h -> this.registeredCellHandlerNames.add(h.getClass().getSimpleName()));
            this.cellTemplateFactory = new CellTemplateFactory(allCellHandlers);

            this.blockDirectiveHandlers = new ArrayList<>();
            this.blockDirectiveHandlers.addAll(this.customBlockDirectiveHandlers);
            this.blockDirectiveHandlers.add(new ForEachDirectiveHandler());
            this.blockDirectiveHandlers.add(new IfDirectiveHandler());
            this.blockDirectiveHandlers.add(new ElseDirectiveHandler());
            this.blockDirectiveHandlers.add(new EndDirectiveHandler());
            this.blockDirectiveHandlers.forEach(h -> this.registeredBlockHandlerNames.add(h.getClass().getSimpleName()));

            return new PoiTemplateEngine(this);
        }
    }

    @Override
    public void process(InputStream templateStream, Map<String, Object> data, OutputStream outputStream) {
        long startTime = System.currentTimeMillis();
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = templateStream.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
            byte[] templateBytes = baos.toByteArray();

            logger.info("Analysis Phase: Starting style and position mapping analysis...");
            long analysisStart = System.currentTimeMillis();
            Map<String, TemplateStyleInfo> allSheetsStyleInfo = new HashMap<>();
            List<String> sheetOrder = new ArrayList<>();
            try (Workbook analysisWorkbook = WorkbookFactory.create(new ByteArrayInputStream(templateBytes))) {
                StyleMappingManager styleMappingManager = new StyleMappingManager();
                for (Sheet sheet : analysisWorkbook) {
                    sheetOrder.add(sheet.getSheetName());
                    allSheetsStyleInfo.put(sheet.getSheetName(), styleMappingManager.extractTemplateStyles(sheet));
                }
            }
            logger.info("Analysis Phase: Completed in {}ms", System.currentTimeMillis() - analysisStart);

            logger.info("Compilation Phase: Starting template pre-compilation...");
            long compileStart = System.currentTimeMillis();
            Map<String, PrecompiledTemplate> compiledTemplates = new HashMap<>();
            try (Workbook compileWorkbook = WorkbookFactory.create(new ByteArrayInputStream(templateBytes))) {
                // 【修复】将 styleInfo 传入编译器，以便其分派合并区域
                TemplateCompiler compiler = new TemplateCompiler(this.cellTemplateFactory, this.blockDirectiveHandlers);
                for (String sheetName : sheetOrder) {
                    Sheet sheet = compileWorkbook.getSheet(sheetName);
                    if (sheet != null) {
                        compiledTemplates.put(sheetName, compiler.compile(sheet, allSheetsStyleInfo.get(sheetName)));
                    }
                }
            }
            logger.info("Compilation Phase: Completed in {}ms", System.currentTimeMillis() - compileStart);

            Map<String, String> stringCache = new WeakHashMap<>();

            logger.info("Execution Phase: Starting single-producer-consumer model...");
            long generateStart = System.currentTimeMillis();
            try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(null, sxssfWindowSize, true, false)) {
                Map<CellStyle, CellStyle> styleCache = new HashMap<>();
                prepareStyles(outputWorkbook, allSheetsStyleInfo.values(), styleCache);

                for (String sheetName : sheetOrder) {
                    logger.info("Processing sheet: {}", sheetName);
                    PrecompiledTemplate compiledTemplate = compiledTemplates.get(sheetName);
                    if (compiledTemplate == null) continue;

                    SXSSFSheet outputSheet = outputWorkbook.createSheet(sheetName);
                    TemplateContext baseContext = new TemplateContext(data, services, globalContext);
                    TemplateStyleInfo styleInfo = allSheetsStyleInfo.get(sheetName);
                    styleInfo.getAllColumnWidths().forEach(outputSheet::setColumnWidth);

                    BlockingQueue<RenderedRow> queue = new ArrayBlockingQueue<>(queueCapacity);
                    Future<?> consumerFuture = startConsumerThread(queue, outputSheet, styleInfo, styleCache, this.objectPool);

                    try {
                        AtomicInteger globalRowCounter = new AtomicInteger(0);
                        executeTemplate(compiledTemplate.getRootBlocks(), baseContext, queue, this.objectPool, globalRowCounter, stringCache);
                    } catch (Exception e) {
                        consumerFuture.cancel(true);
                        throw new RuntimeException("Producer thread failed during template execution", e);
                    } finally {
                        queue.put(RenderedRow.POISON_PILL);
                    }

                    consumerFuture.get();
                    // 【修复】此时 styleInfo 中的合并区域只剩下真正的静态区域了
                    applyStaticMergedRegions(outputSheet, styleInfo);
                }
                logger.info("Writing data to output stream...");
                outputWorkbook.write(outputStream);
            }
            logger.info("Execution Phase: Completed in {}ms", System.currentTimeMillis() - generateStart);
        } catch (Exception e) {
            throw new RuntimeException("Template processing failed", e);
        }
        logger.info("Template processing successful. Total time: {}ms", System.currentTimeMillis() - startTime);
    }

    private void executeTemplate(List<TemplateBlock> blocks, TemplateContext context,
                                 BlockingQueue<RenderedRow> queue, ObjectPool pool,
                                 AtomicInteger globalRowCounter, Map<String, String> stringCache) throws InterruptedException {
        for (TemplateBlock block : blocks) {
            if (block instanceof IfBlock) {
                IfBlock ifBlock = (IfBlock) block;
                if (ifBlock.evaluateCondition(context, this.expressionEvaluator)) {
                    executeTemplate(ifBlock.getThenBlocks(), context, queue, pool, globalRowCounter, stringCache);
                } else {
                    executeTemplate(ifBlock.getElseBlocks(), context, queue, pool, globalRowCounter, stringCache);
                }
            } else if (block instanceof ForEachBlock) {
                ForEachBlock feBlock = (ForEachBlock) block;
                int loopStartRowNo = globalRowCounter.get() + 1;

                Object itemsObject = this.expressionEvaluator.evaluate(feBlock.getCollectionExpression(), context.getAllData());
                if (itemsObject instanceof Iterable) {
                    int index = 0;
                    for (Object item : (Iterable<?>) itemsObject) {
                        TemplateContext itemContext = new TemplateContext(context);
                        itemContext.setVariable(feBlock.getItemName(), item);
                        itemContext.setVariable(feBlock.getIndexName(), index);
                        itemContext.setVariable("currentRowNo", globalRowCounter.incrementAndGet());
                        for (RowTemplate rt : feBlock.getRowTemplates()) {
                            queue.put(rt.produce(itemContext, pool, stringCache, this.expressionEvaluator));
                        }
                        index++;
                    }
                } else {
                    logger.warn("Expression '{}' is not iterable, skipping.", feBlock.getCollectionExpression());
                }

                int loopEndRowNo = globalRowCounter.get();
                // 【修复】无论循环是否执行，都设置变量，以防后续表达式求值失败
                context.setVariable("startRowNo", loopStartRowNo);
                context.setVariable("endRowNo", loopEndRowNo);

            } else if (block instanceof StaticRowsBlock) {
                for (RowTemplate rt : ((StaticRowsBlock) block).getRowTemplates()) {
                    TemplateContext staticContext = new TemplateContext(context);
                    staticContext.setVariable("currentRowNo", globalRowCounter.incrementAndGet());
                    queue.put(rt.produce(staticContext, pool, stringCache, this.expressionEvaluator));
                }
            } else if (block instanceof RootBlock) {
                executeTemplate(((RootBlock) block).getChildren(), context, queue, pool, globalRowCounter, stringCache);
            }
        }
    }

    private Future<?> startConsumerThread(BlockingQueue<RenderedRow> queue, SXSSFSheet sheet, TemplateStyleInfo styleInfo,
                                          Map<CellStyle, CellStyle> styleCache, ObjectPool pool) {
        ExecutorService consumerExecutor = Executors.newSingleThreadExecutor(r -> new Thread(r, "fastexcel-consumer"));
        Runnable consumerTask = () -> {
            try {
                consume(queue, sheet, styleInfo, styleCache, pool);
            } catch (Exception e) {
                logger.error("Consumer thread terminated due to a critical error!", e);
                throw new RuntimeException(e);
            }
        };
        Future<?> future = consumerExecutor.submit(consumerTask);
        consumerExecutor.shutdown();
        return future;
    }

    private void consume(BlockingQueue<RenderedRow> queue, SXSSFSheet sheet, TemplateStyleInfo styleInfo,
                         Map<CellStyle, CellStyle> styleCache, ObjectPool pool) throws InterruptedException {
        Map<Integer, CellStyle[]> rowStyleCache = new HashMap<>();
        int currentRowIndex = 0;
        Map<Integer, Object> lastValuesForMerge = new HashMap<>();
        Map<Integer, Integer> mergeStartRows = new HashMap<>();

        while (true) {
            RenderedRow rowData = queue.take();
            if (rowData == RenderedRow.POISON_PILL) {
                break;
            }

            Row outputRow = sheet.createRow(currentRowIndex);
            CellStyle[] styles = rowStyleCache.computeIfAbsent(rowData.templateRowNum, k -> {
                CellStyle[] newStyles = new CellStyle[rowData.cells.size()];
                for (int i = 0; i < rowData.cells.size(); i++) {
                    newStyles[i] = findCellStyle(rowData.cells.get(i).templateAddress, styleInfo, styleCache);
                }
                return newStyles;
            });

            Float height = styleInfo.getAllRowHeights().get(rowData.templateRowNum);
            if (height != null) {
                outputRow.setHeightInPoints(height);
            }

            for (int i = 0; i < rowData.cells.size(); i++) {
                RenderedCell cellData = rowData.cells.get(i);
                int colIdx = cellData.colIndex;

                if (cellData instanceof MergeableRenderedCell) {
                    handleMergeableCell(sheet, outputRow, (MergeableRenderedCell) cellData, styles[i], currentRowIndex, lastValuesForMerge, mergeStartRows);
                } else {
                    handleRegularCell(sheet, outputRow, cellData, styles[i], currentRowIndex, lastValuesForMerge, mergeStartRows);
                }
            }

            // 在写入行后，应用该行模板附带的静态合并区域
            applyRowStaticMergedRegions(sheet, rowData, currentRowIndex);

            pool.returnRow(rowData);
            currentRowIndex++;
        }

        for (Map.Entry<Integer, Integer> entry : mergeStartRows.entrySet()) {
            int colIdx = entry.getKey();
            int startRow = entry.getValue();
            if (currentRowIndex - 1 > startRow) {
                sheet.addMergedRegion(new CellRangeAddress(startRow, currentRowIndex - 1, colIdx, colIdx));
            }
        }
    }

    // handleMergeableCell, handleRegularCell, completePendingMerge, createAndSetCell, 等方法保持不变...
    private void handleMergeableCell(SXSSFSheet sheet, Row outputRow, MergeableRenderedCell cellData, CellStyle style,
                                     int currentRowIndex, Map<Integer, Object> lastValuesForMerge, Map<Integer, Integer> mergeStartRows) {
        int colIdx = cellData.colIndex;
        Object currentValue = cellData.value;
        Object lastValue = lastValuesForMerge.get(colIdx);

        boolean isSameAsLast = (currentValue instanceof String && lastValue instanceof String)
                ? (currentValue == lastValue)
                : Objects.equals(currentValue, lastValue);

        if (isSameAsLast && currentRowIndex > 0) {
            if (!mergeStartRows.containsKey(colIdx)) {
                mergeStartRows.put(colIdx, currentRowIndex - 1);
            }
            Cell emptyCell = outputRow.createCell(colIdx);
            if (style != null) emptyCell.setCellStyle(style);
        } else {
            completePendingMerge(sheet, colIdx, currentRowIndex, mergeStartRows);
            createAndSetCell(outputRow, cellData, style);
        }
        lastValuesForMerge.put(colIdx, currentValue);
    }

    private void handleRegularCell(SXSSFSheet sheet, Row outputRow, RenderedCell cellData, CellStyle style,
                                   int currentRowIndex, Map<Integer, Object> lastValuesForMerge, Map<Integer, Integer> mergeStartRows) {
        int colIdx = cellData.colIndex;
        completePendingMerge(sheet, colIdx, currentRowIndex, mergeStartRows);
        createAndSetCell(outputRow, cellData, style);
        lastValuesForMerge.remove(colIdx);
    }

    private void completePendingMerge(SXSSFSheet sheet, int colIdx, int currentRowIndex, Map<Integer, Integer> mergeStartRows) {
        if (mergeStartRows.containsKey(colIdx)) {
            int startRow = mergeStartRows.remove(colIdx);
            if (currentRowIndex - 1 > startRow) {
                sheet.addMergedRegion(new CellRangeAddress(startRow, currentRowIndex - 1, colIdx, colIdx));
            }
        }
    }

    /**
     * 此方法现在负责执行自定义渲染逻辑。
     */
    private void createAndSetCell(Row outputRow, RenderedCell cellData, CellStyle style) {
        // 创建单元格是必须的，因为自定义渲染器需要一个 Cell 对象作为操作目标
        Cell outputCell = outputRow.createCell(cellData.colIndex);

        // 检查是否存在自定义渲染器
        if (cellData.customRenderer != null) {
            try {
                // 如果有，就执行它，而不是进行常规的值设置
                cellData.customRenderer.accept(outputCell);
            } catch (Exception e) {
                logger.error("Custom renderer failed for cell at {}:{}. Error: {}",
                        outputRow.getRowNum(), cellData.colIndex, e.getMessage(), e);
                outputCell.setCellValue("##RENDER_ERROR##");
            }
        } else {
            // 否则，走原来的逻辑
            if (cellData.isFormula) {
                try {
                    if (cellData.value != null) {
                        outputCell.setCellFormula(cellData.value.toString());
                    }
                } catch (Exception e) {
                    logger.error("Failed to set formula: '{}' at row {}, col {}", cellData.value, outputRow.getRowNum(), outputCell.getColumnIndex(), e);
                    outputCell.setCellValue("##FORMULA_ERROR##");
                }
            } else {
                setCellValue(outputCell, cellData.value);
            }
        }

        if (style != null) {
            outputCell.setCellStyle(style);
        }
    }

    /**
     * 应用与单行模板关联的静态合并区域。
     *
     * @param sheet           输出的 Sheet。
     * @param rowData         当前渲染的行数据，其中包含模板合并信息。
     * @param currentRowIndex 当前写入的行索引。
     */
    private void applyRowStaticMergedRegions(SXSSFSheet sheet, RenderedRow rowData, int currentRowIndex) {
        List<CellRangeAddress> rowMerges = rowData.getStaticMergedRegions();
        if (rowMerges == null || rowMerges.isEmpty()) {
            return;
        }

        for (CellRangeAddress templateRegion : rowMerges) {
            // 计算新的合并区域，行号基于当前写入位置，列和跨度保持不变
            int rowSpan = templateRegion.getLastRow() - templateRegion.getFirstRow();
            CellRangeAddress newRegion = new CellRangeAddress(
                    currentRowIndex,
                    currentRowIndex + rowSpan,
                    templateRegion.getFirstColumn(),
                    templateRegion.getLastColumn()
            );
            try {
                sheet.addMergedRegion(newRegion);
            } catch (Exception e) {
                logger.warn("Failed to add row-template-based merged region (might overlap): {}", newRegion.formatAsString());
            }
        }
    }

    // findCellStyle, setCellValue, prepareStyles, 等方法保持不变...
    private CellStyle findCellStyle(String templateAddress, TemplateStyleInfo styleInfo, Map<CellStyle, CellStyle> styleCache) {
        if (templateAddress == null) return null;
        CellStyle templateStyle = styleInfo.getCellStyle(templateAddress);
        return (templateStyle != null) ? styleCache.get(templateStyle) : null;
    }

    protected void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        if (value instanceof String) cell.setCellValue((String) value);
        else if (value instanceof Number) cell.setCellValue(((Number) value).doubleValue());
        else if (value instanceof Date) cell.setCellValue((Date) value);
        else if (value instanceof Calendar) cell.setCellValue((Calendar) value);
        else if (value instanceof Boolean) cell.setCellValue((Boolean) value);
        else if (value instanceof RichTextString) cell.setCellValue((RichTextString) value);
        else cell.setCellValue(value.toString());
    }

    private void prepareStyles(Workbook workbook, Collection<TemplateStyleInfo> allStyleInfos, Map<CellStyle, CellStyle> styleCache) {
        Set<CellStyle> uniqueTemplateStyles = new HashSet<>();
        for (TemplateStyleInfo styleInfo : allStyleInfos) {
            uniqueTemplateStyles.addAll(styleInfo.getAllTemplateStyles());
        }
        for (CellStyle templateStyle : uniqueTemplateStyles) {
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(templateStyle);
            styleCache.put(templateStyle, newStyle);
        }
        logger.info("Pre-created {} unique styles.", styleCache.size());
    }

    private void applyStaticMergedRegions(Sheet outputSheet, TemplateStyleInfo styleInfo) {
        for (CellRangeAddress staticRegion : styleInfo.getMergedRegions()) {
            try {
                outputSheet.addMergedRegion(staticRegion);
            } catch (Exception e) {
                logger.warn("Failed to add static merged region (might overlap with dynamic merges): {}", staticRegion.formatAsString());
            }
        }
    }
}