package com.github.jwj.flexcel.engine;

import com.github.jwj.flexcel.engine.model.AnalyzedTemplate;
import com.github.jwj.flexcel.parser.TemplateCompiler;
import com.github.jwj.flexcel.parser.ast.template.PrecompiledTemplate;
import com.github.jwj.flexcel.plugin.block.BlockDirectiveHandler;
import com.github.jwj.flexcel.plugin.cell.CellTemplateFactory;
import com.github.jwj.flexcel.style.StyleMappingManager;
import com.github.jwj.flexcel.style.TemplateStyleInfo;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Flexcel 的通用编译核心。
 * 负责将模板输入流（InputStream）分析并编译成一个与执行引擎无关的 {@link AnalyzedTemplate} 对象。
 * 这个过程包括样式提取和模板语法到AST的转换。
 *
 * @author jwj
 * @since 2025
 */
public final class FlexcelCompiler {

    private static final Logger logger = LoggerFactory.getLogger(FlexcelCompiler.class);

    private final CellTemplateFactory cellTemplateFactory;
    private final List<BlockDirectiveHandler> blockDirectiveHandlers;

    /**
     * 构造一个编译器实例。
     * @param cellTemplateFactory 单元格语法处理器工厂
     * @param blockDirectiveHandlers 块指令处理器列表
     */
    public FlexcelCompiler(CellTemplateFactory cellTemplateFactory, List<BlockDirectiveHandler> blockDirectiveHandlers) {
        this.cellTemplateFactory = cellTemplateFactory;
        this.blockDirectiveHandlers = blockDirectiveHandlers;
    }

    /**
     * 执行分析和编译。
     * @param templateStream 模板文件的输入流
     * @return 封装了编译结果的 {@link AnalyzedTemplate}
     * @throws IOException 如果读取模板流失败
     */
    public AnalyzedTemplate compile(InputStream templateStream) throws IOException {
        // 1. 将输入流缓冲到内存
        byte[] templateBytes = readStreamToBytes(templateStream);

        // 2. 分析阶段：提取样式和元数据
        logger.info("Analysis Phase: Starting style and position mapping analysis...");
        long analysisStart = System.currentTimeMillis();
        // 这里声明的变量名是 allSheetsStyleInfo
        Map<String, TemplateStyleInfo> allSheetsStyleInfo = new HashMap<>();
        List<String> sheetOrder = new ArrayList<>();
        try (Workbook analysisWorkbook = WorkbookFactory.create(new ByteArrayInputStream(templateBytes))) {
            StyleMappingManager styleMappingManager = new StyleMappingManager();
            for (Sheet sheet : analysisWorkbook) {
                String sheetName = sheet.getSheetName();
                sheetOrder.add(sheetName);
                allSheetsStyleInfo.put(sheetName, styleMappingManager.extractTemplateStyles(sheet));
            }
        }
        logger.info("Analysis Phase: Completed in {}ms", System.currentTimeMillis() - analysisStart);


        // 3. 编译阶段：构建AST
        logger.info("Compilation Phase: Starting template pre-compilation...");
        long compileStart = System.currentTimeMillis();
        Map<String, PrecompiledTemplate> compiledTemplates = new HashMap<>();
        try (Workbook compileWorkbook = WorkbookFactory.create(new ByteArrayInputStream(templateBytes))) {
            TemplateCompiler compiler = new TemplateCompiler(this.cellTemplateFactory, this.blockDirectiveHandlers);
            for (String sheetName : sheetOrder) {
                Sheet sheet = compileWorkbook.getSheet(sheetName);
                if (sheet != null) {
                    compiledTemplates.put(sheetName, compiler.compile(sheet, allSheetsStyleInfo.get(sheetName)));
                }
            }
        }
        logger.info("Compilation Phase: Completed in {}ms", System.currentTimeMillis() - compileStart);

        return new AnalyzedTemplate(sheetOrder, compiledTemplates, allSheetsStyleInfo);
    }

    private byte[] readStreamToBytes(InputStream inputStream) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[8192];
        int bytesRead;
        while ((bytesRead = inputStream.read(buffer)) != -1) {
            baos.write(buffer, 0, bytesRead);
        }
        return baos.toByteArray();
    }
}