package com.github.jwj.flexcel.engine.model;

import com.github.jwj.flexcel.parser.ast.template.PrecompiledTemplate;
import com.github.jwj.flexcel.style.TemplateStyleInfo;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 封装了模板经过分析和编译后的结果。
 * 这是一个不可变的DTO，作为通用编译核心与具体执行引擎之间的数据契约。
 * 它不依赖任何具体的Excel写入库（如SXSSF）。
 *
 * @author jwj
 * @since 2025
 */
public final class AnalyzedTemplate {

    private final List<String> sheetOrder;
    private final Map<String, PrecompiledTemplate> compiledSheets;
    private final Map<String, TemplateStyleInfo> sheetStyles;

    public AnalyzedTemplate(List<String> sheetOrder,
                            Map<String, PrecompiledTemplate> compiledSheets,
                            Map<String, TemplateStyleInfo> sheetStyles) {
        // 使用防御性拷贝确保不可变性
        this.sheetOrder = Collections.unmodifiableList(new ArrayList<>(sheetOrder));
        this.compiledSheets = Collections.unmodifiableMap(new HashMap<>(compiledSheets));
        this.sheetStyles = Collections.unmodifiableMap(new HashMap<>(sheetStyles));
    }

    /**
     * 获取模板中工作表的原始顺序。
     * @return 不可修改的工作表名称列表。
     */
    public List<String> getSheetOrder() {
        return sheetOrder;
    }

    /**
     * 获取已编译的工作表AST。
     * Key是工作表名称，Value是其对应的预编译模板（AST的根）。
     * @return 不可修改的Map。
     */
    public Map<String, PrecompiledTemplate> getCompiledSheets() {
        return compiledSheets;
    }

    /**
     * 获取已分析的工作表样式信息。
     * Key是工作表名称，Value是其对应的样式信息。
     * @return 不可修改的Map。
     */
    public Map<String, TemplateStyleInfo> getSheetStyles() {
        return sheetStyles;
    }
}