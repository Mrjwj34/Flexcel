/*
 * MIT License
 *
 * Copyright © 2025 jwj
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
 * associated documentation files (the “Software”), to deal in the Software without restriction, including
 *  without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is furnished to do so, subject
 *  to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial
 * portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
 * PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
 * COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
 * AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

package com.github.jwj.flexcel.parser.ast.block;

import com.github.jwj.flexcel.plugin.cell.CellTemplateFactory;
import com.github.jwj.flexcel.parser.ast.template.CellTemplate;
import com.github.jwj.flexcel.parser.ast.template.RowTemplate;
import com.github.jwj.flexcel.style.TemplateStyleInfo; // 【新增】导入 TemplateStyleInfo
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress; // 【新增】导入 CellRangeAddress
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * BlockBuilder 现在不仅依赖于 CellTemplateFactory 解析单元格，
 * 还负责从 TemplateStyleInfo 中提取并管理行级的静态合并区域。
 * 这确保了循环内的合并区域能被正确地复制和应用。
 */
public class BlockBuilder {
    private static final Logger logger = LoggerFactory.getLogger(BlockBuilder.class);

    private final BlockType type;
    private final int startRow;
    private final String directive;
    private final CellTemplateFactory cellTemplateFactory;
    private final TemplateStyleInfo sheetStyleInfo; // 【新增】持有当前 Sheet 的样式信息引用

    private int endRow = -1;
    private String foreachItemName;
    private String foreachCollectionExpr;
    private final List<TemplateBlock> children = new ArrayList<>();
    private final List<Row> templateRows = new ArrayList<>();
    private List<TemplateBlock> thenBranch = null;

    /**
     * 构造函数，新增 TemplateStyleInfo 作为依赖。
     *
     * @param type              块类型
     * @param startRow          起始行
     * @param directive         指令字符串
     * @param cellTemplateFactory 用于创建 CellTemplate 的工厂实例
     * @param sheetStyleInfo    当前 Sheet 的样式信息，用于提取合并区域等
     */
    public BlockBuilder(BlockType type, int startRow, String directive, CellTemplateFactory cellTemplateFactory, TemplateStyleInfo sheetStyleInfo) {
        this.type = type;
        this.startRow = startRow;
        this.directive = directive;
        this.cellTemplateFactory = Objects.requireNonNull(cellTemplateFactory, "CellTemplateFactory cannot be null");
        this.sheetStyleInfo = Objects.requireNonNull(sheetStyleInfo, "TemplateStyleInfo cannot be null");
    }

    // Getters and setters for builder properties. (无变化)
    public int getStartRow() { return startRow; }
    public void setEndRow(int endRow) { this.endRow = endRow; }
    public List<TemplateBlock> getChildren() { flushStaticRows(); return this.children; }
    public void setForeachDetails(String itemName, String collectionExpr) { this.foreachItemName = itemName; this.foreachCollectionExpr = collectionExpr; }
    public void addChild(TemplateBlock child) { flushStaticRows(); this.children.add(child); }
    public void addTemplateRow(Row row) { this.templateRows.add(row); }
    public BlockType getType() { return this.type; }

    /**
     * Splits the current children into a 'then' branch. Used when an #else directive is found for an #if block. (无变化)
     */
    public void splitForElse() {
        if (this.type == BlockType.IF) {
            flushStaticRows();
            this.thenBranch = new ArrayList<>(this.children);
            this.children.clear();
        }
    }

    /**
     * Processes any pending static rows and adds them as a StaticRowsBlock to the children list.
     * 这里的 compileRow 方法将负责处理合并区域。
     */
    private void flushStaticRows() {
        if (!templateRows.isEmpty()) {
            List<RowTemplate> compiledRows = templateRows.stream()
                    .map(this::compileRow)
                    .collect(Collectors.toList());
            if (!compiledRows.isEmpty()) children.add(new StaticRowsBlock(compiledRows));
            templateRows.clear();
        }
    }

    /**
     * Builds the final TemplateBlock based on the builder's current state.
     */
    public TemplateBlock build() {
        flushStaticRows();
        switch (type) {
            case FOREACH:
                return new ForEachBlock(foreachItemName, foreachCollectionExpr, this.startRow, this.endRow, getChildrenAsRowTemplates());
            case IF:
                return new IfBlock(directive, (thenBranch != null) ? thenBranch : this.children, (thenBranch != null) ? this.children : new ArrayList<>());
            default:
                return new RootBlock(this.children);
        }
    }

    /**
     * Extracts all RowTemplate objects from the child blocks.
     */
    private List<RowTemplate> getChildrenAsRowTemplates() {
        return this.children.stream()
                .filter(b -> b instanceof StaticRowsBlock)
                .flatMap(b -> ((StaticRowsBlock) b).getRowTemplates().stream())
                .collect(Collectors.toList());
    }

    /**
     * Compiles a raw POI Row into a RowTemplate object.
     * 在此处识别并提取与当前行相关的合并区域，并传递给 RowTemplate。
     *
     * @param row The POI Row to compile.
     * @return A new RowTemplate.
     */
    private RowTemplate compileRow(Row row) {
        if (row == null) return new RowTemplate(-1, new ArrayList<>(), new ArrayList<>()); // 【修改】新增合并区域参数

        List<CellTemplate> cellTemplates = new ArrayList<>();
        int lastCellNum = row.getLastCellNum();
        if (lastCellNum < 0) lastCellNum = 0;

        for (int i = 0; i < lastCellNum; i++) {
            Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            cellTemplates.add(compileCell(cell, row.getRowNum(), i));
        }

        List<CellRangeAddress> rowMergedRegions = new ArrayList<>();
        // 从 SheetStyleInfo 中查找与当前行相关的合并区域
        // 注意：这里需要迭代 sheetStyleInfo 中的所有合并区域，而不是通过行号直接获取
        // 因为一个合并区域可能跨多行，但它的起始行号决定了它是哪个 RowTemplate 的一部分。
        // 为了避免复杂性，我们认为一个合并区域的起始行就是它所属的 RowTemplate 的行号。
        // 并且，一旦某个合并区域被 RowTemplate “认领”，就从 sheetStyleInfo 中移除。
        List<CellRangeAddress> allSheetMergedRegions = new ArrayList<>(this.sheetStyleInfo.getMergedRegions()); // 复制一份，防止并发修改

        for (CellRangeAddress region : allSheetMergedRegions) {
            if (region.getFirstRow() == row.getRowNum()) {
                rowMergedRegions.add(region);
                this.sheetStyleInfo.removeMergedRegion(region); // 【关键】从 TemplateStyleInfo 中移除，防止重复处理
            }
        }

        return new RowTemplate(row.getRowNum(), cellTemplates, rowMergedRegions); // 【修改】传入合并区域
    }

    /**
     * Compiles a single POI Cell into a CellTemplate.
     */
    private CellTemplate compileCell(Cell cell, int rowNum, int colIndex) {
        return this.cellTemplateFactory.create(cell, rowNum, colIndex);
    }
}