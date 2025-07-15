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

// File: com/github/jwj/flexcel/parser/precompile/TemplateCompiler.java
package com.github.jwj.flexcel.parser;

import com.github.jwj.flexcel.plugin.cell.CellTemplateFactory;
import com.github.jwj.flexcel.plugin.block.BlockCompilerContext;
import com.github.jwj.flexcel.plugin.block.BlockDirectiveHandler;
import com.github.jwj.flexcel.parser.ast.block.BlockBuilder;
import com.github.jwj.flexcel.parser.ast.block.BlockType;
import com.github.jwj.flexcel.parser.ast.template.PrecompiledTemplate;
import com.github.jwj.flexcel.style.TemplateStyleInfo; // 【新增】导入 TemplateStyleInfo
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Objects;
import java.util.Stack;

/**
 * 模板编译器现在在编译 Sheet 时，接收 TemplateStyleInfo，并将其传递给 BlockBuilder，
 * 以便 BlockBuilder 能够处理行级的静态合并区域。
 */
public class TemplateCompiler {
    private static final Logger logger = LoggerFactory.getLogger(TemplateCompiler.class);

    private final CellTemplateFactory cellTemplateFactory;
    private final List<BlockDirectiveHandler> blockDirectiveHandlers;

    /**
     * 构造函数，保持不变。TemplateStyleInfo 在 compile 方法中接收。
     * @param cellTemplateFactory      负责解析单元格语法的工厂。
     * @param blockDirectiveHandlers   一个有序的块指令处理器列表。
     */
    public TemplateCompiler(CellTemplateFactory cellTemplateFactory, List<BlockDirectiveHandler> blockDirectiveHandlers) {
        this.cellTemplateFactory = Objects.requireNonNull(cellTemplateFactory, "CellTemplateFactory cannot be null.");
        this.blockDirectiveHandlers = Objects.requireNonNull(blockDirectiveHandlers, "BlockDirectiveHandlers cannot be null.");
    }

    /**
     * compile 方法签名，接收 TemplateStyleInfo。
     * @param sheet       要编译的模板 Sheet。
     * @param sheetStyleInfo 该 Sheet 的样式信息（包含原始合并区域）。
     * @return 预编译的模板。
     */
    public PrecompiledTemplate compile(Sheet sheet, TemplateStyleInfo sheetStyleInfo) {
        Stack<BlockBuilder> builderStack = new Stack<>();
        // 创建根 BlockBuilder 时，传入 sheetStyleInfo
        builderStack.push(new BlockBuilder(BlockType.ROOT, -1, "ROOT", this.cellTemplateFactory, sheetStyleInfo)); // 【修改】

        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            String directive = findDirectiveInRow(row);

            if (directive != null) {
                // 传递 sheetStyleInfo，因为 BlockCompilerContext 需要它来创建 BlockBuilder
                boolean handled = dispatchDirective(directive, i, builderStack, sheetStyleInfo);
                if (!handled) {
                    logger.trace("Unknown directive '{}' on row {}. Treating as a static row.", directive, i + 1);
                    builderStack.peek().addTemplateRow(row);
                }
            } else {
                builderStack.peek().addTemplateRow(row);
            }
        }

        while (builderStack.size() > 1) {
            BlockBuilder openBlock = builderStack.peek();
            logger.warn("Template has unclosed block starting on row {}, auto-closing at the end of the sheet.", openBlock.getStartRow() + 1);
            BlockBuilder finishedBlock = builderStack.pop();
            finishedBlock.setEndRow(lastRowNum);
            builderStack.peek().addChild(finishedBlock.build());
        }

        return new PrecompiledTemplate(builderStack.peek().getChildren());
    }

    /**
     * 分派指令给合适的处理器，并传递 TemplateStyleInfo。
     * 因为 BlockCompilerContext 会用它来构建 BlockBuilder。
     *
     * @return 如果指令被成功处理，返回 true；否则返回 false。
     */
    private boolean dispatchDirective(String directive, int rowNum, Stack<BlockBuilder> builderStack, TemplateStyleInfo sheetStyleInfo) { // 【修改】
        BlockCompilerContext context = new BlockCompilerContext(directive, rowNum, builderStack, this.cellTemplateFactory, sheetStyleInfo); // 【修改】
        for (BlockDirectiveHandler handler : this.blockDirectiveHandlers) {
            if (handler.canHandle(directive)) {
                handler.handle(context);
                return true;
            }
        }
        return false;
    }

    // findDirectiveInRow 方法保持不变
    private String findDirectiveInRow(Row row) {
        if (row == null) return null;
        for (Cell cell : row) {
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String text = cell.getStringCellValue();
                if (text != null) {
                    String trimmedText = text.trim();
                    if (trimmedText.startsWith("#")) {
                        return trimmedText;
                    }
                }
            }
        }
        return null;
    }
}