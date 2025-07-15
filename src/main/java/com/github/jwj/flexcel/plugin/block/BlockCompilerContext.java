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

package com.github.jwj.flexcel.plugin.block;

import com.github.jwj.flexcel.plugin.cell.CellTemplateFactory;
import com.github.jwj.flexcel.parser.ast.block.BlockBuilder;
import com.github.jwj.flexcel.style.TemplateStyleInfo; // 【新增】

import java.util.Stack;

/**
 * 块级指令编译上下文。
 * <p>
 * 新增了对 {@link TemplateStyleInfo} 的引用，以便在创建 {@link BlockBuilder} 时传入，
 * 从而处理行级的静态合并区域。
 * </p>
 *
 */
public class BlockCompilerContext {

    private final String directive;
    private final int currentRowNum;
    private final Stack<BlockBuilder> builderStack;
    private final CellTemplateFactory cellTemplateFactory;
    private final TemplateStyleInfo sheetStyleInfo; // 【新增】

    /**
     * 构造函数，新增 sheetStyleInfo 参数。
     */
    public BlockCompilerContext(String directive, int currentRowNum, Stack<BlockBuilder> builderStack, CellTemplateFactory cellTemplateFactory, TemplateStyleInfo sheetStyleInfo) {
        this.directive = directive;
        this.currentRowNum = currentRowNum;
        this.builderStack = builderStack;
        this.cellTemplateFactory = cellTemplateFactory;
        this.sheetStyleInfo = sheetStyleInfo; // 【新增】
    }

    public String getDirective() { return directive; }
    public int getCurrentRowNum() { return currentRowNum; }
    public Stack<BlockBuilder> getBuilderStack() { return builderStack; }
    public CellTemplateFactory getCellTemplateFactory() { return cellTemplateFactory; }

    /**
     * 获取当前 Sheet 的样式信息。
     * @return TemplateStyleInfo 实例。
     */
    public TemplateStyleInfo getSheetStyleInfo() {
        return sheetStyleInfo;
    }
}