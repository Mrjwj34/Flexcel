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

package com.github.jwj.flexcel.plugin.cell;

import com.github.jwj.flexcel.parser.ast.template.CellTemplate;

/**
 * 单元格级语法处理器接口。
 * 该接口的实现类负责解析模板中单个单元格的内容，识别特定的语法（如 #formula, ${!var}），
 * 并将其转换为一个 {@link CellTemplate} 对象。
 * 这种设计将单元格的解析逻辑从 {@code BlockBuilder} 中解耦出来，实现了插件化的扩展。
 */
public interface CellSyntaxHandler {

    /**
     * 判断当前处理器是否能够处理给定的单元格内容。
     *
     * @param rawText 单元格的原始字符串内容。可能为 null。
     * @return 如果可以处理，则返回 {@code true}；否则返回 {@code false}。
     */
    boolean canHandle(String rawText);

    /**
     * 处理单元格内容，并创建一个相应的 {@link CellTemplate}。
     * <p>
     * 只有在 {@link #canHandle(String)} 返回 {@code true} 时，此方法才应该被调用。
     * </p>
     *
     * @param context 包含解析所需全部信息的上下文对象，如单元格地址、索引、原始值等。
     * @return 根据单元格内容创建的 {@link CellTemplate} 实例。
     */
    CellTemplate handle(CellParseContext context);
}