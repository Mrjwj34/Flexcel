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

// File: com/github/jwj/flexcel/core/parser/CellTemplateFactory.java
package com.github.jwj.flexcel.plugin.cell;

import com.github.jwj.flexcel.parser.ast.template.CellTemplate;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;
import java.util.Objects;

/**
 * {@link CellTemplate} 的工厂类，是单元格级语法插件系统的核心协调者。
 * <p>
 * 该工厂持有一个 {@link CellSyntaxHandler} 的列表。当被要求创建一个 {@code CellTemplate} 时，
 * 它会按顺序遍历这些处理器，使用第一个声明能够处理（通过 {@code canHandle} 方法）
 * 当前单元格内容的处理器来实际创建 {@code CellTemplate}。
 * </p>
 * <p>
 * 这实现了责任链设计模式，使得单元格的解析逻辑可以被轻松扩展和定制。
 * 默认的 {@link DefaultCellHandler} 应始终位于处理链的末端，以处理所有其他处理器未处理的情况。
 * </p>
 *
 */
public class CellTemplateFactory {

    private final List<CellSyntaxHandler> handlers;

    /**
     * 使用一个已定义的处理器列表构造工厂。
     *
     * @param handlers {@link CellSyntaxHandler} 的列表。该列表应已按期望的优先级排序。
     *                 通常，最具体的处理器在前，最通用的（如 {@link DefaultCellHandler}）在后。
     */
    public CellTemplateFactory(List<CellSyntaxHandler> handlers) {
        Objects.requireNonNull(handlers, "Handlers list cannot be null.");
        this.handlers = handlers;
    }

    /**
     * 根据给定的单元格信息，创建并返回一个 {@link CellTemplate}。
     * <p>
     * 方法内部会构建一个 {@link CellParseContext}，然后遍历已注册的处理器，
     * 找到第一个可以处理该上下文的处理器，并委托其完成创建工作。
     * </p>
     *
     * @param cell     原始的 POI Cell 对象，可能为 null。
     * @param rowIndex 单元格所在的行索引。
     * @param colIndex 单元格所在的列索引。
     * @return 创建的 {@link CellTemplate} 实例。
     * @throws IllegalStateException 如果没有找到任何可以处理该单元格的处理器
     *         （这种情况理论上不应发生，因为 {@link DefaultCellHandler} 应该总是在链的末尾）。
     */
    public CellTemplate create(Cell cell, int rowIndex, int colIndex) {
        CellParseContext context = new CellParseContext(cell, rowIndex, colIndex);
        String rawText = context.getRawStringValue();

        for (CellSyntaxHandler handler : handlers) {
            // 注意：对于非字符串单元格，rawText 为 null。
            // canHandle 的实现需要考虑到这一点。
            if (handler.canHandle(rawText)) {
                return handler.handle(context);
            }
        }

        // 理论上，由于 DefaultCellHandler 的存在，永远不会到达这里。
        // 但为了代码健壮性，抛出异常。
        throw new IllegalStateException(
                "No suitable CellSyntaxHandler found for cell at " + context.getTemplateAddress()
                        + ". A DefaultCellHandler should always be present."
        );
    }
}