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
import com.github.jwj.flexcel.parser.ast.template.DefaultCellTemplate;
import org.apache.poi.ss.usermodel.CellType;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 负责处理公式语法的 {@link CellSyntaxHandler} 实现。
 * 该处理器识别并解析两种形式的公式：
 * 1. 以 {@code #formula } 为前缀的自定义公式指令 (例如: {@code #formula SUM(A1:A10)}).
 * 2. 模板中原生的 Excel 公式单元格 ({@link CellType#FORMULA}).
 */
public class FormulaCellHandler implements CellSyntaxHandler {

    // 匹配 "#formula " 指令，并捕获后面的所有内容作为表达式。
    // 使用 Pattern.DOTALL 模式以允许表达式跨行。
    private static final Pattern FORMULA_DIRECTIVE_PATTERN = Pattern.compile("^#formula\\s+(.*)$", Pattern.DOTALL);

    @Override
    public boolean canHandle(String rawText) {
        // 如果原始文本不为空且以 "#formula" 开头，则此处理器可以处理
        return rawText != null && rawText.trim().startsWith("#formula");
    }

    @Override
    public CellTemplate handle(CellParseContext context) {
        String expression;
        String rawText = context.getRawStringValue(); // 假定 canHandle 已确保不为 null

        // 提取指令后的表达式
        Matcher matcher = FORMULA_DIRECTIVE_PATTERN.matcher(rawText.trim());
        if (matcher.matches()) {
            expression = matcher.group(1).trim();
        } else {
            // 作为安全保障，如果模式不匹配，则将整个文本视为表达式
            expression = rawText;
        }

        return new DefaultCellTemplate(
                context.getTemplateAddress(),
                context.getColIndex(),
                expression,
                true, // 标记为公式
                false // 不是合并候选项
        );
    }
}