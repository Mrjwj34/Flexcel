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

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 负责处理纵向合并语法的 {@link CellSyntaxHandler} 实现。
 * 该处理器识别并解析形如 {@code ${!variable}} 的语法，
 * 该语法用于在循环渲染时，当连续行的该单元格值相同时自动进行纵向合并。
 */
public class MergeCellHandler implements CellSyntaxHandler {

    // 匹配 ${!variable} 格式，并捕获其中的变量名 'variable'
    private static final Pattern MERGE_VAR_PATTERN = Pattern.compile("^\\$\\{!([^}]+)\\}$");

    @Override
    public boolean canHandle(String rawText) {
        // 如果原始文本不为空且符合 ${!...} 格式，则此处理器可以处理
        return rawText != null && MERGE_VAR_PATTERN.matcher(rawText.trim()).matches();
    }

    @Override
    public CellTemplate handle(CellParseContext context) {
        String rawText = context.getRawStringValue();
        Matcher matcher = MERGE_VAR_PATTERN.matcher(rawText.trim());

        // canHandle 已确保了此处必定匹配成功
        matcher.find();
        // 提取 '!' 后面的纯表达式，例如 "user.name"
        String expression = matcher.group(1);

        return new DefaultCellTemplate(
                context.getTemplateAddress(),
                context.getColIndex(),
                expression,
                false, // 不是公式
                true   // 标记为合并候选项
        );
    }
}