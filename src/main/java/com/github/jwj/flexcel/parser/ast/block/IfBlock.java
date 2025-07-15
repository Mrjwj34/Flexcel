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

import com.github.jwj.flexcel.runtime.TemplateContext;
import com.github.jwj.flexcel.parser.expression.ExpressionEvaluator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class IfBlock implements TemplateBlock {
    private static final Pattern IF_PATTERN = Pattern.compile("#if\\s+(.+)");

    private final String conditionExpression;
    private final List<TemplateBlock> thenBlocks;
    private final List<TemplateBlock> elseBlocks;
    // 不再需要持有独立的求值器实例
    // private final ExpressionEvaluator expressionEvaluator;

    public IfBlock(String directive, List<TemplateBlock> thenBlocks, List<TemplateBlock> elseBlocks) {
        Matcher matcher = IF_PATTERN.matcher(directive);
        if (!matcher.find()) throw new IllegalArgumentException("Invalid #if directive: " + directive);
        this.conditionExpression = matcher.group(1).trim();
        this.thenBlocks = thenBlocks;
        this.elseBlocks = elseBlocks;
        // 【删除】不再需要初始化求值器
        // this.expressionEvaluator = new ExpressionEvaluator();
    }

    private boolean toBoolean(Object value) {
        if (value == null) return false;
        if (value instanceof Boolean) return (Boolean) value;
        if (value instanceof Number) return ((Number) value).doubleValue() != 0;
        if (value instanceof String) return !((String) value).isEmpty();
        return true;
    }

    /**
     * 方法签名改变，接收一个外部传入的求值器。
     * @param context 当前上下文
     * @param evaluator 用于求值的求值器实例
     * @return 条件的布尔结果
     */
    public boolean evaluateCondition(TemplateContext context, ExpressionEvaluator evaluator) {
        Object result = evaluator.evaluate(conditionExpression, context.getAllData());
        return toBoolean(result);
    }

    public List<TemplateBlock> getThenBlocks() {
        return thenBlocks;
    }

    public List<TemplateBlock> getElseBlocks() {
        return elseBlocks;
    }
}