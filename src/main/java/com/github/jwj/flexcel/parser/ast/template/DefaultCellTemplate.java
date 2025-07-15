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

package com.github.jwj.flexcel.parser.ast.template;

import com.github.jwj.flexcel.runtime.dto.MergeableRenderedCell;
import com.github.jwj.flexcel.runtime.dto.RenderedCell;
import com.github.jwj.flexcel.runtime.pool.ObjectPool; // 修改导入
import com.github.jwj.flexcel.runtime.TemplateContext;
import com.github.jwj.flexcel.parser.expression.ExpressionEvaluator;

import java.util.Map;

public class DefaultCellTemplate implements CellTemplate {
    private final String templateAddress;
    private final int colIndex;
    private final String expression;
    private final boolean isFormula;
    private final boolean isMergeCandidate;

    public DefaultCellTemplate(String templateAddress, int colIndex, String expression, boolean isFormula, boolean isMergeCandidate) {
        this.templateAddress = templateAddress;
        this.colIndex = colIndex;
        this.expression = expression;
        this.isFormula = isFormula;
        this.isMergeCandidate = isMergeCandidate;
    }

    @Override
    public RenderedCell produce(TemplateContext context, ObjectPool pool, Map<String, String> stringCache, ExpressionEvaluator evaluator) {
        Object finalValue;

        if (this.expression != null) {
            if (isMergeCandidate) {
                finalValue = evaluator.evaluate(this.expression, context.getAllData());
                if (finalValue instanceof String) {
                    finalValue = stringCache.computeIfAbsent((String) finalValue, k -> k);
                }
            } else {
                finalValue = evaluator.evaluateString(this.expression, context.getAllData());
            }
        } else {
            finalValue = null;
        }

        if (isMergeCandidate) {
            MergeableRenderedCell cell = pool.acquireMergeableCell();
            cell.set(templateAddress, colIndex, finalValue, false, false);
            return cell;
        } else {
            RenderedCell cell = pool.acquireCell();
            cell.set(templateAddress, colIndex, finalValue, isFormula);
            return cell;
        }
    }
}