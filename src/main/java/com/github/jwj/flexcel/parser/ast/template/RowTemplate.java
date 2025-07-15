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

// File: com/github/jwj/flexcel/parser/precompile/template/RowTemplate.java
package com.github.jwj.flexcel.parser.ast.template;

import com.github.jwj.flexcel.runtime.dto.RenderedRow;
import com.github.jwj.flexcel.runtime.pool.ObjectPool;
import com.github.jwj.flexcel.runtime.TemplateContext;
import com.github.jwj.flexcel.parser.expression.ExpressionEvaluator;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Map;

public class RowTemplate {
    private final int templateRowNum;
    private final List<CellTemplate> cellTemplates;
    // 存储与此行模板关联的静态合并区域
    private final List<CellRangeAddress> staticMergedRegions;

    /**
     * 构造函数，增加合并区域参数。
     *
     * @param templateRowNum      模板行号
     * @param cellTemplates       单元格模板列表
     * @param staticMergedRegions 与此行关联的合并区域列表
     */
    public RowTemplate(int templateRowNum, List<CellTemplate> cellTemplates, List<CellRangeAddress> staticMergedRegions) {
        this.templateRowNum = templateRowNum;
        this.cellTemplates = cellTemplates;
        this.staticMergedRegions = staticMergedRegions;
    }

    public int getTemplateRowNum() {
        return templateRowNum;
    }

    /**
     * produce 方法，将合并信息传递给 RenderedRow。
     */
    public RenderedRow produce(TemplateContext context, ObjectPool pool, Map<String, String> stringCache, ExpressionEvaluator evaluator) {
        RenderedRow row = pool.acquireRow();
        row.templateRowNum = this.templateRowNum;
        // 【新增】传递合并信息
        row.setStaticMergedRegions(this.staticMergedRegions);

        for (CellTemplate cellTemplate : this.cellTemplates) {
            row.cells.add(cellTemplate.produce(context, pool, stringCache, evaluator));
        }
        return row;
    }
}