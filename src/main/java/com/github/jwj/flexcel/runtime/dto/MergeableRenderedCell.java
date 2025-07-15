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

package com.github.jwj.flexcel.runtime.dto;

import org.apache.poi.ss.usermodel.Cell;
import java.util.function.Consumer;

public class MergeableRenderedCell extends RenderedCell {
    public boolean shouldMerge;

    public MergeableRenderedCell() {
        super();
    }

    // 重载 set 方法，确保 customRenderer 被正确处理
    public void set(String templateAddress, int colIndex, Object value, boolean isFormula, boolean shouldMerge) {
        super.set(templateAddress, colIndex, value, isFormula);
        this.shouldMerge = shouldMerge;
    }

    // 为可合并单元格添加自定义渲染支持是不常见的，但为了API完整性可以提供。
    // 注意：自定义渲染和纵向合并可能行为冲突，使用时需谨慎。
    @Override
    public void setCustom(String templateAddress, int colIndex, Consumer<Cell> customRenderer) {
        super.setCustom(templateAddress, colIndex, customRenderer);
        this.shouldMerge = false; // 自定义渲染时，禁用合并行为
    }

    @Override
    public void reset() {
        super.reset();
        this.shouldMerge = false;
    }
}