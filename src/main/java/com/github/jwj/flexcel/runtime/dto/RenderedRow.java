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

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class RenderedRow {
    public static final RenderedRow POISON_PILL = new RenderedRow();

    public int templateRowNum;
    public List<RenderedCell> cells;
    // 用于存储与此行模板关联的静态合并区域
    private List<CellRangeAddress> staticMergedRegions;

    public RenderedRow() {
        this.cells = new ArrayList<>(16);
        // 初始化列表
        this.staticMergedRegions = null;
    }

    /**
     * 【新增】设置此行关联的静态合并区域。
     * @param mergedRegions 合并区域列表。
     */
    public void setStaticMergedRegions(List<CellRangeAddress> mergedRegions) {
        this.staticMergedRegions = mergedRegions;
    }

    /**
     * 【新增】获取此行关联的静态合并区域。
     * @return 合并区域列表，可能为 null。
     */
    public List<CellRangeAddress> getStaticMergedRegions() {
        return staticMergedRegions;
    }

    /**
     * 【修改】reset 方法，用于对象池回收。
     */
    public void reset() {
        this.templateRowNum = -1;
        // 清空列列表，但内部的 cell 对象由调用者（对象池）负责归还
        this.cells.clear();
        // 【新增】重置合并区域信息
        this.staticMergedRegions = null;
    }
}