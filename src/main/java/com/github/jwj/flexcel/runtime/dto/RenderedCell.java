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

/**
 * 代表一个已准备好被渲染到 Excel 中的单元格。
 * 新增了对自定义渲染逻辑的支持，允许插件定义超越简单值设置的复杂渲染行为（如插入图片）。
 */
public class RenderedCell {
    public String templateAddress;
    public int colIndex;
    public Object value;
    public boolean isFormula = false;

    /**
     * 一个可执行的自定义渲染动作。
     * 当此字段不为 null 时，消费者线程将执行此动作，而不是进行标准的值设置。
     * Consumer<Cell> 接收一个 Cell 参数，允许插件对这个 Cell 进行任意操作。
     */
    public Consumer<Cell> customRenderer;

    public RenderedCell() {}

    /**
     * 设置标准单元格数据。
     */
    public void set(String templateAddress, int colIndex, Object value, boolean isFormula) {
        this.templateAddress = templateAddress;
        this.colIndex = colIndex;
        this.value = value;
        this.isFormula = isFormula;
        this.customRenderer = null; // 确保清除自定义渲染器
    }

    /**
     * 设置一个带有自定义渲染逻辑的单元格。
     *
     * @param templateAddress 模板地址
     * @param colIndex        列索引
     * @param customRenderer  一个 Consumer，它接收被创建的 Cell 对象并对其执行操作。
     */
    public void setCustom(String templateAddress, int colIndex, Consumer<Cell> customRenderer) {
        this.templateAddress = templateAddress;
        this.colIndex = colIndex;
        this.value = null;
        this.isFormula = false;
        this.customRenderer = customRenderer;
    }

    public void reset() {
        this.templateAddress = null;
        this.colIndex = -1;
        this.value = null;
        this.isFormula = false;
        this.customRenderer = null; // 【新增】在重置时清除渲染器
    }
}