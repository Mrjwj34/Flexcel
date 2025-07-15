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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellAddress;

/**
 * 单元格解析上下文。
 * 这是一个数据传输对象 (DTO)，封装了在解析单个单元格时所需的所有信息。
 * 它被传递给 {@link CellSyntaxHandler} 的 {@code handle} 方法，为处理器提供执行所需的数据。
 */
public final class CellParseContext {

    private final Cell cell; // 可能为 null
    private final int rowIndex;
    private final int colIndex;
    private final String templateAddress;

    /**
     * 构造一个新的解析上下文。
     *
     * @param cell     原始的 POI Cell 对象，可能为 null。
     * @param rowIndex 单元格所在的行索引。
     * @param colIndex 单元格所在的列索引。
     */
    public CellParseContext(Cell cell, int rowIndex, int colIndex) {
        this.cell = cell;
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        // 预先计算并缓存地址，避免重复计算
        this.templateAddress = new CellAddress(rowIndex, colIndex).formatAsString();
    }

    /**
     * 获取原始的 POI Cell 对象。
     * @return Cell 对象，如果模板中对应的位置为空，则返回 null。
     */
    public Cell getCell() {
        return cell;
    }

    /**
     * 获取单元格所在的行索引。
     * @return 行索引。
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /**
     * 获取单元格所在的列索引。
     * @return 列索引。
     */
    public int getColIndex() {
        return colIndex;
    }

    /**
     * 获取单元格的模板地址，例如 "A1"。
     * @return 格式化后的单元格地址字符串。
     */
    public String getTemplateAddress() {
        return templateAddress;
    }

    /**
     * 方便地获取单元格的原始字符串值。
     * @return 如果单元格类型为 STRING，则返回其值；否则返回 null。
     */
    public String getRawStringValue() {
        if (cell != null && cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return null;
    }

    /**
     * 方便地获取单元格的类型。
     * @return 如果单元格存在，返回其类型；否则返回 null。
     */
    public CellType getCellType() {
        return (cell != null) ? cell.getCellType() : null;
    }
}