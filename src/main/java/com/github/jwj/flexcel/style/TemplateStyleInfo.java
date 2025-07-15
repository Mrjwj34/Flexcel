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

package com.github.jwj.flexcel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * 模板样式信息
 * 存储模板中的样式信息和特殊占位符位置
 */
public class TemplateStyleInfo {
    
    // 单元格位置到样式的映射
    private final Map<String, CellStyle> cellStyles = new HashMap<>();
    
    // 特殊占位符单元格位置列表
    private final Set<String> templateCellPositions = new HashSet<>();
    
    // 合并单元格区域
    private final List<CellRangeAddress> mergedRegions = new ArrayList<>();
    
    // 行高信息
    private final Map<Integer, Float> rowHeights = new HashMap<>();
    
    // 列宽信息
    private final Map<Integer, Integer> columnWidths = new HashMap<>();
    
    /**
     * 设置单元格样式
     */
    public void setCellStyle(String cellPosition, CellStyle style) {
        cellStyles.put(cellPosition, style);
    }
    
    /**
     * 获取单元格样式
     */
    public CellStyle getCellStyle(String cellPosition) {
        return cellStyles.get(cellPosition);
    }
    
    /**
     * 获取所有单元格样式
     */
    public Map<String, CellStyle> getAllCellStyles() {
        return new HashMap<>(cellStyles);
    }
    
    /**
     * 添加模板单元格位置
     */
    public void addTemplateCellPosition(String cellPosition) {
        templateCellPositions.add(cellPosition);
    }
    
    /**
     * 获取所有模板单元格位置
     */
    public Set<String> getTemplateCellPositions() {
        return new HashSet<>(templateCellPositions);
    }
    
    /**
     * 检查是否是模板单元格
     */
    public boolean isTemplateCell(String cellPosition) {
        return templateCellPositions.contains(cellPosition);
    }
    
    /**
     * 添加合并单元格区域
     */
    public void addMergedRegion(CellRangeAddress mergedRegion) {
        mergedRegions.add(mergedRegion);
    }
    
    /**
     * 获取所有合并单元格区域
     */
    public List<CellRangeAddress> getMergedRegions() {
        return new ArrayList<>(mergedRegions);
    }
    
    /**
     * 设置行高
     */
    public void setRowHeight(int rowIndex, float height) {
        rowHeights.put(rowIndex, height);
    }
    
    /**
     * 获取行高
     */
    public Float getRowHeight(int rowIndex) {
        return rowHeights.get(rowIndex);
    }
    
    /**
     * 获取所有行高信息
     */
    public Map<Integer, Float> getAllRowHeights() {
        return new HashMap<>(rowHeights);
    }
    
    /**
     * 设置列宽
     */
    public void setColumnWidth(int columnIndex, int width) {
        columnWidths.put(columnIndex, width);
    }
    
    /**
     * 获取列宽
     */
    public Integer getColumnWidth(int columnIndex) {
        return columnWidths.get(columnIndex);
    }
    
    /**
     * 获取所有列宽信息
     */
    public Map<Integer, Integer> getAllColumnWidths() {
        return new HashMap<>(columnWidths);
    }
    
    /**
     * 清空所有样式信息
     */
    public void clear() {
        cellStyles.clear();
        templateCellPositions.clear();
        mergedRegions.clear();
        rowHeights.clear();
        columnWidths.clear();
    }
    /**
     * 从合并区域列表中移除一个特定的合并区域。
     * 主要用于编译阶段，当某个合并区域被识别为属于某个 RowTemplate 时，
     * 将其从 Sheet 级别的静态合并区域中移除，以避免重复处理。
     *
     * @param mergedRegion 要移除的合并区域。
     * @return 如果成功移除则返回 true，否则返回 false。
     */
    public boolean removeMergedRegion(CellRangeAddress mergedRegion) {
        return this.mergedRegions.remove(mergedRegion);
    }
    /**
     * 获取样式信息统计
     */
    public String getStatistics() {
        return String.format("样式信息统计: 单元格样式=%d, 模板单元格=%d, 合并区域=%d, 行高=%d, 列宽=%d",
                cellStyles.size(), templateCellPositions.size(), mergedRegions.size(), 
                rowHeights.size(), columnWidths.size());
    }
    /**
     * 获取模板中所有使用到的、唯一的CellStyle对象集合。
     * 使用HashSet可以自动去重，因为不同的单元格可能引用同一个CellStyle对象。
     * 这是性能优化的关键，确保我们只为每个独特的样式创建一次新样式。
     *
     * @return 一个包含所有唯一CellStyle对象的集合。
     */
    public Collection<CellStyle> getAllTemplateStyles() {
        return new HashSet<>(this.cellStyles.values());
    }
}
