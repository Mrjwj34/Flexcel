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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 样式提取管理器。
 * 职责: 从模板Sheet中提取静态的样式信息，包括单元格样式、行高、列宽和合并区域。
 * 【重构】移除了模板模拟展开的相关逻辑（createPositionMapping），因为这部分功能已由预编译器模型取代。
 */
public class StyleMappingManager {

    private static final Logger logger = LoggerFactory.getLogger(StyleMappingManager.class);

    /**
     * 从模板Sheet中提取完整的静态样式信息。
     * @param templateSheet 模板Sheet。
     * @return 包含所有静态样式信息的 TemplateStyleInfo 对象。
     */
    public TemplateStyleInfo extractTemplateStyles(Sheet templateSheet) {
        logger.debug("Extracting styles from sheet: {}", templateSheet.getSheetName());
        TemplateStyleInfo styleInfo = new TemplateStyleInfo();

        // 提取列宽
        short maxCol = 0;
        for (Row row : templateSheet) {
            if (row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum();
            }
        }
        // 合理设置上限，防止读取过多空列
        for (int i = 0; i < Math.min(maxCol, 256); i++) {
            styleInfo.setColumnWidth(i, templateSheet.getColumnWidth(i));
        }

        // 提取行高和单元格样式
        for (Row row : templateSheet) {
            styleInfo.setRowHeight(row.getRowNum(), row.getHeightInPoints());
            for (Cell cell : row) {
                if (cell.getCellStyle() != null) {
                    styleInfo.setCellStyle(cell.getAddress().formatAsString(), cell.getCellStyle());
                }
            }
        }

        // 提取合并区域
        templateSheet.getMergedRegions().forEach(styleInfo::addMergedRegion);
        logger.debug("Style extraction complete for sheet '{}'. Found {} styles, {} merged regions.",
                templateSheet.getSheetName(), styleInfo.getAllTemplateStyles().size(), styleInfo.getMergedRegions().size());
        return styleInfo;
    }
}