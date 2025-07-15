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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * 默认的 {@link CellSyntaxHandler} 实现，作为处理链的末端。
 * <p>
 * 它负责处理所有其他特定语法处理器不处理的单元格，包括：
 * <ul>
 *   <li>普通文本</li>
 *   <li>标准表达式 (如 {@code ${user.name}})</li>
 *   <li>数字、布尔值、原生日期等非字符串类型单元格</li>
 *   <li>空单元格</li>
 * </ul>
 * </p>
 */
public class DefaultCellHandler implements CellSyntaxHandler {

    /**
     * 默认处理器作为责任链的末端，总是能够处理任何输入。
     *
     * @param rawText 单元格的原始字符串内容。
     * @return 总是返回 {@code true}。
     */
    @Override
    public boolean canHandle(String rawText) {
        return true;
    }

    @Override
    public CellTemplate handle(CellParseContext context) {
        Cell cell = context.getCell();
        // 如果单元格为空，创建一个表示空内容的模板
        if (cell == null) {
            return new DefaultCellTemplate(context.getTemplateAddress(), context.getColIndex(), null, false, false);
        }

        Object expression = null;
        boolean isFormula = false;
        CellType cellType = cell.getCellType();

        // 根据单元格类型提取其内容作为表达式
        switch (cellType) {
            case STRING:
                // 对于字符串，直接使用其内容作为表达式或模板字符串
                expression = cell.getStringCellValue();
                break;
            case FORMULA:
                // 对于原生公式，使用公式字符串作为表达式
                expression = cell.getCellFormula();
                isFormula = true;
                break;
            case NUMERIC:
                // 对于数字，转换为 BigDecimal 的字符串形式以保证精度
                // POI 的日期也存储为数字类型，需特殊判断
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 如果是日期格式，直接在求值时让 POI 返回 Date 对象即可，此处记录原始数值
                    expression = cell.getNumericCellValue();
                } else {
                    expression = new java.math.BigDecimal(cell.getNumericCellValue()).toPlainString();
                }
                break;
            case BOOLEAN:
                // 对于布尔值，转换为字符串
                expression = String.valueOf(cell.getBooleanCellValue());
                break;
            case BLANK:
            case ERROR:
            default:
                // 对于空白、错误或其他类型，表达式为 null
                expression = null;
                break;
        }

        // 注意：这里的 expression 是 Object 类型，因为它可能持有数字或日期值，
        // 而不仅仅是字符串。DefaultCellTemplate 的构造函数和 produce 方法需要能处理这种情况。
        // 但为了统一，我们将非字符串转为字符串，因为最终求值器输入的是字符串。
        String finalExpression = (expression != null) ? String.valueOf(expression) : null;

        return new DefaultCellTemplate(
                context.getTemplateAddress(),
                context.getColIndex(),
                finalExpression, // 传递最终的表达式字符串
                isFormula,
                false
        );
    }
}