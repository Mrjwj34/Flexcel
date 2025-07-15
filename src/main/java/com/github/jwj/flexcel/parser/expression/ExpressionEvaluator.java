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

package com.github.jwj.flexcel.parser.expression;

import java.util.Map;

/**
 * 表达式求值器接口。
 * 定义了将字符串表达式或模板根据给定的上下文数据求值的契约。
 * 这使得引擎可以解耦具体的表达式语言实现。
 */
public interface ExpressionEvaluator {

    /**
     * 对一个可能包含表达式的字符串模板进行求值。
     * - 如果字符串是纯表达式（如 "${user.name}"），则返回表达式的实际结果类型。
     * - 如果是混合模板（如 "你好, ${user.name}!"），则返回渲染后的字符串。
     * - 如果不包含表达式，则原样返回。
     *
     * @param expressionOrTemplate 待求值的表达式或模板字符串。
     * @param contextData          求值所需的上下文数据。
     * @return 求值后的结果。
     */
    Object evaluateString(String expressionOrTemplate, Map<String, Object> contextData);

    /**
     * 对一个纯表达式进行求值。
     *
     * @param expression  待求值的纯表达式字符串 (不含 "${...}")。
     * @param contextData 求值所需的上下文数据。
     * @return 求值后的结果。
     */
    Object evaluate(String expression, Map<String, Object> contextData);
}