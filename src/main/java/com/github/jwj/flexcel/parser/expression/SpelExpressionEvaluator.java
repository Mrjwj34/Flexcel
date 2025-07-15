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

// File: com/github/jwj/flexcel/parser/expression/SpelExpressionEvaluator.java
package com.github.jwj.flexcel.parser.expression;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.expression.MapAccessor;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.Expression;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.PropertyAccessor;
import org.springframework.expression.common.TemplateParserContext;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.ReflectivePropertyAccessor; // 【导入正确的类】
import org.springframework.expression.spel.support.SimpleEvaluationContext;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 使用 Spring Expression Language (SpEL) 的表达式求值器实现。
 * 此版本将安全策略完全封装在内部。默认使用 SimpleEvaluationContext，
 * 它禁用了 T() 表达式等危险操作，但允许通过配置安全的 PropertyAccessor
 * 来访问 Map 和 POJO 的属性，同时支持服务方法调用。
 */
public class SpelExpressionEvaluator implements ExpressionEvaluator {

    private static final Logger logger = LoggerFactory.getLogger(SpelExpressionEvaluator.class);

    private static final ExpressionParser PARSER = new SpelExpressionParser();
    private static final TemplateParserContext TEMPLATE_PARSER_CONTEXT = new TemplateParserContext("${", "}");
    private static final Pattern SINGLE_EXPRESSION_PATTERN = Pattern.compile("^\\$\\{([^}]+)\\}$");

    private static final Map<String, Expression> expressionCache = new ConcurrentHashMap<>(256);
    private static final Map<String, Expression> templateExpressionCache = new ConcurrentHashMap<>(256);

    private final boolean unsafeOperationsEnabled;

    /**
     * 默认构造函数，运行在安全模式下。
     */
    public SpelExpressionEvaluator() {
        this(false);
    }

    /**
     * 构造函数，允许指定是否启用不安全操作。
     * @param unsafeOperationsEnabled 如果为 true，将使用 StandardEvaluationContext，允许所有反射操作。
     */
    public SpelExpressionEvaluator(boolean unsafeOperationsEnabled) {
        this.unsafeOperationsEnabled = unsafeOperationsEnabled;
    }

    // evaluateString 和 evaluate 方法保持不变...
    @Override
    public Object evaluateString(String expressionOrTemplate, Map<String, Object> contextData) {
        if (expressionOrTemplate == null || expressionOrTemplate.isEmpty()) {
            return expressionOrTemplate;
        }

        Matcher matcher = SINGLE_EXPRESSION_PATTERN.matcher(expressionOrTemplate.trim());
        if (matcher.matches()) {
            String pureExpression = matcher.group(1);
            return evaluate(pureExpression, contextData);
        }

        if (expressionOrTemplate.contains("${")) {
            try {
                Expression expr = templateExpressionCache.computeIfAbsent(expressionOrTemplate,
                        key -> PARSER.parseExpression(key, TEMPLATE_PARSER_CONTEXT));
                return expr.getValue(createContext(contextData), String.class);
            } catch (Exception e) {
                logger.warn("Template string evaluation failed: '{}'. Error: {}", expressionOrTemplate, e.getMessage());
                return expressionOrTemplate;
            }
        }
        return expressionOrTemplate;
    }

    @Override
    public Object evaluate(String expression, Map<String, Object> contextData) {
        if (expression == null || expression.trim().isEmpty()) {
            return null;
        }
        try {
            Expression expr = expressionCache.computeIfAbsent(expression, PARSER::parseExpression);
            return expr.getValue(createContext(contextData));
        } catch (Exception e) {
            logger.debug("Expression evaluation failed: '{}'. This is expected in SAFE mode for certain expressions. Error: {}", expression.trim(), e.getMessage());
            return null;
        }
    }


    /**
     * 根据构造时传入的开关，创建并返回合适的 EvaluationContext。
     */
    private EvaluationContext createContext(Map<String, Object> contextData) {
        if (this.unsafeOperationsEnabled) {
            // 危险模式：功能全开，有安全风险
            return new StandardEvaluationContext(contextData);
        } else {
            // 创建一个支持读写的 Builder，这将允许方法调用（对于我们的服务调用是必需的）
            SimpleEvaluationContext.Builder builder = new SimpleEvaluationContext
                    .Builder(new MapAccessor(), new ReflectivePropertyAccessor());
            builder.withInstanceMethods();
            builder.withRootObject(contextData);
            return builder.build();
        }
    }
}