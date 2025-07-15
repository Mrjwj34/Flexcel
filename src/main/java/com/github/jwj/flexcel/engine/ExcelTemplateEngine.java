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

package com.github.jwj.flexcel.engine;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

/**
 * Excel模板引擎接口
 * 支持excelutils的所有特殊占位符功能
 */
public interface ExcelTemplateEngine {

    /**
     * 使用模板生成Excel文件
     * @param templateStream 模板文件输入流
     * @param data 数据模型
     * @param outputStream 输出流
     */
    void process(InputStream templateStream, Map<String, Object> data, OutputStream outputStream);

    /**
     * 注册自定义函数服务
     * 推荐使用Builder模式进行注册
     * @param serviceName 服务名称
     * @param service 服务实例
     */
    void registerService(String serviceName, Object service);

    /**
     * 设置表达式上下文
     * @param context 上下文数据
     */
    void setContext(Map<String, Object> context);
}
