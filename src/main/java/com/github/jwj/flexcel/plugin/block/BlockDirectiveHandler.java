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

package com.github.jwj.flexcel.plugin.block;

/**
 * 块级指令处理器接口。
 * <p>
 * 该接口的实现类负责识别模板中的特定块级指令（如 #if, #foreach, #end）并执行相应的编译时操作，
 * 例如向 {@code BlockBuilder} 栈中推入新的构建器或弹出已完成的构建器。
 * 这将块指令的解析逻辑从 {@code TemplateCompiler} 的主循环中解耦出来，实现了插件化。
 * </p>
 *
 */
public interface BlockDirectiveHandler {

    /**
     * 判断当前处理器是否能够处理给定的指令。
     *
     * @param directive 从模板行中提取出的指令字符串。
     * @return 如果可以处理，则返回 {@code true}；否则返回 {@code false}。
     */
    boolean canHandle(String directive);

    /**
     * 处理指令。
     * <p>
     * 此方法在 {@link #canHandle(String)} 返回 {@code true} 时被调用。
     * 处理器在此方法内执行具体操作，如创建和管理 {@code BlockBuilder}。
     * </p>
     *
     * @param context 包含处理指令所需全部状态和依赖的上下文对象。
     */
    void handle(BlockCompilerContext context);
}