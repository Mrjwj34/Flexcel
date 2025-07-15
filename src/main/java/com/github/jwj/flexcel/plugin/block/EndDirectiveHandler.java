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

// File: com/github/jwj/flexcel/core/parser/syntax/block/EndDirectiveHandler.java
package com.github.jwj.flexcel.plugin.block;

import com.github.jwj.flexcel.parser.ast.block.BlockBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 负责处理 #end 指令的处理器。
 */
public class EndDirectiveHandler implements BlockDirectiveHandler {
    private static final Logger logger = LoggerFactory.getLogger(EndDirectiveHandler.class);
    private static final String DIRECTIVE = "#end";

    @Override
    public boolean canHandle(String directive) {
        return DIRECTIVE.equals(directive);
    }

    @Override
    public void handle(BlockCompilerContext context) {
        logger.debug("Found #end on row {}", context.getCurrentRowNum() + 1);
        if (context.getBuilderStack().size() > 1) { // 确保不弹出根节点
            BlockBuilder finishedBlock = context.getBuilderStack().pop();
            finishedBlock.setEndRow(context.getCurrentRowNum());
            context.getBuilderStack().peek().addChild(finishedBlock.build());
        } else {
            logger.warn("Found orphaned #end directive on row {}", context.getCurrentRowNum() + 1);
        }
    }
}