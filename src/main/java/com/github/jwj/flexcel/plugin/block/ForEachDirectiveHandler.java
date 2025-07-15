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

import com.github.jwj.flexcel.parser.ast.block.BlockBuilder;
import com.github.jwj.flexcel.parser.ast.block.BlockType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 负责处理 #foreach 指令的处理器。
 */
public class ForEachDirectiveHandler implements BlockDirectiveHandler {
    private static final Logger logger = LoggerFactory.getLogger(ForEachDirectiveHandler.class);
    private static final Pattern FOREACH_PATTERN = Pattern.compile("^#foreach\\s+(\\w+)\\s+in\\s+\\$\\{([^}]+)\\}$");

    @Override
    public boolean canHandle(String directive) {
        return FOREACH_PATTERN.matcher(directive).matches();
    }

    @Override
    public void handle(BlockCompilerContext context) {
        logger.debug("Found #foreach on row {}", context.getCurrentRowNum() + 1);
        Matcher matcher = FOREACH_PATTERN.matcher(context.getDirective());
        if (matcher.matches()) {
            BlockBuilder foreachBuilder = new BlockBuilder(
                    BlockType.FOREACH,
                    context.getCurrentRowNum(),
                    context.getDirective(),
                    context.getCellTemplateFactory(),
                    context.getSheetStyleInfo()
            );
            foreachBuilder.setForeachDetails(matcher.group(1), matcher.group(2));
            context.getBuilderStack().push(foreachBuilder);
        }
    }
}