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

package com.github.jwj.flexcel.parser.ast.block;

import com.github.jwj.flexcel.parser.ast.template.RowTemplate;

import java.util.List;

/**
 * 代表 #foreach 循环块。
 */
public class ForEachBlock implements TemplateBlock {
    private final String itemName;
    private final String collectionExpression;
    private final String indexName; // 通常是 "index"
    private final int startRowNo;
    private final int endRowNo;
    private final List<RowTemplate> rowTemplates;

    // 问题3：更新构造函数
    public ForEachBlock(String itemName, String collectionExpression, int startRowNo, int endRowNo, List<RowTemplate> rowTemplates) {
        this.itemName = itemName;
        this.collectionExpression = collectionExpression;
        this.indexName = "index"; // 默认索引变量名
        this.startRowNo = startRowNo;
        this.endRowNo = endRowNo;
        this.rowTemplates = rowTemplates;
    }

    public String getItemName() {
        return itemName;
    }

    public String getCollectionExpression() {
        return collectionExpression;
    }

    public String getIndexName() {
        return indexName;
    }

    public int getStartRowNo() {
        return startRowNo;
    }

    public int getEndRowNo() {
        return endRowNo;
    }

    public List<RowTemplate> getRowTemplates() {
        return rowTemplates;
    }
}