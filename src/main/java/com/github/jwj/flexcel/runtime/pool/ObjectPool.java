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

package com.github.jwj.flexcel.runtime.pool;

import com.github.jwj.flexcel.runtime.dto.MergeableRenderedCell;
import com.github.jwj.flexcel.runtime.dto.RenderedCell;
import com.github.jwj.flexcel.runtime.dto.RenderedRow;

/**
 * 对象池接口，定义了获取和归还渲染所需DTO对象的方法。
 * 这使得引擎可以与具体的池化策略解耦。
 */
public interface ObjectPool {

    /**
     * 从池中获取一个 RenderedRow 对象。
     * @return 一个 RenderedRow 实例。
     */
    RenderedRow acquireRow();

    /**
     * 从池中获取一个 RenderedCell 对象。
     * @return 一个 RenderedCell 实例。
     */
    RenderedCell acquireCell();

    /**
     * 从池中获取一个 MergeableRenderedCell 对象。
     * @return 一个 MergeableRenderedCell 实例。
     */
    MergeableRenderedCell acquireMergeableCell();

    /**
     * 将一个 RenderedRow 对象归还到池中。
     * @param row 要归还的对象。
     */
    void returnRow(RenderedRow row);

    /**
     * 将一个 RenderedCell 对象归还到池中。
     * @param cell 要归还的对象。
     */
    void returnCell(RenderedCell cell);

    /**
     * 将一个 MergeableRenderedCell 对象归还到池中。
     * @param cell 要归还的对象。
     */
    void returnMergeableCell(MergeableRenderedCell cell);
}