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
 * 一个“无操作”的对象池实现。
 * 当对象池被禁用时使用此类。
 * 'acquire' 方法总是创建新对象，'return' 方法不执行任何操作。
 */
public class NoOpObjectPool implements ObjectPool {

    @Override
    public RenderedRow acquireRow() {
        return new RenderedRow();
    }

    @Override
    public RenderedCell acquireCell() {
        return new RenderedCell();
    }

    @Override
    public MergeableRenderedCell acquireMergeableCell() {
        return new MergeableRenderedCell();
    }

    @Override
    public void returnRow(RenderedRow row) {
        // 不执行任何操作
    }

    @Override
    public void returnCell(RenderedCell cell) {
        // 不执行任何操作
    }

    @Override
    public void returnMergeableCell(MergeableRenderedCell cell) {
        // 不执行任何操作
    }
}