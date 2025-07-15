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
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;

/**
 * 基于 {@link BlockingQueue} 的默认对象池实现。
 */
public class DefaultObjectPool implements ObjectPool {
    private final BlockingQueue<RenderedRow> rowPool;
    private final BlockingQueue<RenderedCell> cellPool;
    private final BlockingQueue<MergeableRenderedCell> mergeableCellPool;

    /**
     * 使用指定的容量构造对象池。
     * @param rowCapacity RenderedRow 池的容量。
     * @param cellCapacity RenderedCell 池的容量。
     * @param mergeableCellCapacity MergeableRenderedCell 池的容量。
     */
    public DefaultObjectPool(int rowCapacity, int cellCapacity, int mergeableCellCapacity) {
        this.rowPool = new LinkedBlockingQueue<>(rowCapacity);
        this.cellPool = new LinkedBlockingQueue<>(cellCapacity);
        this.mergeableCellPool = new LinkedBlockingQueue<>(mergeableCellCapacity);
    }

    @Override
    public RenderedRow acquireRow() {
        RenderedRow row = rowPool.poll();
        return (row != null) ? row : new RenderedRow();
    }

    @Override
    public RenderedCell acquireCell() {
        RenderedCell cell = cellPool.poll();
        return (cell != null) ? cell : new RenderedCell();
    }

    @Override
    public MergeableRenderedCell acquireMergeableCell() {
        MergeableRenderedCell cell = mergeableCellPool.poll();
        return (cell != null) ? cell : new MergeableRenderedCell();
    }

    @Override
    public void returnRow(RenderedRow row) {
        if (row == null || row == RenderedRow.POISON_PILL) return;

        for (RenderedCell cell : row.cells) {
            if (cell instanceof MergeableRenderedCell) {
                returnMergeableCell((MergeableRenderedCell) cell);
            } else {
                returnCell(cell);
            }
        }
        row.reset();
        rowPool.offer(row);
    }

    @Override
    public void returnCell(RenderedCell cell) {
        if (cell != null) {
            cell.reset();
            cellPool.offer(cell);
        }
    }

    @Override
    public void returnMergeableCell(MergeableRenderedCell cell) {
        if (cell != null) {
            cell.reset();
            mergeableCellPool.offer(cell);
        }
    }
}