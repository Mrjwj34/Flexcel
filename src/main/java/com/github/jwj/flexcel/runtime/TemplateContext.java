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

// File: com/github/jwj/flexcel/model/TemplateContext.java
package com.github.jwj.flexcel.runtime;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * <p>
 * 模板求值上下文，负责管理模板渲染过程中的变量作用域。
 * 此版本采用父级委托模型（parent delegation model）来处理作用域，解决了变量在嵌套块之间传递的根本问题。
 * </p>
 * <p>
 * 当在当前上下文中查找变量时，如果未找到，它会自动向上委托给父上下文进行查找，
 * 形成一个作用域链。这确保了在子块中可以访问父块定义的变量，并且同级块之间对共享上下文的修改是可见的。
 * </p>
 *
 */
public class TemplateContext {

    // 指向父级上下文的引用，为 null 表示是根上下文
    private final TemplateContext parent;

    // 存储引擎级的共享数据，如 services, global context, 和用户传入的主数据模型。
    // 这个对象在整个上下文中是共享的，而不是逐级复制。
    private final Map<String, Object> sharedData;

    // 仅存储当前作用域定义的局部变量。
    private final Map<String, Object> localScopedVariables;

    /**
     * 构造一个根上下文。
     *
     * @param data          用户传入的主数据模型。
     * @param services      注册的服务。
     * @param globalContext 全局上下文变量。
     */
    public TemplateContext(Map<String, Object> data, Map<String, Object> services, Map<String, Object> globalContext) {
        this.parent = null;
        this.localScopedVariables = new ConcurrentHashMap<>();

        // 将所有共享数据合并到一个 map 中，以简化查找
        this.sharedData = new HashMap<>();
        if (globalContext != null) this.sharedData.putAll(globalContext);
        if (data != null) this.sharedData.putAll(data);
        if (services != null) this.sharedData.putAll(services);
        // 添加一些默认的有用变量
        this.sharedData.putIfAbsent("printDate", new java.util.Date());
    }

    /**
     * 构造一个子上下文。
     *
     * @param parentContext 父级上下文。
     */
    public TemplateContext(TemplateContext parentContext) {
        if (parentContext == null) {
            throw new IllegalArgumentException("Parent context cannot be null for a child context.");
        }
        this.parent = parentContext;
        this.localScopedVariables = new ConcurrentHashMap<>();
        // 子上下文共享父级的共享数据引用
        this.sharedData = parentContext.sharedData;
    }

    /**
     * 在当前作用域设置一个局部变量。
     *
     * @param name  变量名。
     * @param value 变量值。
     */
    public void setVariable(String name, Object value) {
        if (value != null) {
            this.localScopedVariables.put(name, value);
        } else {
            this.localScopedVariables.remove(name);
        }
    }

    /**
     * 根据作用域链查找变量。
     * <p>
     * 查找顺序:
     * 1. 当前上下文的局部变量 (localScopedVariables)。
     * 2. (递归) 父上下文的局部变量。
     * 3. 根上下文的共享数据 (sharedData)。
     * </p>
     *
     * @param name 变量名。
     * @return 找到的变量值，如果都找不到则返回 null。
     */
    public Object getVariable(String name) {
        // 1. 在当前局部作用域查找
        if (this.localScopedVariables.containsKey(name)) {
            return this.localScopedVariables.get(name);
        }

        // 2. 如果有父级，向上委托
        if (this.parent != null) {
            return this.parent.getVariable(name);
        }

        // 3. 如果已经到达根作用域，在共享数据中查找
        return this.sharedData.get(name);
    }

    /**
     * 获取一个包含所有可见变量的 Map，用于传递给表达式求值器。
     * <p>
     * 这个 Map 是动态构建的，它反映了当前作用域链上的所有变量。
     * 子作用域的变量会覆盖父作用域的同名变量（词法作用域）。
     * </p>
     *
     * @return 一个包含所有可访问变量的扁平化 Map。
     */
    public Map<String, Object> getAllData() {
        Map<String, Object> allData = new HashMap<>();

        // 从根节点开始，逐级向下合并作用域变量，实现子级覆盖父级
        if (this.parent != null) {
            allData.putAll(this.parent.getAllData());
        } else {
            // 如果是根节点，先加载共享数据
            allData.putAll(this.sharedData);
        }

        // 最后，用当前作用域的局部变量覆盖
        allData.putAll(this.localScopedVariables);

        return allData;
    }
}