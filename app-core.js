/**
 * ExcelApp - 应用核心模块
 * @module ExcelApp
 * @version 1.1.0-modular
 * @description 模块注册中心、生命周期管理、全局配置
 * 
 * 功能：
 * - 模块注册与获取
 * - 统一初始化入口
 * - 事件总线
 * - 配置管理
 */
window.ExcelApp = (function() {
    'use strict';

    // ==================== 私有状态 ====================
    
    /**
     * 已注册的模块
     * @type {Object.<string, Object>}
     */
    const modules = {};

    /**
     * 模块就绪回调队列
     * @type {Array.<Function>}
     */
    const readyCallbacks = [];

    /**
     * 是否已初始化
     * @type {boolean}
     */
    let isInitialized = false;

    /**
     * 当前正在执行的模块（用于检测循环依赖）
     * @type {string|null}
     */
    let currentExecutingModule = null;

    // ==================== 事件系统 ====================

    /**
     * 事件监听器存储
     * @type {Object.<string, Array.<Function>>}
     */
    const eventListeners = {};

    /**
     * 注册事件监听器
     * @param {string} event - 事件名称
     * @param {Function} callback - 回调函数
     * @returns {Function} 取消订阅的函数
     */
    function on(event, callback) {
        if (typeof callback !== 'function') {
            console.error('❌ ExcelApp.on: 回调必须是函数');
            return () => {};
        }

        if (!eventListeners[event]) {
            eventListeners[event] = [];
        }

        eventListeners[event].push(callback);

        // 返回取消订阅函数
        return function() {
            const index = eventListeners[event].indexOf(callback);
            if (index > -1) {
                eventListeners[event].splice(index, 1);
            }
        };
    }

    /**
     * 触发事件
     * @param {string} event - 事件名称
     * @param {*} data - 事件数据
     */
    function emit(event, data) {
        console.log(`📢 [Event] ${event}`, data);

        const listeners = eventListeners[event];
        if (!listeners || listeners.length === 0) {
            return;
        }

        listeners.forEach(function(callback) {
            try {
                callback(data);
            } catch (err) {
                console.error(`❌ [Event:${event}] 监听器执行错误:`, err);
            }
        });
    }

    /**
     * 移除所有指定事件的监听器
     * @param {string} event - 事件名称（可选，不传则清除全部）
     */
    function off(event) {
        if (event) {
            delete eventListeners[event];
        } else {
            for (let key in eventListeners) {
                delete eventListeners[key];
            }
        }
    }

    // ==================== 模块管理 ====================

    /**
     * 注册模块
     * @param {string} name - 模块名称（唯一标识）
     * @param {Object} moduleInstance - 模块实例
     * @throws {Error} 如果参数无效
     */
    function register(name, moduleInstance) {
        // 参数验证
        if (typeof name !== 'string' || !name.trim()) {
            throw new Error('ExcelApp.register: 模块名称必须是有效字符串');
        }

        if (!moduleInstance || typeof moduleInstance !== 'object') {
            throw new Error('ExcelApp.register: 模块实例必须是对象');
        }

        // 检查重复注册
        if (modules[name]) {
            console.warn(`⚠️ 模块 "${name}" 已被注册，将被覆盖`);
        }

        // 注册模块
        modules[name] = moduleInstance;
        
        console.log(`✅ 模块已注册: ${name}`);

        // 触发注册事件
        emit('module:registered', { name: name, module: moduleInstance });
    }

    /**
     * 获取模块实例
     * @param {string} name - 模块名称
     * @returns {Object|null} 模块实例，如果不存在返回 null
     */
    function getModule(name) {
        // 循环依赖检测
        if (currentExecutingModule && currentExecutingModule === name) {
            console.error(`❌ 检测到循环调用: 模块 "${name}" 正在尝试调用自身`);
            console.trace();
            return null;
        }

        if (!modules[name]) {
            console.warn(`⚠️ 模块未找到: ${name}`);
            console.log('可用模块:', Object.keys(modules));
            return null;
        }

        return modules[name];
    }

    /**
     * 获取所有已注册模块的列表
     * @returns {Array.<string>} 模块名称数组
     */
    function getModuleList() {
        return Object.keys(modules);
    }

    /**
     * 检查模块是否已注册
     * @param {string} name - 模块名称
     * @returns {boolean}
     */
    function hasModule(name) {
        return !!modules[name];
    }

    // ==================== 初始化管理 ====================

    /**
     * 标记当前正在执行的模块（用于循环检测）
     * @param {string|null} moduleName - 模块名称
     */
    function setExecutingModule(moduleName) {
        currentExecutingModule = moduleName;
    }

    /**
     * 注册就绪回调
     * @param {Function} callback - 就绪时调用的函数
     */
    function onReady(callback) {
        if (typeof callback !== 'function') {
            console.error('❌ ExcelApp.onReady: 参数必须是函数');
            return;
        }

        if (isInitialized) {
            // 已经初始化，立即执行
            try {
                callback();
            } catch (err) {
                console.error('❌ 就绪回调执行错误:', err);
            }
        } else {
            readyCallbacks.push(callback);
        }
    }

    /**
     * 通知应用已就绪
     */
    function notifyReady() {
        isInitialized = true;
        
        console.log('🎉 所有模块已就绪');
        console.log('📦 已加载模块:', getModuleList());

        // 执行所有就绪回调
        readyCallbacks.forEach(function(callback) {
            try {
                callback();
            } catch (err) {
                console.error('❌ 就绪回调执行错误:', err);
            }
        });

        // 清空回调队列
        readyCallbacks.length = 0;

        // 触发就绪事件
        emit('app:ready', { 
            timestamp: new Date().toISOString(),
            modules: getModuleList()
        });
    }

    /**
     * 检查是否已初始化
     * @returns {boolean}
     */
    function isReady() {
        return isInitialized;
    }

    // ==================== 配置管理 ====================

    /**
     * 全局配置
     * @type {Object}
     */
    const config = {
        debug: window.location.hostname === 'localhost',
        version: '1.1.0-modular',
        appName: '向心链',
        maxFileSize: 100 * 1024 * 1024,  // 100MB
        chunkSize: 100,
        requestTimeout: 30000  // 30秒
    };

    /**
     * 获取配置项
     * @param {string} key - 配置键名
     * @returns {*} 配置值
     */
    function getConfig(key) {
        if (key) {
            return config[key];
        }
        return Object.assign({}, config);
    }

    /**
     * 设置配置项
     * @param {string} key - 配置键名
     * @param {*} value - 配置值
     */
    function setConfig(key, value) {
        config[key] = value;
        console.log(`⚙️ 配置更新: ${key} =`, value);
        emit('config:changed', { key: key, value: value });
    }

    // ==================== 工具方法 ====================

    /**
     * 安全执行函数（带错误处理）
     * @param {Function} fn - 要执行的函数
     * @param {string} context - 执行上下文描述
     * @returns {*} 函数返回值或 undefined
     */
    function safeExecute(fn, context) {
        if (typeof fn !== 'function') {
            console.error(`❌ safeExecute: ${context} - 参数必须是函数`);
            return undefined;
        }

        try {
            setExecutingModule(context);
            const result = fn();
            setExecutingModule(null);
            return result;
        } catch (err) {
            setExecutingModule(null);
            console.error(`❌ [${context}] 执行错误:`, err);
            emit('error', { context: context, error: err.message });
            return undefined;
        }
    }

    /**
     * 异步安全执行
     * @param {Function} fn - 要执行的异步函数
     * @param {string} context - 执行上下文描述
     * @returns {Promise} Promise 对象
     */
    async function safeExecuteAsync(fn, context) {
        if (typeof fn !== 'function') {
            console.error(`❌ safeExecuteAsync: ${context} - 参数必须是函数`);
            return Promise.reject(new Error('Invalid function'));
        }

        try {
            setExecutingModule(context);
            const result = await fn();
            setExecutingModule(null);
            return result;
        } catch (err) {
            setExecutingModule(null);
            console.error(`❌ [${context}] 异步执行错误:`, err);
            emit('error', { context: context, error: err.message });
            throw err;
        }
    }

    /**
     * 获取应用信息
     * @returns {Object} 应用基本信息
     */
    function getAppInfo() {
        return {
            name: config.appName,
            version: config.version,
            modules: getModuleList(),
            initialized: isInitialized,
            uptime: Date.now(),
            environment: config.debug ? 'development' : 'production'
        };
    }

    // ==================== 公开 API ====================
    
    return {
        // 模块管理
        register: register,
        getModule: getModule,
        getModuleList: getModuleList,
        hasModule: hasModule,

        // 事件系统
        on: on,
        off: off,
        emit: emit,

        // 初始化
        onReady: onReady,
        notifyReady: notifyReady,
        isReady: isReady,
        setExecutingModule: setExecutingModule,

        // 配置
        getConfig: getConfig,
        setConfig: setConfig,

        // 工具
        safeExecute: safeExecute,
        safeExecuteAsync: safeExecuteAsync,
        getAppInfo: getAppInfo
    };
})();

console.log('✅ ExcelApp 核心模块已加载 v' + ExcelApp.getConfig('version'));
