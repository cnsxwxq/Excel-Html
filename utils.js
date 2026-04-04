/**
 * ExcelUtils - 通用工具函数模块
 * @module ExcelUtils
 * @version 1.1.0-modular
 * @description 提供文件处理、格式化、HTML 转义等通用工具函数
 * 
 * 使用方式：
 * - window.ExcelUtils.formatFileSize(bytes)
 * - window.ExcelUtils.escapeHtml(text)
 * - window.ExcelUtils.formatDate(timestamp)
 */
window.ExcelUtils = (function() {
    'use strict';

    /**
     * 格式化文件大小为可读字符串
     * @param {number} bytes - 文件大小（字节）
     * @returns {string} 格式化后的字符串（如 "1.5 MB"）
     * 
     * @example
     * ExcelUtils.formatFileSize(1024)      // => "1.00 KB"
     * ExcelUtils.formatFileSize(1048576)   // => "1.00 MB"
     * ExcelUtils.formatFileSize(500)       // => "500 B"
     */
    function formatFileSize(bytes) {
        if (typeof bytes !== 'number' || isNaN(bytes)) {
            console.warn('⚠️ formatFileSize: 参数必须是数字');
            return '0 B';
        }

        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
        if (bytes < 1024 * 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
        return (bytes / (1024 * 1024 * 1024)).toFixed(2) + ' GB';
    }

    /**
     * HTML 特殊字符转义（防止 XSS）
     * @param {string} text - 需要转义的文本
     * @returns {string} 转义后的安全字符串
     * 
     * @example
     * ExcelUtils.escapeHtml('<script>alert("xss")</script>')
     * // => "&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;"
     */
    function escapeHtml(text) {
        if (text === null || text === undefined) {
            return '';
        }

        const div = document.createElement('div');
        div.textContent = String(text);
        return div.innerHTML;
    }

    /**
     * 格式化时间戳为本地时间字符串
     * @param {number|Date} timestamp - 时间戳或 Date 对象
     * @returns {string} 格式化的时间字符串
     * 
     * @example
     * ExcelUtils.formatDate(Date.now())
     * // => "2026/4/4 13:45"
     */
    function formatDate(timestamp) {
        let date;
        
        if (timestamp instanceof Date) {
            date = timestamp;
        } else if (typeof timestamp === 'number') {
            date = new Date(timestamp);
        } else {
            console.warn('⚠️ formatDate: 无效的时间戳参数');
            return '';
        }

        try {
            return date.toLocaleDateString('zh-CN') + ' ' + 
                   date.toLocaleTimeString('zh-CN', {
                       hour: '2-digit',
                       minute: '2-digit'
                   });
        } catch (err) {
            console.error('❌ formatDate 格式化失败:', err.message);
            return '';
        }
    }

    /**
     * 验证是否为有效的 File 对象
     * @param {*} obj - 待验证的对象
     * @returns {boolean} 是否为 File 对象
     */
    function isFile(obj) {
        if (!obj) return false;
        
        // 浏览器环境检查
        if (typeof File !== 'undefined' && obj instanceof File) {
            return true;
        }
        
        // 兼容性检查（File-like 对象）
        if (obj && typeof obj === 'object' &&
            typeof obj.name === 'string' &&
            typeof obj.size === 'number' &&
            typeof obj.type !== 'undefined') {
            return true;
        }
        
        return false;
    }

    /**
     * 验证是否为有效的 Excel 文件名
     * @param {string} filename - 文件名
     * @returns {boolean} 是否为 .xlsx 或 .xls 文件
     */
    function isExcelFile(filename) {
        if (typeof filename !== 'string' || !filename.trim()) {
            return false;
        }
        return /\.(xlsx|xls)$/i.test(filename);
    }

    /**
     * 深拷贝对象（用于数据隔离）
     * @param {*} obj - 要拷贝的对象
     * @returns {*} 拷贝后的对象
     */
    function deepClone(obj) {
        if (obj === null || typeof obj !== 'object') {
            return obj;
        }

        if (Array.isArray(obj)) {
            return obj.map(item => deepClone(item));
        }

        const cloned = {};
        for (let key in obj) {
            if (obj.hasOwnProperty(key)) {
                cloned[key] = deepClone(obj[key]);
            }
        }

        return cloned;
    }

    /**
     * 生成唯一 ID
     * @param {string} prefix - ID 前缀（可选）
     * @returns {string} 唯一标识符
     */
    function generateId(prefix) {
        const id = Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
        return prefix ? `${prefix}_${id}` : id;
    }

    /**
     * 防抖函数（限制高频调用）
     * @param {Function} func - 要防抖的函数
     * @param {number} wait - 等待时间（毫秒）
     * @returns {Function} 防抖后的函数
     */
    function debounce(func, wait) {
        let timeout;
        return function(...args) {
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, args), wait);
        };
    }

    /**
     * 节流函数（限制执行频率）
     * @param {Function} func - 要节流的函数
     * @param {number} limit - 时间间隔（毫秒）
     * @returns {Function} 节流后的函数
     */
    function throttle(func, limit) {
        let inThrottle;
        return function(...args) {
            if (!inThrottle) {
                func.apply(this, args);
                inThrottle = true;
                setTimeout(() => inThrottle = false, limit);
            }
        };
    }

    // 公开 API
    return {
        formatFileSize: formatFileSize,
        escapeHtml: escapeHtml,
        formatDate: formatDate,
        isFile: isFile,
        isExcelFile: isExcelFile,
        deepClone: deepClone,
        generateId: generateId,
        debounce: debounce,
        throttle: throttle,

        // 版本信息
        version: '1.1.0-modular',
        created: new Date().toISOString()
    };
})();

console.log('✅ ExcelUtils 模块已加载 v' + ExcelUtils.version);
