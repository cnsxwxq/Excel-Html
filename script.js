document.addEventListener('DOMContentLoaded', function() {
    const excelFile = document.getElementById('excelFile');
    const convertBtn = document.getElementById('convertBtn');
    const htmlResult = document.getElementById('htmlResult');
    const copyBtn = document.getElementById('copyBtn');
    const fileUploadWrapper = document.querySelector('.file-upload-wrapper');

    const ConvertModule = {
        isProcessing: false,
        startTime: null,
        totalRows: 0,
        processedRows: 0,
        chunkSize: 100,

        showProgress: function() {
            const section = document.getElementById('convertProgressSection');
            section.style.display = 'block';
            section.classList.remove('success');
            document.getElementById('convertProgressFill').style.width = '0%';
            document.getElementById('convertProgressPercent').textContent = '0%';
            document.getElementById('convertProgressText').textContent = '准备中...';
            document.getElementById('convertProgressTitle').textContent = '正在处理...';
            document.getElementById('convertProgressDetails').textContent = '';
        },

        updateProgress: function(percent, text, detail) {
            document.getElementById('convertProgressFill').style.width = percent + '%';
            document.getElementById('convertProgressPercent').textContent = percent + '%';
            document.getElementById('convertProgressText').textContent = text;
            if (detail) {
                document.getElementById('convertProgressDetails').textContent = detail;
            }
        },

        hideProgress: function(success) {
            const section = document.getElementById('convertProgressSection');
            if (success) {
                section.classList.add('success');
                document.getElementById('convertProgressTitle').textContent = '处理完成！';
                document.getElementById('convertProgressIcon').textContent = '✅';
            } else {
                section.style.display = 'none';
            }
        },

        resetConvertState: function() {
            const section = document.getElementById('convertProgressSection');
            if (section) {
                section.style.display = 'none';
                section.classList.remove('success');
            }
            
            convertBtn.disabled = false;
            convertBtn.innerHTML = '<span class="btn-icon">✨</span>转换为 HTML';
            
            this.isProcessing = false;
        },

        convertAsync: function(file) {
            const self = this;
            
            if (this.isProcessing) {
                return;
            }

            this.isProcessing = true;
            this.startTime = Date.now();
            convertBtn.disabled = true;
            convertBtn.innerHTML = '<span class="btn-icon">⏳</span>处理中...';

            this.showProgress();
            this.updateProgress(0, '正在读取文件...', '文件大小: ' + formatFileSize(file.size));

            const reader = new FileReader();

            reader.onload = function(e) {
                self.updateProgress(10, '正在解析Excel文件...', '这可能需要几秒钟');
                
                requestAnimationFrame(function() {
                    self.parseExcel(e.target.result, file);
                });
            };

            reader.onerror = function() {
                self.isProcessing = false;
                convertBtn.disabled = false;
                convertBtn.innerHTML = '<span class="btn-icon">✨</span>转换为 HTML';
                self.hideProgress(false);
                alert('文件读取失败，请重试');
            };

            reader.readAsArrayBuffer(file);
        },

        parseExcel: function(arrayBuffer, file) {
            const self = this;
            
            try {
                const data = new Uint8Array(arrayBuffer);
                this.updateProgress(20, '正在解析数据结构...', '解析中...');
                
                requestAnimationFrame(function() {
                    try {
                        const workbook = XLSX.read(data, { type: 'array' });
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        
                        self.updateProgress(30, '正在提取数据...', '工作表: ' + firstSheetName);
                        
                        requestAnimationFrame(function() {
                            self.processWorksheet(worksheet, file.name);
                        });
                    } catch (err) {
                        self.handleError('Excel解析失败: ' + err.message);
                    }
                });
            } catch (err) {
                this.handleError('文件处理失败: ' + err.message);
            }
        },

        processWorksheet: function(worksheet, fileName) {
            const self = this;
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            this.totalRows = jsonData.length;
            
            this.updateProgress(40, '正在分析表格结构...', '共 ' + this.totalRows + ' 行数据');

            const originalHeaders = jsonData[0];
            const linkColumnIndexes = [];
            const imageColumnIndexes = [];
            const hideColumnIndexes = [];
            const hideColumnNames = ['拍卖平台', '标的类型', '省份', '城市'];

            originalHeaders.forEach(function(header, index) {
                if (header) {
                    if (header.includes('链接')) {
                        linkColumnIndexes.push(index);
                    }
                    if (header.includes('图片')) {
                        imageColumnIndexes.push(index);
                    }
                    if (hideColumnNames.includes(header.trim())) {
                        hideColumnIndexes.push(index);
                    }
                }
            });

            this.updateProgress(45, '正在生成HTML结构...', '处理 ' + originalHeaders.length + ' 列');

            requestAnimationFrame(function() {
                self.generateHTML(worksheet, jsonData, originalHeaders, hideColumnIndexes, fileName);
            });
        },

        generateHTML: function(worksheet, jsonData, originalHeaders, hideColumnIndexes, fileName) {
            const self = this;
            
            let html = XLSX.utils.sheet_to_html(worksheet, {
                header: '<tr style="background-color: #4CAF50; color: white;">',
                footer: '</tr>',
                tableClass: 'hybrid-table'
            });

            html = html.replace('<table>', '<table class="hybrid-table">');

            this.updateProgress(55, '正在处理DOM结构...', '解析HTML');

            const parser = new DOMParser();
            const doc = parser.parseFromString(html, 'text/html');
            const table = doc.querySelector('table');

            if (!table) {
                this.handleError('HTML生成失败');
                return;
            }

            const rows = table.querySelectorAll('tr');
            this.processRowsChunked(rows, hideColumnIndexes, doc, fileName, table);
        },

        processRowsChunked: function(rows, hideColumnIndexes, doc, fileName, table) {
            const self = this;
            const totalRows = rows.length;
            let currentRow = 0;
            let headers = [];

            const headerRow = rows[0];
            if (hideColumnIndexes.length > 0) {
                const headerCells = headerRow.querySelectorAll('td, th');
                for (let j = hideColumnIndexes.length - 1; j >= 0; j--) {
                    const index = hideColumnIndexes[j];
                    if (headerCells[index]) {
                        headerCells[index].remove();
                    }
                }
            }

            const headerCells = headerRow.querySelectorAll('th, td');
            headers = Array.from(headerCells).map(function(cell) {
                return cell.textContent.trim();
            });

            this.updateProgress(60, '正在处理数据行...', '0 / ' + (totalRows - 1) + ' 行');

            function processChunk() {
                const endIndex = Math.min(currentRow + self.chunkSize, totalRows);

                for (let i = Math.max(1, currentRow); i < endIndex; i++) {
                    const row = rows[i];
                    if (!row) continue;

                    const cells = row.querySelectorAll('td');

                    if (hideColumnIndexes.length > 0) {
                        for (let j = hideColumnIndexes.length - 1; j >= 0; j--) {
                            const index = hideColumnIndexes[j];
                            if (cells[index]) {
                                cells[index].remove();
                            }
                        }
                    }

                    const updatedCells = row.querySelectorAll('td');
                    const minLen = Math.min(headers.length, updatedCells.length);

                    for (let j = 0; j < minLen; j++) {
                        const headerText = headers[j];
                        if (headerText && headerText !== '') {
                            updatedCells[j].setAttribute('data-label', headerText);
                        }

                        if (headers[j] && headers[j].includes('链接')) {
                            const cell = updatedCells[j];
                            let url = cell.getAttribute('data-v') || cell.textContent.trim();
                            if (url && url.match(/^https?:\/\//)) {
                                const a = doc.createElement('a');
                                a.href = url;
                                a.target = '_blank';
                                a.textContent = '详情页 →';
                                cell.innerHTML = '';
                                cell.appendChild(a);
                            }
                        }

                        if (headers[j] && headers[j].includes('图片')) {
                            const cell = updatedCells[j];
                            let url = cell.getAttribute('data-v') || cell.textContent.trim();
                            if (url && url.match(/^https?:\/\//)) {
                                const a = doc.createElement('a');
                                a.href = url;
                                a.target = '_blank';
                                a.textContent = '查看图片 →';
                                cell.innerHTML = '';
                                cell.appendChild(a);
                            }
                        }
                    }
                }

                currentRow = endIndex;
                const progress = 60 + Math.round((currentRow / totalRows) * 35);
                self.updateProgress(progress, '正在处理数据行...', 
                    (currentRow - 1) + ' / ' + (totalRows - 1) + ' 行');

                if (currentRow < totalRows) {
                    requestAnimationFrame(processChunk);
                } else {
                    self.finalizeHTML(table, doc, fileName);
                }
            }

            requestAnimationFrame(processChunk);
        },

        finalizeHTML: function(table, doc, fileName) {
            const self = this;
            
            console.log('[DEBUG] finalizeHTML 开始, 时间:', Date.now() - self.startTime + 'ms');
            this.updateProgress(95, '正在生成最终HTML...', '即将完成');

            try {
                console.log('[DEBUG] 步骤1: 获取table.outerHTML');
                const html = table.outerHTML;
                
                if (!html || html.trim() === '') {
                    console.error('[ERROR] HTML内容为空');
                    self.handleError('HTML内容生成失败：表格内容为空');
                    return;
                }

                console.log('[DEBUG] 步骤2: HTML长度=' + html.length);
                this.updateProgress(96, '正在应用样式...', 'HTML长度: ' + (html.length / 1024).toFixed(1) + 'KB');

                setTimeout(function() {
                    try {
                        console.log('[DEBUG] 步骤3: 开始构建样式化HTML');
                        self.updateProgress(97, '正在构建完整HTML...', '处理中...');
                        
                        const styledHtml = self.createStyledHTML(html);
                        console.log('[DEBUG] 步骤4: 样式化HTML完成, 长度=' + styledHtml.length);

                        self.updateProgress(98, '正在渲染结果...', '即将完成');
                        
                        setTimeout(function() {
                            try {
                                console.log('[DEBUG] 步骤5: 更新DOM');
                                self.updateProgress(99, '正在更新界面...', '最后一步');
                                
                                htmlResult.innerHTML = '<div class="table-wrapper">' + html + '</div>';
                                htmlResult.classList.add('has-content');
                                
                                console.log('[DEBUG] 步骤6: 存储数据到dataset');
                                htmlResult.dataset.html = styledHtml;

                                console.log('[DEBUG] 步骤7: 完成转换, 总耗时:', Date.now() - self.startTime + 'ms');
                                self.updateProgress(100, '处理完成！', '耗时: ' + ((Date.now() - self.startTime) / 1000).toFixed(2) + 's');

                                self.isProcessing = false;
                                convertBtn.disabled = true;
                                convertBtn.innerHTML = '<span class="btn-icon">✅</span>转换完成，请上传新文件';
                                
                                self.hideProgress(true);
                            } catch (renderErr) {
                                console.error('[ERROR] 渲染阶段错误:', renderErr);
                                self.handleError('界面更新失败: ' + renderErr.message);
                            }
                        }, 0);
                    } catch (styleErr) {
                        console.error('[ERROR] 样式化阶段错误:', styleErr);
                        self.handleError('HTML构建失败: ' + styleErr.message);
                    }
                }, 0);
            } catch (htmlErr) {
                console.error('[ERROR] HTML生成阶段错误:', htmlErr);
                self.handleError('HTML生成失败: ' + htmlErr.message);
            }
        },

        createStyledHTML: function(html) {
            return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>向心链——让数据取之于民，用之于民</title>
    <style>
        :root {
            --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            --secondary-gradient: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            --accent-color: #667eea;
            --text-dark: #1a202c;
            --text-light: #718096;
            --bg-color: #f7fafc;
            --card-bg: #ffffff;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.1);
            --shadow-md: 0 4px 6px rgba(0,0,0,0.1);
            --shadow-lg: 0 10px 25px rgba(0,0,0,0.15);
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
            background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
            min-height: 100vh;
            padding: 20px 15px;
            color: #e0e0e0;
            line-height: 1.6;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(10px);
            padding: 25px;
            border-radius: 20px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        h1 {
            text-align: center;
            background: linear-gradient(90deg, #00c6ff, #0072ff, #00c6ff);
            background-size: 200% auto;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            animation: gradient 3s ease infinite;
            font-size: 1.8em;
            margin-bottom: 25px;
            font-weight: 700;
            text-shadow: 0 0 30px rgba(0, 198, 255, 0.3);
            letter-spacing: 1px;
            line-height: 1.4;
        }

        @keyframes gradient {
            0%, 100% { background-position: 0% center; }
            50% { background-position: 100% center; }
        }

        .search-box {
            margin-bottom: 25px;
            padding: 20px;
            background: rgba(255, 255, 255, 0.08);
            border-radius: 15px;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        .search-box label {
            display: block;
            margin-bottom: 10px;
            font-weight: 600;
            color: #00c6ff;
            font-size: 15px;
        }

        .search-box input {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid rgba(0, 198, 255, 0.3);
            border-radius: 8px;
            font-size: 15px;
            background: rgba(255, 255, 255, 0.1);
            color: #e0e0e0;
            transition: all 0.3s ease;
            margin-bottom: 12px;
        }

        .search-box input:focus {
            outline: none;
            border-color: #00c6ff;
            box-shadow: 0 0 15px rgba(0, 198, 255, 0.4);
            background: rgba(255, 255, 255, 0.15);
        }

        .search-box input::placeholder {
            color: rgba(224, 224, 224, 0.5);
        }

        .search-box .button-group {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .search-box button {
            flex: 1;
            min-width: 100px;
            padding: 12px 20px;
            background: linear-gradient(135deg, #00c6ff, #0072ff);
            color: white;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 15px;
            font-weight: 600;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(0, 198, 255, 0.3);
        }

        .search-box button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0, 198, 255, 0.5);
        }

        .search-box button:active {
            transform: translateY(0);
        }

        .search-box .hint {
            margin-top: 12px;
            font-size: 13px;
            color: rgba(224, 224, 224, 0.6);
        }

        .table-wrapper {
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            border-radius: 10px;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        .hybrid-table {
            width: 100%;
            border-collapse: collapse;
        }

        .hybrid-table th {
            background: linear-gradient(135deg, #00c6ff, #0072ff);
            color: white;
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 16px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            white-space: nowrap;
        }

        .hybrid-table td {
            padding: 14px 12px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.08);
            text-align: left;
            color: #e0e0e0;
            transition: all 0.3s ease;
            font-size: 16px;
        }

        .hybrid-table tr {
            background: rgba(255, 255, 255, 0.03);
        }

        .hybrid-table tr:nth-child(even) {
            background: rgba(255, 255, 255, 0.05);
        }

        .hybrid-table tr:hover {
            background: rgba(0, 198, 255, 0.15);
            transform: scale(1.005);
        }

        .hybrid-table a {
            color: #00c6ff;
            text-decoration: none;
            transition: all 0.3s ease;
            font-weight: 500;
        }

        .hybrid-table a:hover {
            text-decoration: underline;
            color: #0072ff;
            text-shadow: 0 0 10px rgba(0, 198, 255, 0.5);
        }

        .hidden {
            display: none !important;
        }

        @media (max-width: 768px) {
            .hybrid-table thead {
                display: none;
            }
            
            .hybrid-table tr:first-child {
                display: none !important;
            }

            .hybrid-table tr {
                display: grid;
                grid-template-columns: 1fr;
                gap: 0;
                margin-bottom: 15px;
                border: 2px solid rgba(255, 255, 255, 0.1);
                border-radius: 10px;
                padding: 0;
                background: rgba(255, 255, 255, 0.05);
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
                position: relative;
                overflow: hidden;
            }

            .hybrid-table tr::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 5px;
                height: 100%;
                background: linear-gradient(135deg, #00c6ff, #0072ff);
            }

            .hybrid-table td[data-label="标题"] {
                order: -1;
                padding: 15px 12px;
                text-align: left;
                font-weight: 700;
                color: #ffffff;
                font-size: 24px;
                border-bottom: 2px solid #00c6ff;
                background: linear-gradient(135deg, rgba(0, 198, 255, 0.2), rgba(0, 114, 255, 0.1));
                line-height: 1.4;
                text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
                word-break: break-word;
            }

            .hybrid-table td[data-label="标题"]::before {
                display: none;
            }

            .hybrid-table td {
                display: block;
                padding: 10px 12px;
                border-bottom: 1px solid rgba(255, 255, 255, 0.08);
                text-align: center;
                position: relative;
                font-size: 21px;
                min-width: 0;
                word-break: break-word;
            }

            .hybrid-table td:last-child {
                border-bottom: none;
            }

            .hybrid-table td[data-label]::before {
                content: attr(data-label) "：";
                display: block;
                position: static;
                font-weight: 600;
                color: #00c6ff;
                text-align: left;
                font-size: 18px;
                margin-bottom: 4px;
                white-space: normal;
                opacity: 0.9;
            }

            .hybrid-table td[data-label="标题"]::before {
                display: none;
            }

            .hybrid-table td[data-label="标题链接"] a {
                display: inline-block;
                width: 145px;
                padding: 6px 4px;
                background: linear-gradient(135deg, #00c6ff, #0072ff);
                color: white;
                border-radius: 6px;
                text-decoration: none;
                font-weight: 600;
                font-size: 21px;
                text-align: center;
                margin-top: 4px;
                transition: all 0.3s ease;
            }

            .hybrid-table td[data-label="标题链接"] a:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 15px rgba(0, 198, 255, 0.4);
            }

            .hybrid-table td[data-label="图片"] a {
                display: inline-block;
                width: 145px;
                padding: 6px 4px;
                background: linear-gradient(135deg, #f093fb, #f5576c);
                color: white;
                border-radius: 6px;
                text-decoration: none;
                font-weight: 600;
                font-size: 21px;
                text-align: center;
                margin-top: 4px;
                transition: all 0.3s ease;
            }

            .hybrid-table td[data-label="图片"] a:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 15px rgba(240, 147, 251, 0.4);
            }
        }

        ::-webkit-scrollbar {
            width: 10px;
            height: 10px;
        }

        ::-webkit-scrollbar-track {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 5px;
        }

        ::-webkit-scrollbar-thumb {
            background: linear-gradient(135deg, #00c6ff, #0072ff);
            border-radius: 5px;
        }

        ::-webkit-scrollbar-thumb:hover {
            background: linear-gradient(135deg, #0072ff, #00c6ff);
        }

        @media (max-width: 1024px) {
            .container {
                padding: 20px;
            }

            h1 {
                font-size: 1.6em;
            }
        }

        @media (max-width: 768px) {
            body {
                padding: 15px 10px;
            }

            .container {
                padding: 15px;
                border-radius: 15px;
            }

            h1 {
                font-size: 1.4em;
                margin-bottom: 20px;
            }

            .search-box {
                padding: 15px;
                margin-bottom: 20px;
            }

            .search-box .button-group {
                flex-direction: column;
            }

            .search-box button {
                width: 100%;
            }
        }

        @media (max-width: 480px) {
            h1 {
                font-size: 1.2em;
            }

            .search-box input {
                font-size: 14px;
                padding: 10px 12px;
            }
        }

        @media (max-height: 500px) and (orientation: landscape) {
            .container {
                padding: 15px;
            }

            h1 {
                font-size: 1.3em;
                margin-bottom: 15px;
            }

            .search-box {
                padding: 10px;
                margin-bottom: 15px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>✨ 向心链——让数据取之于民，用之于民 ✨</h1>
        <div class="search-box">
            <label for="searchInput">🔍 搜索：</label>
            <input type="text" id="searchInput" placeholder="输入关键词，多个用逗号分隔">
            <div class="button-group">
                <button onclick="searchTable()">搜索</button>
                <button onclick="clearSearch()">清除</button>
            </div>
            <div class="hint">💡 提示：可输入多个关键词，用逗号分隔，如：北京，上海</div>
        </div>
        <div class="table-wrapper">
            ${html}
        </div>
    </div>
    <script>
        function searchTable() {
            const input = document.getElementById('searchInput');
            const filter = input.value.toLowerCase();
            const keywords = filter.split(',').map(k => k.trim()).filter(k => k);
            const table = document.querySelector('.hybrid-table');
            const tr = table.querySelectorAll('tr');
            
            for (let i = 1; i < tr.length; i++) {
                const row = tr[i];
                let showRow = false;
                
                if (keywords.length === 0) {
                    showRow = true;
                } else {
                    const rowText = row.textContent.toLowerCase();
                    for (let j = 0; j < keywords.length; j++) {
                        if (rowText.includes(keywords[j])) {
                            showRow = true;
                            break;
                        }
                    }
                }
                
                if (showRow) {
                    row.classList.remove('hidden');
                } else {
                    row.classList.add('hidden');
                }
            }
        }
        
        function clearSearch() {
            document.getElementById('searchInput').value = '';
            const table = document.querySelector('.hybrid-table');
            const tr = table.querySelectorAll('tr');
            for (let i = 1; i < tr.length; i++) {
                tr[i].classList.remove('hidden');
            }
        }
    <\/script>
</body>
</html>`;
        },

        handleError: function(message) {
            this.isProcessing = false;
            convertBtn.disabled = false;
            convertBtn.innerHTML = '<span class="btn-icon">✨</span>转换为 HTML';
            this.hideProgress(false);
            alert('转换失败: ' + message);
        }
    };

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        document.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    const fileInputContainer = document.querySelector('.file-input-container');

    function showFileSuccess(file) {
        const existingMessage = document.querySelector('.file-success-message');
        if (existingMessage) {
            existingMessage.remove();
        }

        const fileInfo = document.createElement('div');
        fileInfo.className = 'file-success-message';
        fileInfo.innerHTML = `
            <div class="file-info-content">
                <div class="file-icon">✅</div>
                <div class="file-details">
                    <div class="file-name">${file.name}</div>
                    <div class="file-meta">
                        <span>大小：${formatFileSize(file.size)}</span>
                        <span>类型：${file.type || 'Excel 文件'}</span>
                        <span>上传时间：${new Date().toLocaleTimeString('zh-CN')}</span>
                    </div>
                </div>
            </div>
        `;

        fileInputContainer.parentNode.insertBefore(fileInfo, fileInputContainer.nextSibling);
    }

    function clearFileSuccessMessage() {
        const existingMessage = document.querySelector('.file-success-message');
        if (existingMessage) {
            existingMessage.remove();
        }
    }

    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
    }

    fileInputContainer.addEventListener('dragenter', () => {
        fileUploadWrapper.classList.add('dragover');
    });

    fileInputContainer.addEventListener('dragover', () => {
        fileUploadWrapper.classList.add('dragover');
    });

    fileInputContainer.addEventListener('dragleave', () => {
        fileUploadWrapper.classList.remove('dragover');
    });

    excelFile.addEventListener('change', function() {
        if (excelFile.files.length && excelFile.files[0].name.match(/\.(xlsx|xls)$/)) {
            ConvertModule.resetConvertState();
            showFileSuccess(excelFile.files[0]);
        }
    });

    fileInputContainer.addEventListener('drop', (e) => {
        fileUploadWrapper.classList.remove('dragover');
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length && files[0].name.match(/\.(xlsx|xls)$/)) {
            excelFile.files = files;
            fileUploadWrapper.style.borderColor = '#43e97b';
            ConvertModule.resetConvertState();
            showFileSuccess(files[0]);
            setTimeout(() => {
                fileUploadWrapper.style.borderColor = '';
            }, 2000);
        } else {
            alert('请上传 Excel 文件（.xlsx 或 .xls 格式）');
        }
    });

    convertBtn.addEventListener('click', function() {
        if (!excelFile.files.length) {
            alert('请选择一个 Excel 文件');
            return;
        }

        const file = excelFile.files[0];
        ConvertModule.convertAsync(file);
    });

    copyBtn.addEventListener('click', function() {
        const htmlContent = htmlResult.dataset.html;
        if (!htmlContent) {
            alert('请先转换 Excel 文件');
            return;
        }
        
        // 直接保存，使用默认参数
        const fileName = 'excel-data';
        const format = 'html';
        const encoding = 'utf-8';
        const includeStyles = true;
        const includeSearch = true;
        
        let finalHtml = htmlContent;
        
        if (!includeStyles || !includeSearch) {
            finalHtml = customizeHtml(htmlContent, includeStyles, includeSearch);
        }
        
        if (encoding === 'gbk') {
            finalHtml = finalHtml.replace(/<meta charset="UTF-8">/g, '<meta charset="GBK">');
        }
        
        const fullFileName = fileName + '.' + format;
        
        performSave(finalHtml, fullFileName, format, encoding);
    });

    function openSaveModal() {
        const saveModal = document.getElementById('saveModal');
        const browserHintText = document.getElementById('browserHintText');
        
        if (window.showSaveFilePicker) {
            browserHintText.textContent = '点击保存后可选择保存位置';
        } else {
            browserHintText.textContent = '文件将保存到浏览器默认下载目录';
        }
        
        updateSavePreviewPath();
        saveModal.style.display = 'flex';
    }

    function closeSaveModal() {
        const saveModal = document.getElementById('saveModal');
        saveModal.style.display = 'none';
    }

    function updateSavePreviewPath() {
        const fileName = document.getElementById('saveFileNameInput').value.trim() || 'excel-data';
        const format = document.getElementById('saveFileFormat').value;
        const previewPath = document.getElementById('savePreviewPath');
        previewPath.textContent = fileName + '.' + format;
    }

    document.getElementById('saveFileNameInput').addEventListener('input', updateSavePreviewPath);
    document.getElementById('saveFileFormat').addEventListener('change', updateSavePreviewPath);

    document.getElementById('closeSaveModal').addEventListener('click', closeSaveModal);
    document.getElementById('saveModalCancel').addEventListener('click', closeSaveModal);

    document.getElementById('saveModal').addEventListener('click', function(e) {
        if (e.target === this) {
            closeSaveModal();
        }
    });

    document.getElementById('saveModalConfirm').addEventListener('click', function() {
        const htmlContent = htmlResult.dataset.html;
        if (!htmlContent) {
            alert('请先转换 Excel 文件');
            closeSaveModal();
            return;
        }

        const fileName = document.getElementById('saveFileNameInput').value.trim() || 'excel-data';
        const format = document.getElementById('saveFileFormat').value;
        const encoding = document.getElementById('saveEncoding').value;
        const includeStyles = document.getElementById('includeStyles').checked;
        const includeSearch = document.getElementById('includeSearch').checked;
        
        let finalHtml = htmlContent;
        
        if (!includeStyles || !includeSearch) {
            finalHtml = customizeHtml(htmlContent, includeStyles, includeSearch);
        }
        
        if (encoding === 'gbk') {
            finalHtml = finalHtml.replace(/<meta charset="UTF-8">/g, '<meta charset="GBK">');
        }
        
        const fullFileName = fileName + '.' + format;
        
        performSave(finalHtml, fullFileName, format, encoding);
    });

    function customizeHtml(html, includeStyles, includeSearch) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        
        if (!includeStyles) {
            const styleTag = doc.querySelector('style');
            if (styleTag) {
                styleTag.remove();
            }
        }
        
        if (!includeSearch) {
            const searchBox = doc.querySelector('.search-box');
            if (searchBox) {
                searchBox.remove();
            }
            
            const scripts = doc.querySelectorAll('script');
            scripts.forEach(function(script) {
                if (script.textContent.includes('searchTable') || script.textContent.includes('clearSearch')) {
                    script.remove();
                }
            });
        }
        
        return doc.documentElement.outerHTML;
    }

    function performSave(htmlContent, fileName, format, encoding) {
        if (window.showSaveFilePicker) {
            const mimeType = 'text/html';
            const options = {
                suggestedName: fileName,
                types: [{
                    description: 'HTML 文件',
                    accept: {
                        [mimeType]: ['.' + format]
                    }
                }],
                excludeAcceptAllOption: false
            };

            let saveHandle = null;
            const self = this;
            
            window.showSaveFilePicker(options)
                .then(function(handle) {
                    saveHandle = handle;
                    return handle.createWritable();
                })
                .then(function(writable) {
                    const blob = new Blob([htmlContent], { type: mimeType + ';charset=' + encoding });
                    
                    return writable.write(blob).then(function() {
                        writable.close();
                        closeSaveModal();
                        showSaveSuccess(saveHandle.name, false);
                    });
                })
                .catch(function(err) {
                    if (err.name === 'AbortError') {
                        return;
                    }
                    closeSaveModal();
                    showSaveError(err.message);
                });
        } else {
            try {
                const blob = new Blob([htmlContent], { type: 'text/html;charset=' + encoding });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = fileName;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                closeSaveModal();
                showSaveSuccess(fileName, true);
            } catch (err) {
                closeSaveModal();
                showSaveError(err.message);
            }
        }
    }

    function showSaveSuccess(fileName, isDownload) {
        const message = isDownload 
            ? '文件 "' + fileName + '" 已保存到下载目录'
            : '文件 "' + fileName + '" 已保存成功';
        
        const existingMessage = document.querySelector('.success-message');
        if (existingMessage) {
            existingMessage.remove();
        }
        
        const successMessage = document.createElement('div');
        successMessage.className = 'success-message';
        successMessage.textContent = '✅ ' + message;
        htmlResult.parentNode.appendChild(successMessage);
        
        setTimeout(function() {
            successMessage.style.opacity = '0';
            setTimeout(function() {
                successMessage.remove();
            }, 300);
        }, 3000);
    }

    function showSaveError(errorMessage) {
        const existingMessage = document.querySelector('.error-message');
        if (existingMessage) {
            existingMessage.remove();
        }
        
        const errorMessageDiv = document.createElement('div');
        errorMessageDiv.className = 'error-message';
        errorMessageDiv.style.cssText = 'background: linear-gradient(135deg, #fca5a5 0%, #f87171 100%); color: #7f1d1d; padding: 15px 20px; border-radius: 10px; margin-top: 15px; text-align: center; font-weight: 600; animation: slideDown 0.3s ease-out;';
        errorMessageDiv.textContent = '❌ 保存失败：' + errorMessage;
        htmlResult.parentNode.appendChild(errorMessageDiv);
        
        setTimeout(function() {
            errorMessageDiv.style.opacity = '0';
            setTimeout(function() {
                errorMessageDiv.remove();
            }, 300);
        }, 5000);
    }
});

(function() {
    'use strict';

    const MergeModule = {
        files: [],
        fileData: [],
        mergedData: null,
        headers: [],
        validationResult: null,
        validationPassed: false,
        mergeCompleted: false,
        currentPage: 1,
        pageSize: 50,
        filteredData: null,
        sortColumn: null,
        sortDirection: 'asc',
        logs: [],
        isProcessing: false,
        startTime: null,

        init: function() {
            this.bindEvents();
            this.initLogPanel();
        },

        bindEvents: function() {
            const self = this;

            document.querySelectorAll('.tab-btn').forEach(function(btn) {
                btn.addEventListener('click', function() {
                    self.switchTab(this.dataset.tab);
                });
            });

            const mergeDropZone = document.getElementById('mergeDropZone');
            const mergeFiles = document.getElementById('mergeFiles');

            mergeDropZone.addEventListener('dragenter', function(e) {
                this.classList.add('dragover');
            });

            mergeDropZone.addEventListener('dragover', function(e) {
                e.preventDefault();
                this.classList.add('dragover');
            });

            mergeDropZone.addEventListener('dragleave', function(e) {
                this.classList.remove('dragover');
            });

            mergeDropZone.addEventListener('drop', function(e) {
                e.preventDefault();
                this.classList.remove('dragover');
                const files = Array.from(e.dataTransfer.files).filter(function(f) {
                    return f.name.match(/\.(xlsx|xls)$/i);
                });
                if (files.length > 0) {
                    self.addFiles(files);
                } else {
                    self.showError('文件格式错误', '请上传 Excel 文件（.xlsx 或 .xls 格式）');
                }
            });

            mergeFiles.addEventListener('change', function() {
                const files = Array.from(this.files);
                if (files.length > 0) {
                    self.addFiles(files);
                }
            });

            document.getElementById('clearAllFiles').addEventListener('click', function() {
                self.clearAllFiles();
            });

            document.getElementById('enableDedup').addEventListener('change', function() {
                const dedupConfig = document.getElementById('dedupConfig');
                if (this.checked) {
                    dedupConfig.classList.add('show');
                } else {
                    dedupConfig.classList.remove('show');
                }
            });

            document.getElementById('dedupColumn').addEventListener('change', function() {
                self.updateDedupExplanation(this.value);
            });

            document.getElementById('validateBtn').addEventListener('click', function() {
                self.validateFiles();
            });

            document.getElementById('mergeBtn').addEventListener('click', function() {
                self.mergeFiles();
            });

            document.getElementById('previewSearchBtn').addEventListener('click', function() {
                self.searchPreview();
            });

            document.getElementById('previewClearBtn').addEventListener('click', function() {
                self.clearPreviewSearch();
            });

            document.getElementById('previewSearch').addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    self.searchPreview();
                }
            });

            document.getElementById('saveMergedBtn').addEventListener('click', function() {
                self.saveMergedFile();
            });

            document.getElementById('closeModal').addEventListener('click', function() {
                self.hideModal();
            });

            document.getElementById('modalOkBtn').addEventListener('click', function() {
                self.hideModal();
            });

            document.getElementById('errorModal').addEventListener('click', function(e) {
                if (e.target === this) {
                    self.hideModal();
                }
            });

            document.getElementById('toggleLog').addEventListener('click', function() {
                self.toggleLogPanel();
            });
        },

        switchTab: function(tabId) {
            document.querySelectorAll('.tab-btn').forEach(function(btn) {
                btn.classList.remove('active');
            });
            document.querySelectorAll('.tab-content').forEach(function(content) {
                content.classList.remove('active');
            });

            document.querySelector('[data-tab="' + tabId + '"]').classList.add('active');
            document.getElementById(tabId + '-tab').classList.add('active');
        },

        addFiles: function(files) {
            const self = this;
            let hasNewFiles = false;

            files.forEach(function(file) {
                const exists = self.files.some(function(f) {
                    return f.name === file.name && f.size === file.size;
                });

                if (!exists) {
                    self.files.push(file);
                    self.log('info', '添加文件: ' + file.name);
                    hasNewFiles = true;
                }
            });

            if (hasNewFiles) {
                this.validationPassed = false;
                this.mergeCompleted = false;
                this.resetMergeResult();
                this.updateMergeButton();
            }

            this.renderFileList();
        },

        removeFile: function(index) {
            const file = this.files[index];
            this.files.splice(index, 1);
            this.fileData.splice(index, 1);
            this.mergeCompleted = false;
            this.log('info', '移除文件: ' + file.name);
            this.renderFileList();
            this.updateMergeButton();
            this.resetValidation();
            this.resetMergeResult();
        },

        clearAllFiles: function() {
            this.files = [];
            this.fileData = [];
            this.mergeCompleted = false;
            this.log('info', '清空所有文件');
            this.renderFileList();
            this.updateMergeButton();
            this.resetValidation();
            this.resetMergeResult();
        },

        renderFileList: function() {
            const fileList = document.getElementById('mergeFileList');
            const fileListContent = document.getElementById('fileListContent');
            const fileCount = document.getElementById('fileCount');

            if (this.files.length === 0) {
                fileList.style.display = 'none';
                return;
            }

            fileList.style.display = 'block';
            fileCount.textContent = this.files.length;

            const self = this;
            fileListContent.innerHTML = this.files.map(function(file, index) {
                const status = self.getFileStatus(index);
                return '<div class="file-item ' + status.class + '" data-index="' + index + '">' +
                    '<div class="file-item-info">' +
                    '<div class="file-item-icon">' + status.icon + '</div>' +
                    '<div class="file-item-details">' +
                    '<div class="file-item-name">' + self.escapeHtml(file.name) + '</div>' +
                    '<div class="file-item-meta">' +
                    '<span>大小: ' + self.formatFileSize(file.size) + '</span>' +
                    '<span>修改: ' + self.formatDate(file.lastModified) + '</span>' +
                    '</div>' +
                    '</div>' +
                    '</div>' +
                    '<span class="file-item-status ' + status.statusClass + '">' + status.text + '</span>' +
                    '<button class="file-item-remove" onclick="window.MergeModule.removeFile(' + index + ')">移除</button>' +
                    '</div>';
            }).join('');
        },

        getFileStatus: function(index) {
            if (!this.fileData[index]) {
                return {
                    class: '',
                    icon: '📄',
                    text: '待验证',
                    statusClass: 'processing'
                };
            }

            const data = this.fileData[index];
            if (data.error) {
                return {
                    class: 'error',
                    icon: '❌',
                    text: '错误',
                    statusClass: 'invalid'
                };
            }

            if (data.validated) {
                return {
                    class: 'success',
                    icon: '✅',
                    text: '验证通过',
                    statusClass: 'valid'
                };
            }

            return {
                class: '',
                icon: '📄',
                text: '待验证',
                statusClass: 'processing'
            };
        },

        updateMergeButton: function() {
            const mergeBtn = document.getElementById('mergeBtn');
            const shouldDisable = this.files.length < 2 || !this.validationPassed || this.mergeCompleted;
            mergeBtn.disabled = shouldDisable;
            
            if (this.mergeCompleted) {
                mergeBtn.innerHTML = '<span class="btn-icon">✅</span>合并完成';
            } else if (this.isProcessing) {
                mergeBtn.innerHTML = '<span class="btn-icon">⏳</span>合并中...';
            } else {
                mergeBtn.innerHTML = '<span class="btn-icon">🔗</span>开始合并';
            }
        },

        validateFiles: function() {
            const self = this;

            if (this.files.length < 2) {
                this.showError('文件数量不足', '请至少选择2个Excel文件进行合并');
                return;
            }

            this.showProgress();
            this.updateProgress(0, '开始验证文件...');
            this.log('info', '开始验证 ' + this.files.length + ' 个文件');

            this.fileData = [];
            let processed = 0;
            const total = this.files.length;

            const validateNext = function(index) {
                if (index >= total) {
                    self.completeValidation();
                    return;
                }

                const file = self.files[index];
                self.updateProgress(
                    Math.round((index / total) * 100),
                    '正在验证: ' + file.name
                );

                self.readExcelFile(file, function(err, data) {
                    if (err) {
                        self.fileData[index] = { error: err.message };
                        self.log('error', '验证失败: ' + file.name + ' - ' + err.message);
                    } else {
                        self.fileData[index] = data;
                        self.log('success', '验证成功: ' + file.name + ' (' + data.rows.length + ' 行)');
                    }

                    processed++;
                    validateNext(index + 1);
                });
            };

            validateNext(0);
        },

        readExcelFile: function(file, callback) {
            const self = this;
            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    if (workbook.SheetNames.length === 0) {
                        callback(new Error('文件中没有工作表'));
                        return;
                    }

                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

                    if (jsonData.length === 0) {
                        callback(new Error('工作表为空'));
                        return;
                    }

                    const headers = jsonData[0].map(function(h) {
                        return String(h || '').trim();
                    });

                    const rows = jsonData.slice(1).filter(function(row) {
                        return row.some(function(cell) {
                            return cell !== '' && cell !== null && cell !== undefined;
                        });
                    });

                    callback(null, {
                        headers: headers,
                        rows: rows,
                        workbook: workbook,
                        sheet: firstSheet,
                        fileName: file.name
                    });
                } catch (err) {
                    callback(new Error('文件解析失败: ' + err.message));
                }
            };

            reader.onerror = function() {
                callback(new Error('文件读取失败'));
            };

            reader.readAsArrayBuffer(file);
        },

        completeValidation: function() {
            const self = this;
            this.updateProgress(100, '验证完成');

            const validFiles = this.fileData.filter(function(d) {
                return d && !d.error;
            });

            if (validFiles.length < 2) {
                this.hideProgress();
                this.validationPassed = false;
                this.showValidationResult(false, '有效文件不足', '至少需要2个有效文件才能进行合并');
                this.updateMergeButton();
                this.renderFileList();
                return;
            }

            const structureResult = this.validateStructure(validFiles);

            if (!structureResult.valid) {
                this.hideProgress();
                this.validationPassed = false;
                this.showValidationResult(false, '表结构不一致', structureResult.message);
                this.showStructureComparison(structureResult.comparison);
                this.updateMergeButton();
                this.renderFileList();
                return;
            }

            this.fileData.forEach(function(data) {
                if (data && !data.error) {
                    data.validated = true;
                }
            });

            this.headers = validFiles[0].headers;
            this.updateDedupColumns();

            this.hideProgress();
            this.validationPassed = true;
            this.showValidationResult(true, '验证通过', '所有文件表结构一致，可以进行合并');
            this.showStructureComparison(structureResult.comparison);
            this.updateMergeButton();
            this.renderFileList();

            this.log('success', '验证完成，可以开始合并');
        },

        validateStructure: function(fileDataList) {
            const self = this;
            const baseHeaders = fileDataList[0].headers;
            const comparison = {
                files: fileDataList.map(function(d) {
                    return d.fileName;
                }),
                headers: baseHeaders,
                results: []
            };

            for (let i = 1; i < fileDataList.length; i++) {
                const currentHeaders = fileDataList[i].headers;
                const result = {
                    fileName: fileDataList[i].fileName,
                    matches: [],
                    mismatches: []
                };

                if (currentHeaders.length !== baseHeaders.length) {
                    return {
                        valid: false,
                        message: '文件 "' + fileDataList[i].fileName + '" 的列数不一致（期望 ' + baseHeaders.length + ' 列，实际 ' + currentHeaders.length + ' 列）',
                        comparison: comparison
                    };
                }

                for (let j = 0; j < baseHeaders.length; j++) {
                    if (currentHeaders[j] === baseHeaders[j]) {
                        result.matches.push(j);
                    } else {
                        result.mismatches.push({
                            index: j,
                            expected: baseHeaders[j],
                            actual: currentHeaders[j]
                        });
                    }
                }

                comparison.results.push(result);

                if (result.mismatches.length > 0) {
                    const mismatchDetails = result.mismatches.map(function(m) {
                        return '第' + (m.index + 1) + '列: 期望"' + m.expected + '"，实际"' + m.actual + '"';
                    }).join('；');

                    return {
                        valid: false,
                        message: '文件 "' + fileDataList[i].fileName + '" 的列名不一致：' + mismatchDetails,
                        comparison: comparison
                    };
                }
            }

            return {
                valid: true,
                comparison: comparison
            };
        },

        showValidationResult: function(success, title, message) {
            const validationResult = document.getElementById('validationResult');
            const validationContent = document.getElementById('validationContent');

            validationResult.style.display = 'block';

            const resultClass = success ? 'success' : 'error';
            const icon = success ? '✅' : '❌';

            validationContent.innerHTML = '<div class="validation-result ' + resultClass + '">' +
                '<div class="validation-title">' +
                '<span class="validation-icon">' + icon + '</span>' +
                title +
                '</div>' +
                '<div class="validation-details">' + message + '</div>' +
                '</div>';
        },

        showStructureComparison: function(comparison) {
            if (!comparison) return;

            const validationContent = document.getElementById('validationContent');
            let html = '<div class="structure-comparison">' +
                '<h4 style="margin-bottom: 15px; color: var(--text-dark);">表结构对比</h4>' +
                '<table class="structure-table">' +
                '<thead><tr><th>列号</th><th>列名</th>';

            comparison.files.forEach(function(fileName) {
                html += '<th>' + this.escapeHtml(fileName.substring(0, 15)) + '</th>';
            }, this);

            html += '</tr></thead><tbody>';

            const self = this;
            comparison.headers.forEach(function(header, colIndex) {
                html += '<tr>';
                html += '<td>' + (colIndex + 1) + '</td>';
                html += '<td>' + self.escapeHtml(header) + '</td>';

                comparison.files.forEach(function(fileName, fileIndex) {
                    if (fileIndex === 0) {
                        html += '<td class="match">✅</td>';
                        return;
                    }

                    const result = comparison.results[fileIndex - 1];
                    if (!result) {
                        html += '<td class="match">✅</td>';
                        return;
                    }

                    const isMismatch = result.mismatches.some(function(m) {
                        return m.index === colIndex;
                    });

                    if (isMismatch) {
                        const mismatch = result.mismatches.find(function(m) {
                            return m.index === colIndex;
                        });
                        html += '<td class="mismatch" title="期望: ' + self.escapeHtml(mismatch.expected) + ', 实际: ' + self.escapeHtml(mismatch.actual) + '">❌</td>';
                    } else {
                        html += '<td class="match">✅</td>';
                    }
                });

                html += '</tr>';
            });

            html += '</tbody></table></div>';
            validationContent.innerHTML += html;
        },

        updateDedupColumns: function() {
            const select = document.getElementById('dedupColumn');
            select.innerHTML = '<option value="all">所有列完全相同</option>';

            this.headers.forEach(function(header, index) {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = header;
                select.appendChild(option);
            });

            this.updateDedupExplanation('all');
        },

        updateDedupExplanation: function(selectedValue) {
            const explanation = document.getElementById('dedupModeExplanation');
            if (!explanation) return;

            const titleEl = explanation.querySelector('.dedup-explanation-title');
            const descEl = explanation.querySelector('.dedup-explanation-desc');

            if (selectedValue === 'all') {
                explanation.classList.remove('example-mode');
                titleEl.textContent = '当前模式：全列匹配去重';
                descEl.innerHTML = '只有当一行数据的<strong>所有列</strong>内容都完全相同时，才会被视为重复行并移除。';
            } else {
                const columnName = this.headers[parseInt(selectedValue)] || '未知列';
                explanation.classList.add('example-mode');
                titleEl.textContent = '当前模式：单列匹配去重';
                descEl.innerHTML = '只根据<strong>' + this.escapeHtml(columnName) + '</strong>列的内容判断重复。只要该列值相同，无论其他列是否相同，都会被视为重复行。';
            }
        },

        mergeFiles: function() {
            const self = this;

            if (this.isProcessing) {
                return;
            }

            this.isProcessing = true;
            this.startTime = Date.now();
            this.updateMergeButton();

            const validFiles = this.fileData.filter(function(d) {
                return d && !d.error;
            });

            if (validFiles.length < 2) {
                this.showError('文件数量不足', '请先验证文件');
                this.isProcessing = false;
                this.updateMergeButton();
                return;
            }

            this.showProgress();
            this.updateProgress(0, '正在合并Excel文件，请稍候...');
            this.log('info', '开始合并 ' + validFiles.length + ' 个文件');

            const enableDedup = document.getElementById('enableDedup').checked;
            const skipEmptyRows = document.getElementById('skipEmptyRows').checked;
            const dedupColumn = document.getElementById('dedupColumn').value;

            const allRows = [];
            let totalRows = 0;

            validFiles.forEach(function(fileData) {
                totalRows += fileData.rows.length;
            });

            let processedRows = 0;

            validFiles.forEach(function(fileData, fileIndex) {
                self.updateProgress(
                    Math.round((processedRows / totalRows) * 80),
                    '正在处理: ' + fileData.fileName
                );

                fileData.rows.forEach(function(row) {
                    if (skipEmptyRows) {
                        const isEmpty = row.every(function(cell) {
                            return cell === '' || cell === null || cell === undefined;
                        });
                        if (isEmpty) return;
                    }

                    allRows.push(row.slice());
                    processedRows++;
                });
            });

            this.updateProgress(80, '正在合并数据...');
            this.log('info', '合并完成，共 ' + allRows.length + ' 行');

            let dedupCount = 0;
            if (enableDedup) {
                this.updateProgress(85, '正在去重...');
                const beforeDedup = allRows.length;
                this.deduplicateRows(allRows, dedupColumn);
                dedupCount = beforeDedup - allRows.length;
                this.log('info', '去重完成，移除 ' + dedupCount + ' 行重复数据');
            }

            this.updateProgress(95, '正在生成结果...');

            this.mergedData = {
                headers: this.headers.slice(),
                rows: allRows
            };

            const processTime = ((Date.now() - this.startTime) / 1000).toFixed(2);

            this.updateProgress(100, '合并完成！');
            this.hideProgress(true);
            this.showMergeResult(allRows.length, validFiles.length, processTime, dedupCount);
            this.showPreview();
            this.showSaveSection();

            this.isProcessing = false;
            this.mergeCompleted = true;
            this.updateMergeButton();
            
            this.log('success', '合并完成！总行数: ' + allRows.length);
            
            this.showMergeSuccessNotification(allRows.length, validFiles.length);
            
            setTimeout(function() {
                self.scrollToResults();
            }, 500);
        },

        deduplicateRows: function(rows, dedupColumn) {
            const seen = new Set();

            rows.sort(function(a, b) {
                return 0;
            });

            let writeIndex = 0;
            for (let i = 0; i < rows.length; i++) {
                let key;
                if (dedupColumn === 'all') {
                    key = rows[i].join('\x00');
                } else {
                    key = String(rows[i][parseInt(dedupColumn)] || '');
                }

                if (!seen.has(key)) {
                    seen.add(key);
                    if (writeIndex !== i) {
                        rows[writeIndex] = rows[i];
                    }
                    writeIndex++;
                }
            }

            rows.length = writeIndex;
        },

        showMergeResult: function(totalRows, totalFiles, processTime, dedupRows) {
            const resultSection = document.getElementById('mergeResultSection');
            resultSection.style.display = 'block';

            document.getElementById('totalRows').textContent = totalRows.toLocaleString();
            document.getElementById('totalFiles').textContent = totalFiles;
            document.getElementById('processTime').textContent = processTime + 's';

            const dedupStat = document.getElementById('dedupStat');
            if (dedupRows > 0) {
                dedupStat.style.display = 'flex';
                document.getElementById('dedupRows').textContent = dedupRows.toLocaleString();
            } else {
                dedupStat.style.display = 'none';
            }
        },

        showPreview: function() {
            if (!this.mergedData) return;

            const previewSection = document.getElementById('previewSection');
            previewSection.style.display = 'block';

            this.filteredData = this.mergedData.rows.slice();
            this.currentPage = 1;
            this.renderPreviewTable();
        },

        renderPreviewTable: function() {
            const table = document.getElementById('previewTable');
            const previewCount = document.getElementById('previewCount');

            const data = this.filteredData || this.mergedData.rows;
            const totalRows = data.length;
            const totalPages = Math.ceil(totalRows / this.pageSize);

            previewCount.textContent = Math.min(this.pageSize, totalRows - (this.currentPage - 1) * this.pageSize);

            let html = '<thead><tr>';
            this.mergedData.headers.forEach(function(header, index) {
                const sortClass = this.sortColumn === index ?
                    (this.sortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc') : '';
                html += '<th class="' + sortClass + '" onclick="window.MergeModule.sortTable(' + index + ')">' +
                    this.escapeHtml(header) + '</th>';
            }, this);
            html += '</tr></thead><tbody>';

            const start = (this.currentPage - 1) * this.pageSize;
            const end = Math.min(start + this.pageSize, totalRows);

            for (let i = start; i < end; i++) {
                html += '<tr>';
                this.mergedData.headers.forEach(function(header, index) {
                    const cellValue = data[i][index];
                    html += '<td>' + this.escapeHtml(String(cellValue !== undefined ? cellValue : '')) + '</td>';
                }, this);
                html += '</tr>';
            }

            html += '</tbody>';
            table.innerHTML = html;

            this.renderPagination(totalPages);
        },

        renderPagination: function(totalPages) {
            const pagination = document.getElementById('pagination');

            if (totalPages <= 1) {
                pagination.innerHTML = '';
                return;
            }

            let html = '<button class="page-btn" onclick="window.MergeModule.goToPage(1)" ' +
                (this.currentPage === 1 ? 'disabled' : '') + '>首页</button>';
            html += '<button class="page-btn" onclick="window.MergeModule.goToPage(' + (this.currentPage - 1) + ')" ' +
                (this.currentPage === 1 ? 'disabled' : '') + '>上一页</button>';

            const startPage = Math.max(1, this.currentPage - 2);
            const endPage = Math.min(totalPages, this.currentPage + 2);

            for (let i = startPage; i <= endPage; i++) {
                html += '<button class="page-btn ' + (i === this.currentPage ? 'active' : '') + '" ' +
                    'onclick="window.MergeModule.goToPage(' + i + ')">' + i + '</button>';
            }

            html += '<button class="page-btn" onclick="window.MergeModule.goToPage(' + (this.currentPage + 1) + ')" ' +
                (this.currentPage === totalPages ? 'disabled' : '') + '>下一页</button>';
            html += '<button class="page-btn" onclick="window.MergeModule.goToPage(' + totalPages + ')" ' +
                (this.currentPage === totalPages ? 'disabled' : '') + '>末页</button>';

            pagination.innerHTML = html;
        },

        goToPage: function(page) {
            const data = this.filteredData || this.mergedData.rows;
            const totalPages = Math.ceil(data.length / this.pageSize);

            if (page < 1) page = 1;
            if (page > totalPages) page = totalPages;

            this.currentPage = page;
            this.renderPreviewTable();
        },

        sortTable: function(columnIndex) {
            if (!this.filteredData) return;

            if (this.sortColumn === columnIndex) {
                this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                this.sortColumn = columnIndex;
                this.sortDirection = 'asc';
            }

            const self = this;
            this.filteredData.sort(function(a, b) {
                let valA = a[columnIndex];
                let valB = b[columnIndex];

                if (valA === undefined || valA === null) valA = '';
                if (valB === undefined || valB === null) valB = '';

                const numA = parseFloat(valA);
                const numB = parseFloat(valB);

                let comparison = 0;
                if (!isNaN(numA) && !isNaN(numB)) {
                    comparison = numA - numB;
                } else {
                    comparison = String(valA).localeCompare(String(valB), 'zh-CN');
                }

                return self.sortDirection === 'asc' ? comparison : -comparison;
            });

            this.currentPage = 1;
            this.renderPreviewTable();
        },

        searchPreview: function() {
            const searchTerm = document.getElementById('previewSearch').value.trim().toLowerCase();

            if (!searchTerm) {
                this.filteredData = this.mergedData.rows.slice();
            } else {
                this.filteredData = this.mergedData.rows.filter(function(row) {
                    return row.some(function(cell) {
                        return String(cell || '').toLowerCase().includes(searchTerm);
                    });
                });
            }

            this.currentPage = 1;
            this.renderPreviewTable();
            this.log('info', '搜索: "' + searchTerm + '"，找到 ' + this.filteredData.length + ' 条记录');
        },

        clearPreviewSearch: function() {
            document.getElementById('previewSearch').value = '';
            this.filteredData = this.mergedData.rows.slice();
            this.currentPage = 1;
            this.renderPreviewTable();
        },

        showSaveSection: function() {
            document.getElementById('saveSection').style.display = 'block';
        },

        saveMergedFile: function() {
            const self = this;
            
            if (!this.mergedData) {
                this.showError('没有数据', '请先合并文件');
                return;
            }

            const defaultFileName = 'merged_excel';
            const format = 'xlsx';
            const fullFileName = defaultFileName + '.' + format;

            const ws = XLSX.utils.aoa_to_sheet([this.mergedData.headers].concat(this.mergedData.rows));
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Merged Data');

            const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

            if (window.showSaveFilePicker) {
                const options = {
                    suggestedName: fullFileName,
                    types: [{
                        description: 'Excel 文件',
                        accept: {
                            [mimeType]: ['.xlsx']
                        }
                    }],
                    excludeAcceptAllOption: false
                };

                let saveHandle = null;
                
                window.showSaveFilePicker(options)
                    .then(function(handle) {
                        saveHandle = handle;
                        return handle.createWritable();
                    })
                    .then(function(writable) {
                        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                        const blob = new Blob([wbout], { type: mimeType });
                        
                        return writable.write(blob).then(function() {
                            writable.close();
                            self.log('success', '文件已保存: ' + saveHandle.name);
                            self.showSuccess('保存成功', '文件 "' + saveHandle.name + '" 已保存成功！');
                        });
                    })
                    .catch(function(err) {
                        if (err.name === 'AbortError') {
                            self.log('info', '用户取消了保存操作');
                            return;
                        }
                        self.showError('保存失败', '保存文件时发生错误: ' + err.message);
                    });
            } else {
                try {
                    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                    const blob = new Blob([wbout], { type: mimeType });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = fullFileName;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    
                    self.log('success', '文件已下载: ' + fullFileName);
                    self.showSuccess('下载成功', '文件 "' + fullFileName + '" 已开始下载！\n\n提示：您的浏览器不支持系统保存对话框，已使用下载方式保存。');
                } catch (err) {
                    self.showError('保存失败', '保存文件时发生错误: ' + err.message);
                }
            }
        },

        resetValidation: function() {
            document.getElementById('validationResult').style.display = 'none';
            document.getElementById('validationContent').innerHTML = '';
            this.validationPassed = false;
            this.updateMergeButton();
        },

        resetMergeResult: function() {
            document.getElementById('mergeResultSection').style.display = 'none';
            document.getElementById('previewSection').style.display = 'none';
            document.getElementById('saveSection').style.display = 'none';
            this.mergedData = null;
            this.filteredData = null;
        },

        showProgress: function() {
            document.getElementById('progressSection').style.display = 'block';
            document.getElementById('progressFill').style.width = '0%';
            document.getElementById('progressPercent').textContent = '0%';
            document.getElementById('progressText').textContent = '准备中...';
            document.getElementById('progressDetails').innerHTML = '';
        },

        updateProgress: function(percent, text) {
            document.getElementById('progressFill').style.width = percent + '%';
            document.getElementById('progressPercent').textContent = percent + '%';
            document.getElementById('progressText').textContent = text;

            const details = document.getElementById('progressDetails');
            const p = document.createElement('p');
            p.textContent = new Date().toLocaleTimeString() + ' - ' + text;
            details.appendChild(p);
            details.scrollTop = details.scrollHeight;
        },

        hideProgress: function(success) {
            const section = document.getElementById('progressSection');
            if (success) {
                section.classList.add('success');
                setTimeout(function() {
                    section.style.display = 'none';
                    section.classList.remove('success');
                }, 1500);
            } else {
                setTimeout(function() {
                    section.style.display = 'none';
                }, 500);
            }
        },

        showMergeSuccessNotification: function(totalRows, fileCount) {
            const existingNotification = document.querySelector('.merge-success-notification');
            if (existingNotification) {
                existingNotification.remove();
            }

            const notification = document.createElement('div');
            notification.className = 'merge-success-notification';
            notification.innerHTML = 
                '<div class="notification-content">' +
                    '<div class="notification-icon">✅</div>' +
                    '<div class="notification-text">' +
                        '<div class="notification-title">Excel文件合并成功！</div>' +
                        '<div class="notification-desc">已合并 ' + fileCount + ' 个文件，共 ' + totalRows.toLocaleString() + ' 行数据</div>' +
                        '<div class="notification-hint">已自动为您跳转到结果展示区域</div>' +
                    '</div>' +
                    '<button class="notification-close" onclick="this.parentElement.parentElement.remove()">×</button>' +
                '</div>';

            document.body.appendChild(notification);

            setTimeout(function() {
                notification.classList.add('show');
            }, 10);

            setTimeout(function() {
                notification.classList.remove('show');
                setTimeout(function() {
                    if (notification.parentElement) {
                        notification.remove();
                    }
                }, 300);
            }, 5000);
        },

        scrollToResults: function() {
            const resultSection = document.getElementById('mergeResultSection');
            if (resultSection && resultSection.style.display !== 'none') {
                resultSection.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });

                resultSection.classList.add('highlight-section');
                setTimeout(function() {
                    resultSection.classList.remove('highlight-section');
                }, 2000);
            }
        },

        showError: function(title, message) {
            this.showModal('⚠️', title, message, 'error');
            this.log('error', title + ': ' + message);
        },

        showSuccess: function(title, message) {
            this.showModal('✅', title, message, 'success');
        },

        showModal: function(icon, title, message, type) {
            const modal = document.getElementById('errorModal');
            const modalIcon = document.querySelector('.modal-icon');
            const modalTitle = document.getElementById('errorTitle');
            const modalBody = document.getElementById('errorMessage');

            modalIcon.textContent = icon;
            modalTitle.textContent = title;
            modalBody.innerHTML = '<p>' + this.escapeHtml(message) + '</p>';

            modal.style.display = 'flex';
        },

        hideModal: function() {
            document.getElementById('errorModal').style.display = 'none';
        },

        initLogPanel: function() {
            const logPanel = document.getElementById('logPanel');
            logPanel.style.display = 'none';
        },

        log: function(type, message) {
            const now = new Date();
            const timeStr = now.toLocaleTimeString();

            this.logs.push({ type: type, message: message, time: timeStr });

            const logPanel = document.getElementById('logPanel');
            const logContent = document.getElementById('logContent');

            logPanel.style.display = 'block';

            const entry = document.createElement('div');
            entry.className = 'log-entry ' + type;
            entry.innerHTML = '<span class="log-time">[' + timeStr + ']</span>' + this.escapeHtml(message);
            logContent.appendChild(entry);
            logContent.scrollTop = logContent.scrollHeight;

            if (this.logs.length > 100) {
                this.logs.shift();
                logContent.removeChild(logContent.firstChild);
            }
        },

        toggleLogPanel: function() {
            const logContent = document.getElementById('logContent');
            const toggleBtn = document.getElementById('toggleLog');

            if (logContent.style.display === 'none') {
                logContent.style.display = 'block';
                toggleBtn.textContent = '▼';
            } else {
                logContent.style.display = 'none';
                toggleBtn.textContent = '▲';
            }
        },

        formatFileSize: function(bytes) {
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
            return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
        },

        formatDate: function(timestamp) {
            const date = new Date(timestamp);
            return date.toLocaleDateString('zh-CN') + ' ' + date.toLocaleTimeString('zh-CN', {
                hour: '2-digit',
                minute: '2-digit'
            });
        },

        escapeHtml: function(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
    };

    window.MergeModule = MergeModule;

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            MergeModule.init();
        });
    } else {
        MergeModule.init();
    }
})();
