document.addEventListener('DOMContentLoaded', function() {
    const excelFile = document.getElementById('excelFile');
    const convertBtn = document.getElementById('convertBtn');
    const htmlResult = document.getElementById('htmlResult');
    const copyBtn = document.getElementById('copyBtn');
    const fileUploadWrapper = document.querySelector('.file-upload-wrapper');

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

        setTimeout(() => {
            fileInfo.style.opacity = '0';
            setTimeout(() => fileInfo.remove(), 300);
        }, 5000);
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
        const reader = new FileReader();

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            const originalHeaders = jsonData[0];
            const linkColumnIndexes = [];
            const imageColumnIndexes = [];
            const hideColumnIndexes = [];
            const hideColumnNames = ['拍卖平台', '标的类型', '省份', '城市'];
            
            originalHeaders.forEach((header, index) => {
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
            
            let html = XLSX.utils.sheet_to_html(worksheet, {
                header: '<tr style="background-color: #4CAF50; color: white;">',
                footer: '</tr>',
                tableClass: 'hybrid-table'
            });
            
            html = html.replace('<table>', '<table class="hybrid-table">');
            
            const parser = new DOMParser();
            const doc = parser.parseFromString(html, 'text/html');
            const table = doc.querySelector('table');
            
            if (table) {
                const rows = table.querySelectorAll('tr');
                
                // 先删除隐藏列
                if (hideColumnIndexes.length > 0) {
                    for (let i = 0; i < rows.length; i++) {
                        const cells = rows[i].querySelectorAll('td, th');
                        for (let j = hideColumnIndexes.length - 1; j >= 0; j--) {
                            const index = hideColumnIndexes[j];
                            if (cells[index]) {
                                cells[index].remove();
                            }
                        }
                    }
                }
                
                // 获取删除隐藏列后的表头（可能是 th 或 td）
                const headerRow = rows[0];
                const headerCells = headerRow.querySelectorAll('th, td');
                const headers = Array.from(headerCells).map(cell => cell.textContent.trim());
                
                // 处理链接和图片列（需要重新计算索引）
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    const cells = row.querySelectorAll('td');
                    
                    // 设置 data-label 属性
                    const minLen = Math.min(headers.length, cells.length);
                    for (let j = 0; j < minLen; j++) {
                        const headerText = headers[j];
                        if (headerText && headerText !== '') {
                            cells[j].setAttribute('data-label', headerText);
                        }
                        
                        // 处理链接列
                        if (headers[j] && headers[j].includes('链接')) {
                            const cell = cells[j];
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
                        
                        // 处理图片列
                        if (headers[j] && headers[j].includes('图片')) {
                            const cell = cells[j];
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
                
                html = table.outerHTML;
            }
            
            const styledHtml = `<!DOCTYPE html>
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
            /* 完全隐藏表头行 */
            .hybrid-table thead {
                display: none;
            }
            
            /* 隐藏第一行（表头行）在移动端的所有显示 */
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

            /* 标题列放在最前面 - 使用 CSS order 属性强制排到第一行 */
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

            /* 其他字段 - 改为上下布局，标签在上，内容在下 */
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

            /* 卡片模式下显示字段名 - 只对有 data-label 的单元格生效 */
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

            /* 标题字段不显示字段名 */
            .hybrid-table td[data-label="标题"]::before {
                display: none;
            }

            /* 标题链接特殊处理 */
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

            /* 图片链接特殊处理 */
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
            
            htmlResult.innerHTML = '<div class="table-wrapper">' + html + '</div>';
            htmlResult.classList.add('has-content');
            htmlResult.dataset.html = styledHtml;
            
            convertBtn.textContent = '✅ 转换成功！再次转换';
            setTimeout(() => {
                convertBtn.innerHTML = '<span class="btn-icon">✨</span>转换为 HTML';
            }, 2000);
        };

        reader.readAsArrayBuffer(file);
    });

    copyBtn.addEventListener('click', function() {
        const htmlContent = htmlResult.dataset.html;
        if (!htmlContent) {
            alert('请先转换 Excel 文件');
            return;
        }
        
        const blob = new Blob([htmlContent], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'excel-data.htm';
        document.body.appendChild(a);
        a.click();
        
        setTimeout(function() {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);
        
        const existingMessage = document.querySelector('.success-message');
        if (existingMessage) {
            existingMessage.remove();
        }
        
        const successMessage = document.createElement('div');
        successMessage.className = 'success-message';
        successMessage.textContent = '✅ HTML 文件已保存成功！';
        htmlResult.parentNode.appendChild(successMessage);
        
        setTimeout(function() {
            successMessage.style.display = 'none';
        }, 3000);
    });
});
