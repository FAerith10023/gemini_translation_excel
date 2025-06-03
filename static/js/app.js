// Excel 翻译器 Web 应用 JavaScript

class ExcelTranslatorApp {
    constructor() {
        this.apiKey = '';
        this.uploadedFile = null;
        this.currentFilename = '';
        this.downloadFilename = '';
        this.terminologyDownloadFilename = '';
        this.terminologyMatchedFilename = '';  // 用于翻译的术语库匹配后文件名
        
        this.initializeEventListeners();
        this.checkFormValidity();
    }

    initializeEventListeners() {
        // API 密钥切换显示/隐藏
        document.getElementById('toggleApiKey').addEventListener('click', () => {
            this.toggleApiKeyVisibility();
        });

        // API 连接测试
        document.getElementById('testConnection').addEventListener('click', () => {
            this.testApiConnection();
        });

        // 文件上传
        document.getElementById('fileInput').addEventListener('change', (e) => {
            this.handleFileUpload(e);
        });

        // 关键词快捷按钮
        document.querySelectorAll('.keyword-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.selectKeyword(e.target.dataset.keyword);
            });
        });

        // 术语库匹配
        document.getElementById('terminologyMatch').addEventListener('click', () => {
            this.startTerminologyMatch();
        });

        // 开始翻译
        document.getElementById('startTranslation').addEventListener('click', () => {
            this.startTranslation();
        });

        // 下载术语库匹配结果
        document.getElementById('downloadTerminologyResult').addEventListener('click', () => {
            this.downloadTerminologyResult();
        });

        // 术语库匹配后继续翻译
        document.getElementById('translateAfterTerminology').addEventListener('click', () => {
            this.translateAfterTerminology();
        });

        // 下载结果
        document.getElementById('downloadResult').addEventListener('click', () => {
            this.downloadResult();
        });

        // 新的翻译
        document.getElementById('newTranslation').addEventListener('click', () => {
            this.resetForm();
        });

        // 表单验证
        document.getElementById('apiKey').addEventListener('input', () => {
            this.checkFormValidity();
        });

        document.getElementById('fileInput').addEventListener('change', () => {
            this.checkFormValidity();
        });
    }

    toggleApiKeyVisibility() {
        const apiKeyInput = document.getElementById('apiKey');
        const toggleBtn = document.getElementById('toggleApiKey');
        const icon = toggleBtn.querySelector('i');

        if (apiKeyInput.type === 'password') {
            apiKeyInput.type = 'text';
            icon.className = 'bi bi-eye-slash';
        } else {
            apiKeyInput.type = 'password';
            icon.className = 'bi bi-eye';
        }
    }

    async testApiConnection() {
        const apiKey = document.getElementById('apiKey').value.trim();
        const testBtn = document.getElementById('testConnection');
        const resultDiv = document.getElementById('connectionResult');

        if (!apiKey) {
            this.showToast('错误', 'API 密钥不能为空', 'error');
            return;
        }

        // 显示加载状态
        const originalText = testBtn.innerHTML;
        testBtn.innerHTML = '<span class="loading-spinner"></span> 测试中...';
        testBtn.disabled = true;

        try {
            const response = await fetch('/api/test-connection', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ api_key: apiKey })
            });

            const data = await response.json();

            if (data.success) {
                resultDiv.innerHTML = `<div class="status-success">
                    <i class="bi bi-check-circle"></i> ${data.message}
                </div>`;
                this.showToast('成功', 'API 连接测试成功', 'success');
                this.apiKey = apiKey;
            } else {
                resultDiv.innerHTML = `<div class="status-error">
                    <i class="bi bi-x-circle"></i> ${data.message}
                </div>`;
                this.showToast('错误', data.message, 'error');
            }
        } catch (error) {
            resultDiv.innerHTML = `<div class="status-error">
                <i class="bi bi-x-circle"></i> 连接失败: ${error.message}
            </div>`;
            this.showToast('错误', `连接失败: ${error.message}`, 'error');
        } finally {
            testBtn.innerHTML = originalText;
            testBtn.disabled = false;
            this.checkFormValidity();
        }
    }

    async handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        // 验证文件类型
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        
        if (!allowedTypes.includes(file.type) && 
            !file.name.toLowerCase().endsWith('.xlsx') && 
            !file.name.toLowerCase().endsWith('.xls')) {
            this.showToast('错误', '请选择 .xlsx 或 .xls 格式的文件', 'error');
            event.target.value = '';
            return;
        }

        // 验证文件大小 (16MB)
        if (file.size > 16 * 1024 * 1024) {
            this.showToast('错误', '文件大小不能超过 16MB', 'error');
            event.target.value = '';
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (data.success) {
                this.currentFilename = data.filename;
                this.uploadedFile = file;
                this.showToast('成功', '文件上传成功', 'success');
                
                // 显示文件信息
                await this.displayFileInfo(data);
                
                this.checkFormValidity();
            } else {
                this.showToast('错误', data.message, 'error');
                event.target.value = '';
            }
        } catch (error) {
            this.showToast('错误', `上传失败: ${error.message}`, 'error');
            event.target.value = '';
        }
    }

    async displayFileInfo(uploadData) {
        const fileInfoDiv = document.getElementById('fileInfo');
        const fileDetailsDiv = document.getElementById('fileDetails');

        // 基本文件信息
        let infoHtml = `
            <p><strong>文件名:</strong> ${uploadData.original_name}</p>
            <p><strong>文件大小:</strong> ${this.formatFileSize(uploadData.size)}</p>
        `;

        try {
            // 获取详细文件分析
            const response = await fetch(`/api/file-info/${uploadData.filename}`);
            const data = await response.json();

            if (data.success && data.info) {
                if (data.info.error) {
                    infoHtml += `<p><strong>分析结果:</strong> ${data.info.error}</p>`;
                } else {
                    infoHtml += `
                        <p><strong>工作表数量:</strong> ${data.info.total_sheets}</p>
                        <p><strong>包含中文的单元格:</strong> ${data.info.chinese_cells}</p>
                    `;
                    
                    if (data.info.sheets && data.info.sheets.length > 0) {
                        infoHtml += `<p><strong>工作表名称:</strong> ${data.info.sheets.join(', ')}</p>`;
                    }
                }
            }
        } catch (error) {
            infoHtml += `<p><strong>分析结果:</strong> 无法分析文件内容</p>`;
        }

        fileDetailsDiv.innerHTML = infoHtml;
        fileInfoDiv.style.display = 'block';
    }

    selectKeyword(keyword) {
        const keywordsInput = document.getElementById('keywords');
        keywordsInput.value = keyword;

        // 更新按钮状态
        document.querySelectorAll('.keyword-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        
        event.target.classList.add('active');
        
        this.showToast('已选择', `专业领域: ${keyword}`, 'info');
    }

    async startTranslation() {
        const apiKey = document.getElementById('apiKey').value.trim();
        const keywords = document.getElementById('keywords').value.trim();

        if (!apiKey || !this.currentFilename) {
            this.showToast('错误', '请确保已输入 API 密钥并上传文件', 'error');
            return;
        }

        // 显示进度
        document.getElementById('translationProgress').style.display = 'block';
        document.getElementById('startTranslation').disabled = true;

        try {
            const response = await fetch('/api/translate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    api_key: apiKey,
                    filename: this.currentFilename,
                    keywords: keywords
                })
            });

            const data = await response.json();

            if (data.success) {
                this.downloadFilename = data.download_filename;
                this.showTranslationResult(data);
                this.showToast('成功', '翻译完成！', 'success');
            } else {
                this.showToast('错误', data.message, 'error');
            }
        } catch (error) {
            this.showToast('错误', `翻译失败: ${error.message}`, 'error');
        } finally {
            document.getElementById('translationProgress').style.display = 'none';
            document.getElementById('startTranslation').disabled = false;
        }
    }

    showTranslationResult(data) {
        const resultSection = document.getElementById('resultSection');
        const resultContent = document.getElementById('resultContent');

        const resultHtml = `
            <div class="alert alert-success">
                <h6><i class="bi bi-check-circle"></i> 翻译成功完成！</h6>
                <p class="mb-2"><strong>原文件:</strong> ${this.uploadedFile.name}</p>
                <p class="mb-2"><strong>译文件:</strong> ${data.download_filename}</p>
                <p class="mb-0"><strong>文件大小:</strong> ${this.formatFileSize(data.output_size)}</p>
            </div>
        `;

        resultContent.innerHTML = resultHtml;
        resultSection.style.display = 'block';

        // 滚动到结果区域
        resultSection.scrollIntoView({ behavior: 'smooth' });
    }

    downloadResult() {
        if (this.downloadFilename) {
            window.location.href = `/api/download/${this.downloadFilename}`;
        } else {
            this.showToast('错误', '没有可下载的文件', 'error');
        }
    }

    async startTerminologyMatch() {
        const apiKey = document.getElementById('apiKey').value.trim();

        if (!apiKey || !this.currentFilename) {
            this.showToast('错误', '请确保已输入 API 密钥并上传文件', 'error');
            return;
        }

        // 显示进度
        document.getElementById('terminologyProgress').style.display = 'block';
        document.getElementById('terminologyMatch').disabled = true;

        try {
            const response = await fetch('/api/terminology-match', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    api_key: apiKey,
                    filename: this.currentFilename
                })
            });

            const data = await response.json();

            if (data.success) {
                this.terminologyDownloadFilename = data.download_filename;
                this.terminologyMatchedFilename = data.matched_filename;  // 保存用于翻译的文件名
                this.showTerminologyResult(data);
                this.showToast('成功', `术语库匹配完成！共替换 ${data.replacement_count} 个术语`, 'success');
            } else {
                this.showToast('错误', data.message, 'error');
            }
        } catch (error) {
            this.showToast('错误', `术语库匹配失败: ${error.message}`, 'error');
        } finally {
            document.getElementById('terminologyProgress').style.display = 'none';
            document.getElementById('terminologyMatch').disabled = false;
        }
    }

    showTerminologyResult(data) {
        const resultSection = document.getElementById('terminologyResultSection');
        const resultContent = document.getElementById('terminologyResultContent');

        resultContent.innerHTML = `
            <div class="alert alert-success">
                <h6><i class="bi bi-check-circle"></i> 术语库匹配成功</h6>
                <p class="mb-1"><strong>匹配文件:</strong> ${data.download_filename}</p>
                <p class="mb-1"><strong>替换术语数:</strong> ${data.replacement_count} 个</p>
                <p class="mb-0"><strong>文件大小:</strong> ${this.formatFileSize(data.file_size)}</p>
            </div>
            <div class="alert alert-info">
                <p class="mb-0"><i class="bi bi-info-circle"></i> 
                   术语库匹配已完成，您可以下载匹配后的文件，或继续进行AI翻译。</p>
            </div>
        `;

        resultSection.style.display = 'block';
        resultSection.scrollIntoView({ behavior: 'smooth' });
    }

    async downloadTerminologyResult() {
        if (!this.terminologyDownloadFilename) {
            this.showToast('错误', '没有可下载的术语库匹配文件', 'error');
            return;
        }

        try {
            const response = await fetch(`/api/download/${this.terminologyDownloadFilename}`);
            
            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = this.terminologyDownloadFilename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                this.showToast('成功', '术语库匹配文件下载完成', 'success');
            } else {
                this.showToast('错误', '下载文件失败', 'error');
            }
        } catch (error) {
            this.showToast('错误', `下载失败: ${error.message}`, 'error');
        }
    }

    async translateAfterTerminology() {
        // 使用术语库匹配后的文件进行翻译
        if (!this.terminologyMatchedFilename) {
            this.showToast('错误', '没有术语库匹配文件可用于翻译', 'error');
            return;
        }

        const apiKey = document.getElementById('apiKey').value.trim();
        const keywords = document.getElementById('keywords').value.trim();

        // 显示进度
        document.getElementById('translationProgress').style.display = 'block';
        document.getElementById('translateAfterTerminology').disabled = true;

        try {
            const response = await fetch('/api/translate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    api_key: apiKey,
                    filename: this.terminologyMatchedFilename,  // 使用术语库匹配后的文件
                    keywords: keywords
                })
            });

            const data = await response.json();

            if (data.success) {
                this.downloadFilename = data.download_filename;
                this.showTranslationResult(data);
                this.showToast('成功', '翻译完成！', 'success');
            } else {
                this.showToast('错误', data.message, 'error');
            }
        } catch (error) {
            this.showToast('错误', `翻译失败: ${error.message}`, 'error');
        } finally {
            document.getElementById('translationProgress').style.display = 'none';
            document.getElementById('translateAfterTerminology').disabled = false;
        }
    }

    resetForm() {
        // 重置表单
        document.getElementById('apiKey').value = '';
        document.getElementById('fileInput').value = '';
        document.getElementById('keywords').value = '';

        // 重置状态
        this.apiKey = '';
        this.uploadedFile = null;
        this.currentFilename = '';
        this.downloadFilename = '';
        this.terminologyDownloadFilename = '';
        this.terminologyMatchedFilename = '';

        // 隐藏信息区域
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('terminologyResultSection').style.display = 'none';
        document.getElementById('resultSection').style.display = 'none';
        document.getElementById('connectionResult').innerHTML = '';

        // 重置按钮状态
        document.querySelectorAll('.keyword-btn').forEach(btn => {
            btn.classList.remove('active');
        });

        this.checkFormValidity();
        this.showToast('已重置', '表单已重置，可以开始新的翻译', 'info');
    }

    checkFormValidity() {
        const apiKey = document.getElementById('apiKey').value.trim();
        const hasFile = this.currentFilename !== '';
        
        const isValid = apiKey && hasFile;
        
        document.getElementById('startTranslation').disabled = !isValid;
        document.getElementById('terminologyMatch').disabled = !isValid;
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    showToast(title, message, type = 'info') {
        const toast = document.getElementById('toast');
        const toastTitle = document.getElementById('toastTitle');
        const toastMessage = document.getElementById('toastMessage');

        // 设置内容
        toastTitle.textContent = title;
        toastMessage.textContent = message;

        // 设置样式
        const toastHeader = toast.querySelector('.toast-header');
        toastHeader.className = 'toast-header';
        
        switch (type) {
            case 'success':
                toastHeader.style.backgroundColor = '#198754';
                break;
            case 'error':
                toastHeader.style.backgroundColor = '#dc3545';
                break;
            case 'warning':
                toastHeader.style.backgroundColor = '#ffc107';
                break;
            default:
                toastHeader.style.backgroundColor = '#0dcaf0';
        }

        // 显示 Toast
        const bsToast = new bootstrap.Toast(toast);
        bsToast.show();
    }
}

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    new ExcelTranslatorApp();
}); 