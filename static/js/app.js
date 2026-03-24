/**
 * SPSS ANOVA 统计分析平台 - 前端逻辑
 */

// 全局状态
const state = {
    currentFile: null,
    currentFileData: null,  // 存储文件的 ArrayBuffer 副本
    columns: [],
    previewData: [],
    allData: [],
    analysisResults: null
};

// DOM 元素
const elements = {
    uploadZone: document.getElementById('uploadZone'),
    fileInput: document.getElementById('fileInput'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    fileSize: document.getElementById('fileSize'),
    sampleColumn: document.getElementById('sampleColumn'),
    valueColumnsContainer: document.getElementById('valueColumnsContainer'),
    selectAllIndicators: document.getElementById('selectAllIndicators'),
    deselectAllIndicators: document.getElementById('deselectAllIndicators'),
    analyzeBtn: document.getElementById('analyzeBtn'),
    downloadBtn: document.getElementById('downloadBtn'),
    previewContainer: document.getElementById('previewContainer'),
    dataInfo: document.getElementById('dataInfo'),
    colCount: document.getElementById('colCount'),
    rowCount: document.getElementById('rowCount'),
    resultsSection: document.getElementById('resultsSection'),
    resultsContainer: document.getElementById('resultsContainer'),
    alertContainer: document.getElementById('alertContainer'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    mergeParallel: document.getElementById('mergeParallel'),
    mergeOptions: document.getElementById('mergeOptions'),
    mergeSuffixLength: document.getElementById('mergeSuffixLength')
};

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    console.log('App initialized');
});

function initializeEventListeners() {
    // 上传区域点击
    elements.uploadZone.addEventListener('click', (e) => {
        // 防止点击文件信息区域时触发两次
        if (e.target.closest('.file-info')) return;
        elements.fileInput.click();
    });

    // 文件选择
    elements.fileInput.addEventListener('change', handleFileSelect);

    // 拖拽上传
    elements.uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.uploadZone.classList.add('dragover');
    });

    elements.uploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.uploadZone.classList.remove('dragover');
    });

    elements.uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.uploadZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            elements.fileInput.files = files;
            handleFileSelect({ target: elements.fileInput });
        }
    });

    // 分析按钮
    elements.analyzeBtn.addEventListener('click', performAnalysis);

    // 下载按钮
    elements.downloadBtn.addEventListener('click', downloadReport);

    // 合并平行样选项
    if (elements.mergeParallel) {
        elements.mergeParallel.addEventListener('change', (e) => {
            if (elements.mergeOptions) {
                elements.mergeOptions.style.display = e.target.checked ? 'block' : 'none';
            }
        });
    }

    // 全选/全不选按钮
    if (elements.selectAllIndicators) {
        elements.selectAllIndicators.addEventListener('click', () => {
            const checkboxes = elements.valueColumnsContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => cb.checked = true);
        });
    }

    if (elements.deselectAllIndicators) {
        elements.deselectAllIndicators.addEventListener('click', () => {
            const checkboxes = elements.valueColumnsContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => cb.checked = false);
        });
    }
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) {
        console.log('No file selected');
        return;
    }

    console.log('File selected:', file.name, file.size);

    // 验证文件类型
    const allowedTypes = ['.xlsx', '.xls'];
    const fileExt = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedTypes.includes(fileExt)) {
        showAlert('请上传Excel文件 (.xlsx 或 .xls)', 'error');
        return;
    }

    // 存储文件信息并创建副本
    state.currentFile = file;

    // 读取文件内容为 ArrayBuffer 以便重复使用
    const reader = new FileReader();
    reader.onload = function(e) {
        state.currentFileData = e.target.result;
        console.log('File data cached, size:', state.currentFileData.byteLength);

        // 显示文件信息
        elements.fileName.textContent = file.name;
        elements.fileSize.textContent = formatFileSize(file.size);
        elements.fileInfo.classList.add('show');
        elements.uploadZone.classList.add('has-file');

        // 读取列信息 - 使用新的 File 对象
        const fileForUpload = new File([state.currentFileData], file.name, { type: file.type });
        readFileColumns(fileForUpload);
    };
    reader.onerror = function(e) {
        showAlert('读取文件失败', 'error');
        console.error('FileReader error:', e);
    };
    reader.readAsArrayBuffer(file);
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function readFileColumns(file) {
    console.log('Reading file columns...');
    showLoading(true);

    const formData = new FormData();
    formData.append('file', file);

    // 创建 AbortController 用于超时控制
    const controller = new AbortController();
    const timeoutId = setTimeout(() => {
        console.log('Request timeout, aborting...');
        controller.abort();
    }, 30000); // 30秒超时

    try {
        console.log('Sending request to /get_columns...');
        const response = await fetch('/get_columns', {
            method: 'POST',
            body: formData,
            signal: controller.signal
        });

        clearTimeout(timeoutId);
        console.log('Response received:', response.status);

        if (!response.ok) {
            let errorMessage = `服务器错误: ${response.status}`;
            try {
                const errorData = await response.json();
                if (errorData.error) errorMessage = errorData.error;
            } catch (e) {
                // 如果无法解析JSON，使用默认错误消息
            }
            throw new Error(errorMessage);
        }

        const result = await response.json();
        console.log('Result:', result);

        if (result.error) {
            throw new Error(result.error);
        }

        // 更新状态
        state.columns = result.columns;
        state.previewData = result.preview;
        state.allData = result.all_data || result.preview;

        // 更新列选择器（传入智能识别结果）
        updateColumnSelectors(result.columns, result.detected);

        // 更新数据预览
        updatePreview(result.preview, result.shape);

        // 启用按钮
        elements.sampleColumn.disabled = false;
        elements.analyzeBtn.disabled = false;
        if (elements.selectAllIndicators) elements.selectAllIndicators.disabled = false;
        if (elements.deselectAllIndicators) elements.deselectAllIndicators.disabled = false;

        showAlert(`成功读取文件：${result.shape[0]} 行 × ${result.shape[1]} 列`, 'success');

    } catch (error) {
        clearTimeout(timeoutId);
        console.error('Error reading file:', error);

        if (error.name === 'AbortError') {
            showAlert('请求超时，请检查服务器是否正常运行', 'error');
        } else if (error.name === 'TypeError' && error.message.includes('fetch')) {
            showAlert('无法连接到服务器，请确保服务器已启动 (python app.py)', 'error');
        } else {
            showAlert('读取文件失败: ' + error.message, 'error');
        }
    } finally {
        showLoading(false);
    }
}

function updateColumnSelectors(columns, detected) {
    if (!columns || columns.length === 0) {
        console.warn('No columns received');
        return;
    }

    // 使用智能识别的结果
    const detectedSample = detected?.sample_column || columns[0];
    const detectedIndicators = detected?.indicator_columns || [];

    console.log('智能识别结果:', {
        sample_column: detectedSample,
        indicator_columns: detectedIndicators
    });

    // 样品名称列 - 使用智能识别的样品列
    elements.sampleColumn.innerHTML = columns.map((col) =>
        `<option value="${escapeHtml(col)}" ${col === detectedSample ? 'selected' : ''}>${escapeHtml(col)} ${col === detectedSample ? '✓ 样品' : ''}</option>`
    ).join('');

    // 指标数据列 - 使用多选框
    // 只显示检测到的指标列作为选项
    if (detectedIndicators.length > 0) {
        elements.valueColumnsContainer.innerHTML = detectedIndicators.map((col, index) => {
            const colId = `indicator_${index}`;
            return `
                <div class="checkbox-item">
                    <input type="checkbox" id="${colId}" value="${escapeHtml(col)}" checked>
                    <label for="${colId}">
                        ${escapeHtml(col)}
                        <span class="auto-indicator">📈 自动识别</span>
                    </label>
                </div>
            `;
        }).join('');
    } else {
        elements.valueColumnsContainer.innerHTML = `
            <div style="color: var(--text-secondary); text-align: center; padding: 1rem;">
                <i class="fas fa-exclamation-triangle"></i> 未检测到数值型指标列
            </div>
        `;
    }

    // 显示自动识别提示
    if (detectedIndicators.length > 0) {
        const indicatorNames = detectedIndicators.slice(0, 3).join('、');
        const moreIndicator = detectedIndicators.length > 3 ? ` 等` : '';
        showAlert(`✓ 智能识别成功！样品列: "${detectedSample}"，指标列: ${indicatorNames}${moreIndicator} (${detectedIndicators.length}个)`, 'success');
    } else {
        showAlert('⚠ 未检测到数值型指标列，请检查数据格式', 'warning');
    }

    console.log('Column selectors updated with auto-detection');

    // 同步绘图列选择器
    const xSel = document.getElementById('plotXCol');
    if (xSel) {
        xSel.innerHTML = '<option value="">-- 选择列 --</option>' + getColumnOptions('');
        if (detected && detected.sample_column) xSel.value = detected.sample_column;
    }
    // 同步热图列选择器
    const hmRowSel = document.getElementById('heatmapRowCol');
    if (hmRowSel) {
        hmRowSel.innerHTML = '<option value="">-- 选择行标签列 --</option>' + getColumnOptions('');
    }
    const hmValSel = document.getElementById('heatmapValueCols');
    if (hmValSel && state.columns) {
        hmValSel.innerHTML = state.columns.map(c =>
            `<label style="display:flex;align-items:center;gap:0.4rem;padding:2px 0;cursor:pointer;">
                <input type="checkbox" value="${c}" style="width:15px;height:15px;cursor:pointer;">
                <span style="font-size:0.85rem;">${c}</span>
            </label>`
        ).join('');
    }
    // 同步 PCA 列选择器
    const pcaValSel = document.getElementById('pcaValueCols');
    if (pcaValSel && state.columns) {
        pcaValSel.innerHTML = state.columns.map(c =>
            `<label style="display:flex;align-items:center;gap:0.4rem;padding:2px 0;cursor:pointer;">
                <input type="checkbox" value="${c}" style="width:15px;height:15px;cursor:pointer;">
                <span style="font-size:0.85rem;">${c}</span>
            </label>`
        ).join('');
    }
    refreshPlotColumnSelectors();
}

// PCA 分组管理
const PCA_COLORS = ['#1f77b4','#ff7f0e','#2ca02c','#d62728','#9467bd','#8c564b','#e377c2','#7f7f7f','#bcbd22','#17becf'];

function pcaGetRowLabels() {
    const allData = state.allData && state.allData.length > 0 ? state.allData : state.previewData;
    if (!allData || allData.length === 0) return [];
    // 找第一个非数值列作为样本名列（通常是指标名列）
    const firstRow = allData[0];
    const cols = Object.keys(firstRow);
    const labelCol = cols.find(c => isNaN(parseFloat(firstRow[c]))) || cols[0];
    return allData.map((row, i) => ({ idx: i, label: String(row[labelCol] ?? i) }));
}

// 获取所有已被其他组选中的idx集合
function pcaGetUsedIndices(excludeRow) {
    const used = new Set();
    document.querySelectorAll('#pcaGroupRows .pca-group-row').forEach(row => {
        if (row === excludeRow) return;
        row.querySelectorAll('.pca-row-check:checked').forEach(cb => used.add(parseInt(cb.value)));
    });
    return used;
}

// 刷新所有组的checkbox禁用状态
function pcaRefreshDisabled() {
    document.querySelectorAll('#pcaGroupRows .pca-group-row').forEach(row => {
        const used = pcaGetUsedIndices(row);
        row.querySelectorAll('.pca-row-check').forEach(cb => {
            const idx = parseInt(cb.value);
            if (used.has(idx) && !cb.checked) {
                cb.disabled = true;
                cb.parentElement.style.opacity = '0.35';
                cb.parentElement.style.cursor = 'not-allowed';
            } else {
                cb.disabled = false;
                cb.parentElement.style.opacity = '';
                cb.parentElement.style.cursor = 'pointer';
            }
        });
    });
}

function pcaAddGroupRow(name='', color='', selectedIndices=[]) {
    const container = document.getElementById('pcaGroupRows');
    if (!container) return;
    const grpIdx = container.querySelectorAll('.pca-group-row').length;
    const c = color || PCA_COLORS[grpIdx % PCA_COLORS.length];
    const labels = pcaGetRowLabels();

    const row = document.createElement('div');
    row.className = 'pca-group-row';
    row.style.cssText = 'border:1px solid #f0c060;border-radius:6px;padding:0.5rem;margin-bottom:0.5rem;background:#fffbeb;';

    const selSet = new Set(selectedIndices.map(Number));
    const optionsHtml = labels.map(({idx, label}) =>
        `<label style="display:flex;align-items:center;gap:0.3rem;padding:1px 4px;cursor:pointer;font-size:0.82rem;white-space:nowrap;">
            <input type="checkbox" class="pca-row-check" value="${idx}" ${selSet.has(idx)?'checked':''} style="cursor:pointer;">
            <span>${escapeHtml(label)}</span>
        </label>`
    ).join('');

    row.innerHTML = `
        <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.4rem;">
            <input type="text" class="pca-grp-name form-select" placeholder="组名" value="${escapeHtml(name)}" style="width:120px;padding:0.3rem 0.5rem;font-size:0.82rem;">
            <input type="color" class="pca-grp-color" value="${c}" style="height:30px;width:36px;padding:1px;border:1px solid #ccc;border-radius:4px;cursor:pointer;">
            <span style="font-size:0.78rem;color:#92400e;">选择属于此组的数据行：</span>
            <button type="button" style="margin-left:auto;background:#ef4444;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:0.85rem;padding:2px 8px;" onclick="this.closest('.pca-group-row').remove();pcaRefreshDisabled();">删除组</button>
        </div>
        <div class="pca-row-list" style="display:flex;flex-wrap:wrap;gap:2px;max-height:120px;overflow-y:auto;border:1px solid #e5c97a;border-radius:4px;padding:4px;background:#fff;">${optionsHtml}</div>`;
    container.appendChild(row);
    // 监听勾选变化，刷新禁用状态
    row.querySelectorAll('.pca-row-check').forEach(cb => cb.addEventListener('change', pcaRefreshDisabled));
    pcaRefreshDisabled();
}

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('pcaAddGroupBtn')?.addEventListener('click', () => {
        if (!state.previewData || state.previewData.length === 0) {
            showAlert('请先上传数据文件', 'error'); return;
        }
        pcaAddGroupRow();
    });
});

function updatePreview(data, shape) {
    if (!data || data.length === 0) {
        console.warn('No preview data');
        return;
    }

    // 更新数据信息
    elements.dataInfo.style.display = 'flex';
    elements.colCount.textContent = `${shape[1]} 列`;
    elements.rowCount.textContent = `${shape[0]} 行`;

    // 创建预览表格
    let columns = Object.keys(data[0]);

    // 将样品名称列移到最左侧
    const sampleCol = elements.sampleColumn.value;
    if (sampleCol && columns.includes(sampleCol)) {
        columns = [sampleCol, ...columns.filter(c => c !== sampleCol)];
    }

    let html = `
        <table class="preview-table">
            <thead>
                <tr>
                    ${columns.map(col => `<th>${escapeHtml(col)}</th>`).join('')}
                </tr>
            </thead>
            <tbody>
                ${data.map(row => `
                    <tr>
                        ${columns.map(col => `<td>${escapeHtml(String(row[col] ?? ''))}</td>`).join('')}
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    if (shape[0] > 10) {
        html += `<div style="text-align: center; padding: 1rem; color: var(--text-secondary); font-size: 0.875rem;">
            <i class="fas fa-ellipsis-h"></i> 还有 ${shape[0] - 10} 行数据未显示
        </div>`;
    }

    elements.previewContainer.innerHTML = html;
    console.log('Preview updated');
}

async function performAnalysis() {
    if (!state.currentFile || !state.currentFileData) {
        showAlert('请先上传文件', 'error');
        return;
    }

    console.log('Starting analysis...');
    showLoading(true);

    // 使用存储的文件数据创建新的 File 对象
    const file = new File([state.currentFileData], state.currentFile.name, { type: state.currentFile.type });

    const formData = new FormData();
    formData.append('file', file);
    formData.append('sample_column', elements.sampleColumn.value);

    // 获取选中的指标列（多选框）
    const checkedBoxes = elements.valueColumnsContainer.querySelectorAll('input[type="checkbox"]:checked');
    const selectedIndicators = Array.from(checkedBoxes).map(cb => cb.value);
    if (selectedIndicators.length > 0) {
        formData.append('value_columns', selectedIndicators.join(','));
    }

    // 添加合并平行样参数
    const mergeParallel = elements.mergeParallel ? elements.mergeParallel.checked : false;
    const mergeSuffixLength = elements.mergeSuffixLength ? parseInt(elements.mergeSuffixLength.value) || 1 : 1;
    formData.append('merge_parallel', mergeParallel);
    formData.append('merge_suffix_length', mergeSuffixLength);
    console.log('Selected indicators:', selectedIndicators, 'Merge parallel:', mergeParallel, 'Suffix length:', mergeSuffixLength);

    // 创建 AbortController 用于超时控制
    const controller = new AbortController();
    const timeoutId = setTimeout(() => {
        console.log('Analysis timeout, aborting...');
        controller.abort();
    }, 60000); // 60秒超时

    try {
        console.log('Sending request to /upload...');
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData,
            signal: controller.signal
        });

        clearTimeout(timeoutId);
        console.log('Analysis response:', response.status);

        if (!response.ok) {
            let errorMessage = `服务器错误: ${response.status}`;
            try {
                const errorData = await response.json();
                if (errorData.error) errorMessage = errorData.error;
            } catch (e) {}
            throw new Error(errorMessage);
        }

        const result = await response.json();
        console.log('Analysis result:', result);

        if (result.error) {
            throw new Error(result.error);
        }

        state.analysisResults = result;

        // 显示结果
        displayResults(result);

        // 启用下载按钮
        elements.downloadBtn.disabled = false;

        showAlert('分析完成！', 'success');

    } catch (error) {
        clearTimeout(timeoutId);
        console.error('Analysis error:', error);

        if (error.name === 'AbortError') {
            showAlert('分析超时，请检查服务器是否正常运行或数据是否过大', 'error');
        } else if (error.name === 'TypeError' && error.message.includes('fetch')) {
            showAlert('无法连接到服务器，请确保服务器已启动', 'error');
        } else {
            showAlert('分析失败: ' + error.message, 'error');
        }
    } finally {
        showLoading(false);
    }
}

function displayResults(result) {
    if (!result || !result.indicators) {
        console.warn('No results to display');
        return;
    }

    elements.resultsSection.style.display = 'block';

    const indicators = result.indicators || [];
    const preview = result.preview || {};
    const summaryTable = result.summary_table || [];

    let html = '';

    // ── 多指标汇总表 ──
    if (indicators.length > 1 && summaryTable.length > 0) {
        html += `
        <div class="result-card">
            <div class="result-header">
                <div class="result-title"><i class="fas fa-table"></i> 多指标汇总（均值 ± 标准差，Duncan字母）</div>
            </div>
            <div style="overflow-x:auto; margin-top:1rem;">
                <table class="preview-table">
                    <thead><tr>
                        <th>样品名称</th>
                        ${indicators.map(ind => `<th colspan="3" style="text-align:center">${escapeHtml(String(ind))}</th>`).join('')}
                    </tr>
                    <tr>
                        <th></th>
                        ${indicators.map(() => `<th>均值</th><th>标准差</th><th>Duncan</th>`).join('')}
                    </tr></thead>
                    <tbody>
                        ${summaryTable.map(row => `
                        <tr>
                            <td><strong>${escapeHtml(row.sample)}</strong></td>
                            ${indicators.map(ind => {
                                const k = String(ind);
                                const mean = row[k+'_mean'] !== null ? row[k+'_mean'] : '-';
                                const std  = row[k+'_std']  !== null ? row[k+'_std']  : '-';
                                const dun  = row[k+'_duncan'] || '-';
                                return `<td>${mean}</td><td>${std}</td><td style="font-weight:bold;color:var(--primary-color)">${dun}</td>`;
                            }).join('')}
                        </tr>`).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
    }

    // ── 每个指标的 ANOVA / Levene 卡片 ──
    indicators.forEach((indicator, index) => {
        const data = preview[indicator];
        if (!data) return;

        const isSignificant = data.anova_significant;
        const isLeveneSignificant = data.levene_significant;

        // 单指标时也显示该指标的样品汇总
        let singleSummaryHtml = '';
        if (indicators.length === 1 && summaryTable.length > 0) {
            const k = String(indicator);
            singleSummaryHtml = `
            <div class="collapse-panel open">
                <div class="collapse-header" onclick="toggleCollapse(this)">
                    <div class="collapse-title"><i class="fas fa-list-ol"></i> 各样品统计结果</div>
                    <i class="fas fa-chevron-down collapse-icon"></i>
                </div>
                <div class="collapse-content"><div class="collapse-body">
                    <table class="preview-table">
                        <thead><tr><th>样品名称</th><th>均值</th><th>标准差</th><th>Duncan</th></tr></thead>
                        <tbody>
                            ${summaryTable.map(row => `
                            <tr>
                                <td><strong>${escapeHtml(row.sample)}</strong></td>
                                <td>${row[k+'_mean'] !== null ? row[k+'_mean'] : '-'}</td>
                                <td>${row[k+'_std']  !== null ? row[k+'_std']  : '-'}</td>
                                <td style="font-weight:bold;color:var(--primary-color)">${row[k+'_duncan'] || '-'}</td>
                            </tr>`).join('')}
                        </tbody>
                    </table>
                </div></div>
            </div>`;
        }

        html += `
            <div class="result-card">
                <div class="result-header">
                    <div class="result-title"><i class="fas fa-flask"></i> 指标: ${escapeHtml(String(indicator))}</div>
                    <span class="result-indicator">结果 ${index + 1}/${indicators.length}</span>
                </div>

                ${singleSummaryHtml}

                <div class="collapse-panel open">
                    <div class="collapse-header" onclick="toggleCollapse(this)">
                        <div class="collapse-title"><i class="fas fa-chart-pie"></i> 单因素方差分析 (ANOVA)</div>
                        <i class="fas fa-chevron-down collapse-icon"></i>
                    </div>
                    <div class="collapse-content"><div class="collapse-body">
                        <div class="stats-grid">
                            <div class="stat-box">
                                <div class="stat-label">F 统计量</div>
                                <div class="stat-value">${data.f_statistic !== undefined ? data.f_statistic : 'N/A'}</div>
                            </div>
                            <div class="stat-box">
                                <div class="stat-label">P 值</div>
                                <div class="stat-value ${isSignificant ? 'stat-significant' : 'stat-not-significant'}">
                                    ${data.p_value !== undefined ? data.p_value : 'N/A'}
                                </div>
                            </div>
                            <div class="stat-box">
                                <div class="stat-label">显著性</div>
                                <div class="stat-value ${isSignificant ? 'stat-significant' : 'stat-not-significant'}">
                                    ${isSignificant ? '显著 *' : '不显著'}
                                </div>
                            </div>
                        </div>
                        <div style="margin-top:1rem;font-size:0.875rem;color:var(--text-secondary);">
                            ${isSignificant
                                ? '<i class="fas fa-check-circle" style="color:var(--success-color);"></i> 不同组间存在显著差异 (p < 0.05)'
                                : '<i class="fas fa-times-circle" style="color:var(--error-color);"></i> 不同组间无显著差异 (p ≥ 0.05)'}
                        </div>
                    </div></div>
                </div>

                <div class="collapse-panel">
                    <div class="collapse-header" onclick="toggleCollapse(this)">
                        <div class="collapse-title"><i class="fas fa-balance-scale"></i> 方差齐性检验 (Levene's Test)</div>
                        <i class="fas fa-chevron-down collapse-icon"></i>
                    </div>
                    <div class="collapse-content"><div class="collapse-body">
                        <div class="stats-grid">
                            <div class="stat-box">
                                <div class="stat-label">Levene 统计量</div>
                                <div class="stat-value">${data.levene_significant !== undefined ? (isLeveneSignificant ? '显著' : '不显著') : 'N/A'}</div>
                            </div>
                            <div class="stat-box">
                                <div class="stat-label">P 值</div>
                                <div class="stat-value ${isLeveneSignificant ? 'stat-not-significant' : 'stat-significant'}">
                                    ${data.levene_significant !== undefined ? (isLeveneSignificant ? '< 0.05' : '≥ 0.05') : 'N/A'}
                                </div>
                            </div>
                            <div class="stat-box">
                                <div class="stat-label">结论</div>
                                <div class="stat-value ${isLeveneSignificant ? 'stat-not-significant' : 'stat-significant'}">
                                    ${isLeveneSignificant !== undefined ? (isLeveneSignificant ? '方差不齐' : '方差齐性') : 'N/A'}
                                </div>
                            </div>
                        </div>
                    </div></div>
                </div>

                <div style="margin-top:1rem;padding:1rem;background:var(--light-cyan);border-radius:10px;font-size:0.875rem;color:var(--text-secondary);">
                    <i class="fas fa-info-circle"></i>
                    <strong>完整报告</strong>（描述性统计、ANOVA表、LSD、Duncan）可点击下载Excel
                </div>
            </div>
        `;
    });

    elements.resultsContainer.innerHTML = html;

    // 滚动到结果区域
    elements.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    console.log('Results displayed');
}

function toggleCollapse(header) {
    const panel = header.parentElement;
    panel.classList.toggle('open');
}

async function downloadReport() {
    if (!state.currentFile || !state.currentFileData) {
        showAlert('请先上传并分析文件', 'error');
        return;
    }

    console.log('Downloading report...');
    showLoading(true);

    // 使用存储的文件数据创建新的 File 对象
    const file = new File([state.currentFileData], state.currentFile.name, { type: state.currentFile.type });

    const formData = new FormData();
    formData.append('file', file);
    formData.append('sample_column', elements.sampleColumn.value);

    // 获取选中的指标列（多选框）
    const checkedBoxes = elements.valueColumnsContainer.querySelectorAll('input[type="checkbox"]:checked');
    const selectedIndicators = Array.from(checkedBoxes).map(cb => cb.value);
    if (selectedIndicators.length > 0) {
        formData.append('value_columns', selectedIndicators.join(','));
    }

    // 添加合并平行样参数
    const mergeParallel = elements.mergeParallel ? elements.mergeParallel.checked : false;
    const mergeSuffixLength = elements.mergeSuffixLength ? parseInt(elements.mergeSuffixLength.value) || 1 : 1;
    formData.append('merge_parallel', mergeParallel);
    formData.append('merge_suffix_length', mergeSuffixLength);
    console.log('Download report - Selected indicators:', selectedIndicators, 'Merge parallel:', mergeParallel, 'Suffix length:', mergeSuffixLength);

    // 创建 AbortController 用于超时控制
    const controller = new AbortController();
    const timeoutId = setTimeout(() => {
        console.log('Download timeout, aborting...');
        controller.abort();
    }, 60000); // 60秒超时

    try {
        const response = await fetch('/download_report', {
            method: 'POST',
            body: formData,
            signal: controller.signal
        });

        clearTimeout(timeoutId);

        if (!response.ok) {
            let errorMessage = `下载失败: ${response.status}`;
            try {
                const error = await response.json();
                if (error.error) errorMessage = error.error;
            } catch (e) {}
            throw new Error(errorMessage);
        }

        // 下载文件
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ANOVA_分析报告.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        showAlert('报告下载成功！', 'success');

    } catch (error) {
        clearTimeout(timeoutId);
        console.error('Download error:', error);

        if (error.name === 'AbortError') {
            showAlert('下载超时，请检查服务器是否正常运行', 'error');
        } else if (error.name === 'TypeError' && error.message.includes('fetch')) {
            showAlert('无法连接到服务器', 'error');
        } else {
            showAlert('下载失败: ' + error.message, 'error');
        }
    } finally {
        showLoading(false);
    }
}

function showAlert(message, type) {
    console.log(`Alert [${type}]:`, message);

    const alertDiv = document.createElement('div');

    // 根据类型选择图标
    let iconClass = 'fa-check-circle';
    if (type === 'error') iconClass = 'fa-exclamation-circle';
    if (type === 'warning') iconClass = 'fa-exclamation-triangle';

    alertDiv.className = `alert alert-${type}`;
    alertDiv.innerHTML = `
        <i class="fas ${iconClass} alert-icon"></i>
        <span>${escapeHtml(message)}</span>
    `;

    elements.alertContainer.appendChild(alertDiv);

    // 3秒后自动移除
    setTimeout(() => {
        if (alertDiv.parentNode) {
            alertDiv.remove();
        }
    }, 5000);
}

function showLoading(show) {
    if (show) {
        elements.loadingOverlay.classList.add('show');
    } else {
        elements.loadingOverlay.classList.remove('show');
    }
}

function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

// ─────────────────────────────────────────────
// 绘图模块
// ─────────────────────────────────────────────

const plotState = {
    seriesCount: 0,
    lastImageB64: null,
    lastExcelB64: null,
    lastOpjuB64: null
};

function initPlotModule() {
    const enablePlot = document.getElementById('enablePlot');
    const plotConfig = document.getElementById('plotConfig');
    const addSeriesBtn = document.getElementById('addSeriesBtn');
    const generatePlotBtn = document.getElementById('generatePlotBtn');
    const downloadPlotBtn = document.getElementById('downloadPlotBtn');
    const downloadPlotExcelBtn = document.getElementById('downloadPlotExcelBtn');

    const downloadPlotOpjuBtn = document.getElementById('downloadPlotOpjuBtn');

    if (!enablePlot) return;

    enablePlot.addEventListener('change', () => {
        plotConfig.classList.toggle('show', enablePlot.checked);
        if (enablePlot.checked && plotState.seriesCount === 0) {
            addSeries();  // 默认加一行
        }
    });

    addSeriesBtn.addEventListener('click', addSeries);

    // 取色器与文本框双向同步
    ['Min', 'Mid', 'Max'].forEach(k => {
        const picker = document.getElementById(`heatmapColor${k}`);
        const hex = document.getElementById(`heatmapColor${k}Hex`);
        if (picker && hex) {
            picker.addEventListener('input', () => { hex.value = picker.value.toUpperCase(); });
            hex.addEventListener('input', () => { if (/^#[0-9A-Fa-f]{6}$/.test(hex.value)) picker.value = hex.value; });
        }
    });

    // 图表类型切换
    const chartTypeRadios = document.querySelectorAll('input[name="chartType"]');
    chartTypeRadios.forEach(radio => {
        radio.addEventListener('change', () => {
            updateSeriesHeaderForChartType(radio.value);

            const isHeatmap = radio.value === 'heatmap';
            const isRadar = radio.value === 'radar';
            const isPca = radio.value === 'pca';

            // 清空现有系列并重新添加（热图/PCA不需要系列配置）
            const container = document.getElementById('seriesContainer');
            if (container && !isHeatmap && !isPca) {
                container.innerHTML = '';
                plotState.seriesCount = 0;
                addSeries();  // 添加一个默认系列
            }

            // 系列配置区域（热图/PCA不需要）
            const seriesSection = document.getElementById('seriesSection');
            if (seriesSection) {
                seriesSection.style.display = (isHeatmap || isPca) ? 'none' : 'block';
            }

            // 基础轴配置（热图/PCA不需要）
            const axisConfigSection = document.getElementById('axisConfigSection');
            if (axisConfigSection) {
                axisConfigSection.style.display = (isHeatmap || isPca) ? 'none' : 'grid';
            }

            // 网格线开关（雷达图和热图不需要）
            const gridToggle = document.getElementById('gridToggleSection');
            if (gridToggle) {
                gridToggle.style.display = (isRadar || isHeatmap) ? 'none' : 'block';
            }

            // 柱状图组间距（仅柱状图显示）
            const barSpacingSection = document.getElementById('barSpacingSection');
            if (barSpacingSection) {
                barSpacingSection.style.display = (radio.value === 'bar') ? 'block' : 'none';
            }

            // 雷达图转置选项
            const transposeSection = document.getElementById('radarTransposeSection');
            if (transposeSection) {
                transposeSection.style.display = isRadar ? 'block' : 'none';
            }

            // 雷达图范围选项
            const rangeSection = document.getElementById('radarRangeSection');
            if (rangeSection) {
                rangeSection.style.display = isRadar ? 'block' : 'none';
            }

            // 热图配置
            const heatmapSection = document.getElementById('heatmapSection');
            if (heatmapSection) {
                heatmapSection.style.display = isHeatmap ? 'block' : 'none';
            }

            // PCA 配置
            const pcaSection = document.getElementById('pcaSection');
            if (pcaSection) {
                pcaSection.style.display = isPca ? 'block' : 'none';
            }

            // 热图/PCA不需要图例和数据标签字号
            const fsLegendGroup = document.getElementById('fsLegendGroup');
            const fsDataLabelGroup = document.getElementById('fsDataLabelGroup');
            if (fsLegendGroup) fsLegendGroup.style.display = isHeatmap ? 'none' : '';
            if (fsDataLabelGroup) fsDataLabelGroup.style.display = isHeatmap ? 'none' : '';
        });
    });

    // 柱状图组间距滑块同步
    const barSpacingSlider = document.getElementById('barGroupSpacing');
    const barSpacingInput = document.getElementById('barGroupSpacingVal');
    if (barSpacingSlider && barSpacingInput) {
        barSpacingSlider.addEventListener('input', () => { barSpacingInput.value = barSpacingSlider.value; });
        barSpacingInput.addEventListener('input', () => { barSpacingSlider.value = barSpacingInput.value; });
    }

    generatePlotBtn.addEventListener('click', generatePlot);

    downloadPlotBtn.addEventListener('click', () => {
        if (!plotState.lastImageB64) return;
        const a = document.createElement('a');
        a.href = 'data:image/png;base64,' + plotState.lastImageB64;
        a.download = '图表.png';
        a.click();
    });

    downloadPlotExcelBtn.addEventListener('click', () => {
        if (!plotState.lastExcelB64) return;
        const a = document.createElement('a');
        a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + plotState.lastExcelB64;
        a.download = '图表数据.xlsx';
        a.click();
    });

    if (downloadPlotOpjuBtn) {
        downloadPlotOpjuBtn.addEventListener('click', () => {
            if (!plotState.lastOpjuB64) {
                showAlert('未能生成 Origin 图表文件', 'warning');
                return;
            }
            const a = document.createElement('a');
            a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + plotState.lastOpjuB64;
            a.download = '图表数据_Origin.xlsx';
            a.click();
        });
    }
}

function getColumnOptions(selectedVal) {
    const cols = state.columns || [];
    return cols.map(c =>
        `<option value="${escapeHtml(c)}" ${c === selectedVal ? 'selected' : ''}>${escapeHtml(c)}</option>`
    ).join('');
}

function updateSeriesHeaderForChartType(chartType) {
    const header = document.getElementById('seriesHeader');
    if (!header) return;

    if (chartType === 'line') {
        header.style.gridTemplateColumns = '1fr 1fr 1fr 100px 80px 140px 48px';
        header.innerHTML = `
            <div class="plot-small-label">Y 值列</div>
            <div class="plot-small-label">标准差列（可选）</div>
            <div class="plot-small-label">标签列（Duncan等，可选）</div>
            <div class="plot-small-label">线条样式</div>
            <div class="plot-small-label">粗细</div>
            <div class="plot-small-label">颜色</div>
            <div></div>
        `;
    } else if (chartType === 'radar') {
        header.style.gridTemplateColumns = '1fr 100px 80px 100px 140px 48px';
        header.innerHTML = `
            <div class="plot-small-label">数据列</div>
            <div class="plot-small-label">线条样式</div>
            <div class="plot-small-label">粗细</div>
            <div class="plot-small-label">标记</div>
            <div class="plot-small-label">颜色</div>
            <div></div>
        `;
    } else {
        header.style.gridTemplateColumns = '1fr 1fr 1fr 140px 48px';
        header.innerHTML = `
            <div class="plot-small-label">Y 值列</div>
            <div class="plot-small-label">标准差列（可选）</div>
            <div class="plot-small-label">标签列（Duncan等，可选）</div>
            <div class="plot-small-label">颜色</div>
            <div></div>
        `;
    }
}

function refreshPlotColumnSelectors() {
    // 更新 X 轴列选择器
    const xSel = document.getElementById('plotXCol');
    if (xSel) {
        const cur = xSel.value;
        xSel.innerHTML = '<option value="">-- 选择列 --</option>' + getColumnOptions(cur);
    }
    // 更新所有系列的列选择器
    document.querySelectorAll('.series-y-col, .series-std-col, .series-lbl-col').forEach(sel => {
        const cur = sel.value;
        sel.innerHTML = '<option value="">-- 无 --</option>' + getColumnOptions(cur);
    });
}

const DEFAULT_COLORS = ['#8FBC8F','#DAA520','#87CEEB','#E8A0A0','#9B8EC4','#F4A460'];

function addSeries() {
    const container = document.getElementById('seriesContainer');
    if (!container) return;

    const idx = plotState.seriesCount++;
    const color = DEFAULT_COLORS[idx % DEFAULT_COLORS.length];
    const cols = state.columns || [];
    const colOpts = '<option value="">-- 无 --</option>' + getColumnOptions('');

    // 检查当前图表类型
    const chartType = document.querySelector('input[name="chartType"]:checked')?.value || 'bar';
    const isLineMode = chartType === 'line';
    const isRadarMode = chartType === 'radar';

    const row = document.createElement('div');
    row.className = 'plot-series-row' + (isLineMode ? ' line-mode' : '') + (isRadarMode ? ' radar-mode' : '');
    row.dataset.idx = idx;

    if (isRadarMode) {
        // 雷达图模式：只需要数据列、线条样式、粗细、标记、颜色
        row.innerHTML = `
            <select class="form-select series-data-col">${colOpts}</select>
            <select class="form-select series-line-style" style="padding:0.35rem 0.5rem;">
                <option value="-">实线 —</option>
                <option value="--">虚线 - -</option>
                <option value="-.">点画线 -·-</option>
                <option value=":">点线 ···</option>
            </select>
            <input type="number" class="form-select series-line-width" value="2" min="0.5" max="5" step="0.5" style="padding:0.35rem 0.5rem;" placeholder="粗细">
            <select class="form-select series-marker" style="padding:0.35rem 0.5rem;">
                <option value="o">圆形 ●</option>
                <option value="s">方形 ■</option>
                <option value="^">三角 ▲</option>
                <option value="D">菱形 ◆</option>
                <option value="v">倒三角 ▼</option>
                <option value="<">左三角 ◀</option>
                <option value=">">右三角 ▶</option>
            </select>
            <div style="display:flex;align-items:center;gap:4px;">
                <input type="color" class="color-swatch series-color" value="${color}" style="width:38px;height:38px;padding:2px;cursor:pointer;border:1px solid #ccc;border-radius:6px;">
                <input type="text" class="series-color-hex form-select" value="${color}" style="width:82px;font-size:0.78rem;padding:0.3rem 0.4rem;" maxlength="7" placeholder="#rrggbb">
            </div>
            <button class="remove-series" title="删除此系列"><i class="fas fa-times"></i></button>
        `;
    } else if (isLineMode) {
        // 折线图模式：添加线条样式和粗细控件
        row.innerHTML = `
            <select class="form-select series-y-col">${colOpts}</select>
            <select class="form-select series-std-col">${colOpts}</select>
            <select class="form-select series-lbl-col">${colOpts}</select>
            <select class="form-select series-line-style" style="padding:0.35rem 0.5rem;">
                <option value="-">实线 —</option>
                <option value="--">虚线 - -</option>
                <option value="-.">点画线 -·-</option>
                <option value=":">点线 ···</option>
            </select>
            <input type="number" class="form-select series-line-width" value="2" min="0.5" max="5" step="0.5" style="padding:0.35rem 0.5rem;" placeholder="粗细">
            <div style="display:flex;align-items:center;gap:4px;">
                <input type="color" class="color-swatch series-color" value="${color}" style="width:38px;height:38px;padding:2px;cursor:pointer;border:1px solid #ccc;border-radius:6px;">
                <input type="text" class="series-color-hex form-select" value="${color}" style="width:82px;font-size:0.78rem;padding:0.3rem 0.4rem;" maxlength="7" placeholder="#rrggbb">
            </div>
            <button class="remove-series" title="删除此系列"><i class="fas fa-times"></i></button>
        `;
    } else {
        // 柱状图模式：原有布局
        row.innerHTML = `
            <select class="form-select series-y-col">${colOpts}</select>
            <select class="form-select series-std-col">${colOpts}</select>
            <select class="form-select series-lbl-col">${colOpts}</select>
            <div style="display:flex;align-items:center;gap:4px;">
                <input type="color" class="color-swatch series-color" value="${color}" style="width:38px;height:38px;padding:2px;cursor:pointer;border:1px solid #ccc;border-radius:6px;">
                <input type="text" class="series-color-hex form-select" value="${color}" style="width:82px;font-size:0.78rem;padding:0.3rem 0.4rem;" maxlength="7" placeholder="#rrggbb">
            </div>
            <button class="remove-series" title="删除此系列"><i class="fas fa-times"></i></button>
        `;
    }

    // 同步颜色选择器 ↔ 文本框
    const colorInput = row.querySelector('.series-color');
    const hexInput = row.querySelector('.series-color-hex');
    colorInput.addEventListener('input', () => { hexInput.value = colorInput.value; });
    hexInput.addEventListener('input', () => {
        const v = hexInput.value.trim();
        if (/^#[0-9a-fA-F]{6}$/.test(v)) colorInput.value = v;
    });
    row.querySelector('.remove-series').addEventListener('click', () => {
        row.remove();
    });
    container.appendChild(row);
}

async function generatePlot() {
    if (!state.previewData || state.previewData.length === 0) {
        showAlert('请先上传数据文件', 'error');
        return;
    }

    // 获取图表类型（提前，PCA 不需要 xCol）
    const chartType = document.querySelector('input[name="chartType"]:checked')?.value || 'bar';
    const boldConfig = {
        title: document.getElementById('boldTitle')?.checked || false,
        axis_label: document.getElementById('boldAxisLabel')?.checked || false,
        tick: document.getElementById('boldTick')?.checked || false,
        legend: document.getElementById('boldLegend')?.checked || false,
        data_label: document.getElementById('boldDataLabel')?.checked || false,
    };

    if (chartType === 'pca') {
        const valueCols = Array.from(document.getElementById('pcaValueCols')?.querySelectorAll('input[type="checkbox"]:checked') || []).map(cb => cb.value);
        const pcX = parseInt(document.getElementById('pcaCompX')?.value) || 1;
        const pcY = parseInt(document.getElementById('pcaCompY')?.value) || 2;
        const showEllipse = document.getElementById('pcaShowEllipse')?.checked ?? true;
        const showLabels = document.getElementById('pcaShowLabels')?.checked ?? true;
        const showGrid = document.getElementById('pcaShowGrid')?.checked ?? true;

        // 收集分组信息（从勾选框读取）
        const groupRows = document.querySelectorAll('#pcaGroupRows .pca-group-row');
        const groupsMap = Array.from(groupRows).map(row => ({
            name: row.querySelector('.pca-grp-name')?.value || '',
            color: row.querySelector('.pca-grp-color')?.value || '#1f77b4',
            indices: Array.from(row.querySelectorAll('.pca-row-check:checked')).map(cb => parseInt(cb.value))
        })).filter(g => g.indices.length > 0);

        const allData = state.allData && state.allData.length > 0 ? state.allData : state.previewData;

        const toNum = (id) => { const v = document.getElementById(id)?.value; return v !== '' && v != null ? parseFloat(v) : null; };

        const payload = {
            chart_type: 'pca',
            data: allData,
            value_cols: valueCols,
            sample_col: document.getElementById('sampleColumn')?.value || '',
            groups_map: groupsMap,
            pc_x: pcX,
            pc_y: pcY,
            xlabel: document.getElementById('pcaXLabel')?.value || '',
            ylabel: document.getElementById('pcaYLabel')?.value || '',
            show_ellipse: showEllipse,
            show_labels: showLabels,
            label_mode: document.getElementById('pcaLabelMode')?.value || 'auto',
            show_grid: showGrid,
            x_min: toNum('pcaXMin'), x_max: toNum('pcaXMax'),
            y_min: toNum('pcaYMin'), y_max: toNum('pcaYMax'),
            axis_fontsize: parseInt(document.getElementById('pcaAxisFontSize')?.value) || 13,
            axis_color: document.getElementById('pcaAxisColor')?.value || '#000000',
            tick_fontsize: parseInt(document.getElementById('pcaTickFontSize')?.value) || 11,
            tick_color: document.getElementById('pcaTickColor')?.value || '#000000',
            label_fontsize: parseInt(document.getElementById('pcaLabelFontSize')?.value) || 8,
            legend_fontsize: parseInt(document.getElementById('pcaLegendFontSize')?.value) || 10,
            dot_size: parseInt(document.getElementById('pcaDotSize')?.value) || 60,
        };

        showLoading(true);
        try {
            const resp = await fetch('/api/plot', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await resp.json();
            if (result.error) throw new Error(result.error);

            plotState.lastImageB64 = result.image;
            plotState.lastExcelB64 = result.excel;
            plotState.lastOpjuB64 = null;

            document.getElementById('plotResult').innerHTML = `<img class="plot-result-img" src="data:image/png;base64,${result.image}" alt="PCA图">`;
            document.getElementById('downloadPlotBtn').style.display = 'inline-flex';
            document.getElementById('downloadPlotExcelBtn').style.display = 'inline-flex';
            const opjuBtn = document.getElementById('downloadPlotOpjuBtn');
            if (opjuBtn) opjuBtn.style.display = 'none';
            showAlert('PCA 图生成成功！', 'success');
        } catch (e) {
            showAlert('生成 PCA 图失败: ' + e.message, 'error');
        } finally {
            showLoading(false);
        }
        return;
    }

    const xCol = document.getElementById('plotXCol').value;
    if (!xCol) { showAlert('请选择 X 轴数据列', 'error'); return; }

    const rows = document.querySelectorAll('#seriesContainer .plot-series-row');

    if (chartType === 'heatmap') {
        const rowCol = document.getElementById('heatmapRowCol')?.value || '';
        const valueCols = Array.from(document.getElementById('heatmapValueCols')?.querySelectorAll('input[type="checkbox"]:checked') || []).map(cb => cb.value);
        if (valueCols.length === 0) { showAlert('请选择至少一个数值列', 'error'); return; }
        const normalize = document.getElementById('heatmapNormalize')?.checked ?? true;
        const clusterRows = document.getElementById('heatmapClusterRows')?.checked ?? true;
        const clusterCols = document.getElementById('heatmapClusterCols')?.checked ?? true;
        showLoading(true);
        try {
            const payload = {
                chart_type: 'heatmap',
                data: state.allData,
                row_col: rowCol,
                value_cols: valueCols,
                normalize,
                cluster_rows: clusterRows,
                cluster_cols: clusterCols,
                bold: boldConfig,
                title: '',
                vmin: parseFloat(document.getElementById('heatmapVmin')?.value) || -2,
                vmax: parseFloat(document.getElementById('heatmapVmax')?.value) || 2,
                vstep: parseFloat(document.getElementById('heatmapVstep')?.value) || 0.5,
                color_min: document.getElementById('heatmapColorMinHex')?.value || '#1446AF',
                color_mid: document.getElementById('heatmapColorMidHex')?.value || '#EDE7AD',
                color_max: document.getElementById('heatmapColorMaxHex')?.value || '#CF1C1D',
                cbar_fontsize: parseInt(document.getElementById('heatmapCbarFontSize')?.value) || 7,
                cbar_bold: document.getElementById('heatmapCbarBold')?.checked || false,
                font_sizes: {
                    title: parseInt(document.getElementById('fsTitleSize')?.value) || 12,
                    axis_label: parseInt(document.getElementById('fsAxisLabel')?.value) || 10,
                    tick: parseInt(document.getElementById('fsTick')?.value) || 8,
                    legend: parseInt(document.getElementById('fsLegend')?.value) || 9
                }
            };
            const resp = await fetch('/api/plot', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await resp.json();
            if (!resp.ok || result.error) { showAlert('生成图表失败: ' + (result.error || '未知错误'), 'error'); return; }
            plotState.lastImageB64 = result.image;
            plotState.lastExcelB64 = result.excel;
            plotState.lastOpjuB64 = result.opju || null;
            document.getElementById('plotResult').innerHTML = `<img class="plot-result-img" src="data:image/png;base64,${result.image}" alt="图表">`;
            document.getElementById('downloadPlotBtn').style.display = 'inline-flex';
            document.getElementById('downloadPlotExcelBtn').style.display = 'inline-flex';
            const opjuBtn2 = document.getElementById('downloadPlotOpjuBtn');
            if (opjuBtn2) opjuBtn2.style.display = result.opju ? 'inline-flex' : 'none';
            showAlert('图表生成成功！', 'success');
        } catch(e) {
            showAlert('请求失败: ' + e.message, 'error');
        } finally {
            showLoading(false);
        }
        return;
    }

    if (rows.length === 0) { showAlert('请至少添加一个 Y 轴系列', 'error'); return; }

    if (chartType === 'radar') {
        // 雷达图模式
        const seriesCols = [], colors = [], lineStyles = [], lineWidths = [], markerStyles = [];
        rows.forEach(row => {
            const dataCol = row.querySelector('.series-data-col');
            if (dataCol) seriesCols.push(dataCol.value);
            colors.push(row.querySelector('.series-color').value);
            const styleSelect = row.querySelector('.series-line-style');
            const widthInput = row.querySelector('.series-line-width');
            const markerSelect = row.querySelector('.series-marker');
            lineStyles.push(styleSelect ? styleSelect.value : '-');
            lineWidths.push(widthInput ? widthInput.value : '2');
            markerStyles.push(markerSelect ? markerSelect.value : 'o');
        });

        if (seriesCols.some(c => !c)) { showAlert('请为每个系列选择数据列', 'error'); return; }

        // 检查是否需要转置数据
        const shouldTranspose = document.getElementById('transposeRadarData')?.checked || false;
        const yMinVal = document.getElementById('radarYMin')?.value;
        const yMaxVal = document.getElementById('radarYMax')?.value;
        const yStepVal = document.getElementById('radarYStep')?.value;

        showLoading(true);
        try {
            const payload = {
                chart_type: 'radar',
                data: state.previewData,
                axes_col: xCol,
                series_cols: seriesCols,
                colors: colors,
                line_styles: lineStyles,
                line_widths: lineWidths,
                marker_styles: markerStyles,
                transpose: shouldTranspose,
                y_min: yMinVal !== '' ? parseFloat(yMinVal) : null,
                y_max: yMaxVal !== '' ? parseFloat(yMaxVal) : null,
                y_step: yStepVal !== '' ? parseFloat(yStepVal) : null,
                grid_color: document.getElementById('radarGridColor')?.value || '#808080',
                grid_width: parseFloat(document.getElementById('radarGridWidth')?.value) || 0.8,
                spoke_color: document.getElementById('radarSpokeColor')?.value || '#808080',
                spoke_width: parseFloat(document.getElementById('radarSpokeWidth')?.value) || 1.2,
                axis_label_pad: parseFloat(document.getElementById('radarAxisLabelPad')?.value) || 20,
                tick_label_angle: parseFloat(document.getElementById('radarTickAngle')?.value) || 45,
                bold: boldConfig,
                title: '',
                font_sizes: {
                    title: parseInt(document.getElementById('fsTitleSize')?.value) || 14,
                    axis_label: parseInt(document.getElementById('radarAxisLabelSize')?.value) || parseInt(document.getElementById('fsAxisLabel')?.value) || 11,
                    tick: parseInt(document.getElementById('radarTickSize')?.value) || parseInt(document.getElementById('fsTick')?.value) || 10,
                    legend: parseInt(document.getElementById('fsLegend')?.value) || 10
                }
            };

            const resp = await fetch('/api/plot', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const result = await resp.json();
            if (result.error) throw new Error(result.error);

            plotState.lastImageB64 = result.image;
            plotState.lastExcelB64 = result.excel;
            plotState.lastOpjuB64 = result.opju || null;

            const plotResult = document.getElementById('plotResult');
            plotResult.innerHTML = `<img class="plot-result-img" src="data:image/png;base64,${result.image}" alt="图表">`;

            document.getElementById('downloadPlotBtn').style.display = 'inline-flex';
            document.getElementById('downloadPlotExcelBtn').style.display = 'inline-flex';
            const opjuBtn = document.getElementById('downloadPlotOpjuBtn');
            if (opjuBtn) {
                opjuBtn.style.display = result.opju ? 'inline-flex' : 'none';
            }

            showAlert('图表生成成功！', 'success');
        } catch (e) {
            showAlert('生成图表失败: ' + e.message, 'error');
        } finally {
            showLoading(false);
        }
        return;
    }

    // 柱状图和折线图模式
    const yCols = [], stdCols = [], lblCols = [], colors = [];
    const lineStyles = [], lineWidths = [];
    rows.forEach(row => {
        yCols.push(row.querySelector('.series-y-col').value);
        stdCols.push(row.querySelector('.series-std-col').value);
        lblCols.push(row.querySelector('.series-lbl-col').value);
        colors.push(row.querySelector('.series-color').value);

        // 折线图特有参数
        if (chartType === 'line') {
            const styleSelect = row.querySelector('.series-line-style');
            const widthInput = row.querySelector('.series-line-width');
            lineStyles.push(styleSelect ? styleSelect.value : '-');
            lineWidths.push(widthInput ? widthInput.value : '2');
        }
    });

    if (yCols.some(c => !c)) { showAlert('请为每个系列选择 Y 值列', 'error'); return; }

    // 用完整数据（previewData 只有前10行，需要用分析结果的汇总数据）
    // 优先用 summary_table，否则用 previewData
    let plotData = state.previewData;
    let finalYCols = [...yCols];
    let finalStdCols = [...stdCols];
    let finalLblCols = [...lblCols];

    if (state.analysisResults && state.analysisResults.summary_table && state.analysisResults.summary_table.length > 0) {
        const st = state.analysisResults.summary_table;
        const sampleCol = elements.sampleColumn.value;
        if (xCol === sampleCol) {
            // summary_table 的列名格式是 {indicator}_mean / _std / _duncan
            // 把用户选的原始列名映射过去
            plotData = st.map(r => {
                const obj = { [xCol]: r.sample };
                Object.keys(r).forEach(k => { if (k !== 'sample') obj[k] = r[k]; });
                return obj;
            });
            // 重新映射列名：如果用户选了 '300'，数据里对应 '300_mean'
            finalYCols = yCols.map(c => {
                if (plotData[0] && !(c in plotData[0]) && (c + '_mean') in plotData[0]) return c + '_mean';
                return c;
            });
            finalStdCols = stdCols.map((c, i) => {
                if (!c) return c;
                if (plotData[0] && !(c in plotData[0]) && (c + '_std') in plotData[0]) return c + '_std';
                return c;
            });
            finalLblCols = lblCols.map((c, i) => {
                if (!c) return c;
                if (plotData[0] && !(c in plotData[0]) && (c + '_duncan') in plotData[0]) return c + '_duncan';
                return c;
            });
        }
    }

    showLoading(true);
    try {
        const payload = {
            chart_type: chartType,
            data: plotData,
            x_col: xCol,
            y_cols: finalYCols,
            std_cols: finalStdCols,
            label_cols: finalLblCols,
            colors: colors,
            x_label: document.getElementById('plotXLabel').value,
            y_label: document.getElementById('plotYLabel').value,
            show_grid: document.getElementById('showGridLines')?.checked || false,
            bar_inner_gap: parseFloat(document.getElementById('barGroupSpacingVal')?.value) || 0.9,
            bold: boldConfig,
            font_sizes: {
                title:      parseInt(document.getElementById('fsTitleSize')?.value) || 14,
                axis_label: parseInt(document.getElementById('fsAxisLabel')?.value) || 13,
                tick:       parseInt(document.getElementById('fsTick')?.value) || 11,
                legend:     parseInt(document.getElementById('fsLegend')?.value) || 10,
                data_label: parseInt(document.getElementById('fsDataLabel')?.value) || 9
            }
        };

        // 折线图特有参数
        if (chartType === 'line') {
            payload.line_styles = lineStyles;
            payload.line_widths = lineWidths;
        }

        const resp = await fetch('/api/plot', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const result = await resp.json();
        if (result.error) throw new Error(result.error);

        plotState.lastImageB64 = result.image;
        plotState.lastExcelB64 = result.excel;
        plotState.lastOpjuB64 = result.opju || null;

        const plotResult = document.getElementById('plotResult');
        plotResult.innerHTML = `<img class="plot-result-img" src="data:image/png;base64,${result.image}" alt="图表">`;

        document.getElementById('downloadPlotBtn').style.display = 'inline-flex';
        document.getElementById('downloadPlotExcelBtn').style.display = 'inline-flex';
        const opjuBtn = document.getElementById('downloadPlotOpjuBtn');
        if (opjuBtn) {
            opjuBtn.style.display = result.opju ? 'inline-flex' : 'none';
        }

        showAlert('图表生成成功！', 'success');
    } catch (e) {
        showAlert('生成图表失败: ' + e.message, 'error');
    } finally {
        showLoading(false);
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initPlotModule();
});
