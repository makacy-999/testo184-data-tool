/**
 * Testo 184 温度计数据读取汇总工具 - 前端交互逻辑
 */

// ─── 全局状态 ──────────────────────────────────────────────────────────────
const state = {
    sessionId: null,
    sessionName: '',
    points: [],           // [{point_number, serial_number, status}]
    currentStep: 1,
    currentPointIndex: 0,
    detectedDevice: null,
    completedDevices: [], // [{point_number, serial_number, record_count}]
    totalRecords: 0,
};

// ─── API 工具 ──────────────────────────────────────────────────────────────
async function api(url, options = {}) {
    const headers = { 'Content-Type': 'application/json' };
    // 若传 FormData，不设置 JSON header 也不 JSON 序列化
    if (options.isForm) {
        delete options.isForm;
        const res = await fetch(url, {
            ...options,
            body: options.body, // FormData / File 等直接传
        });
        if (url.endsWith('/export')) return res;
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || `请求失败 (${res.status})`);
        return data;
    }
    const res = await fetch(url, {
        headers,
        ...options,
        body: options.body ? JSON.stringify(options.body) : undefined,
    });
    if (url.endsWith('/export')) return res; // 文件下载不解析 JSON
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `请求失败 (${res.status})`);
    return data;
}

// ─── Toast 通知 ────────────────────────────────────────────────────────────
function toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    el.textContent = message;
    container.appendChild(el);
    setTimeout(() => el.remove(), 4000);
}

// ─── 状态更新 ──────────────────────────────────────────────────────────────
function setStatus(text) {
    document.getElementById('status-text').textContent = text;
}

// ─── 步骤导航 ──────────────────────────────────────────────────────────────
function goToStep(step) {
    state.currentStep = step;
    // 更新步骤条
    document.querySelectorAll('.step-item').forEach(el => {
        const s = parseInt(el.dataset.step);
        el.classList.remove('active', 'done');
        if (s === step) el.classList.add('active');
        else if (s < step) el.classList.add('done');
    });
    // 切换面板
    document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
    document.getElementById(`panel-step${step}`).classList.add('active');
}

// ─── Step 1: 配置测点 ─────────────────────────────────────────────────────
function generatePoints() {
    const count = parseInt(document.getElementById('point-count').value) || 1;
    const grid = document.getElementById('points-grid');
    grid.innerHTML = '';
    state.points = [];

    for (let i = 1; i <= count; i++) {
        const div = document.createElement('div');
        div.className = 'point-item';
        div.innerHTML = `
            <span class="point-num">#${i}</span>
            <input type="text" placeholder="测点名称（可选）" value="${i}" data-index="${i - 1}">
        `;
        grid.appendChild(div);
        state.points.push({ point_number: String(i), serial_number: '' });
    }

    // 监听输入变化
    grid.querySelectorAll('input').forEach(input => {
        input.addEventListener('input', (e) => {
            const idx = parseInt(e.target.dataset.index);
            state.points[idx].point_number = e.target.value || String(idx + 1);
        });
    });
}

async function startReading() {
    // 收集当前测点数据
    const grid = document.getElementById('points-grid');
    const inputs = grid.querySelectorAll('input');
    inputs.forEach((input, idx) => {
        if (state.points[idx]) {
            state.points[idx].point_number = input.value || String(idx + 1);
        }
    });

    if (state.points.length === 0) {
        toast('请先生成测点', 'warning');
        return;
    }

    setStatus('正在创建会话...');
    try {
        const sessionName = document.getElementById('session-name').value;
        const session = await api('/api/sessions', {
            method: 'POST',
            body: { name: sessionName, points: state.points },
        });
        state.sessionId = session.id;
        state.sessionName = session.name;
        state.points = session.points;
        state.currentPointIndex = 0;
        state.completedDevices = [];
        state.totalRecords = 0;

        document.getElementById('session-info').textContent = `会话: ${session.name}`;
        toast('会话已创建，开始读取设备', 'success');

        // 更新 Step 2 UI
        updateReadUI();
        goToStep(2);
        setStatus('就绪 - 请插入第一个温度计');
    } catch (e) {
        toast(`创建失败: ${e.message}`, 'error');
        setStatus('创建会话失败');
    }
}

// ─── Step 2: 读取设备 ─────────────────────────────────────────────────────
function updateReadUI() {
    const total = state.points.length;
    const done = state.currentPointIndex;
    const pct = total > 0 ? (done / total * 100) : 0;

    document.getElementById('read-progress-fill').style.width = `${pct}%`;
    document.getElementById('read-progress-text').textContent = `${done} / ${total} 已读取`;

    if (done < total) {
        const pt = state.points[done];
        document.getElementById('current-point-badge').textContent = `测点 #${pt.point_number}`;
        document.getElementById('current-point-status').textContent = '请插入温度计，然后点击「扫描设备」';
        document.getElementById('btn-scan').style.display = '';
        document.getElementById('btn-confirm-read').style.display = 'none';
        document.getElementById('device-result').style.display = 'none';
    } else {
        document.getElementById('current-point-badge').textContent = '✅ 全部完成';
        document.getElementById('current-point-status').textContent = '所有设备已读取完成！';
        document.getElementById('btn-scan').style.display = 'none';
        document.getElementById('btn-confirm-read').style.display = 'none';
        document.getElementById('device-result').style.display = 'none';
        // 自动跳到 Step 3
        setTimeout(() => {
            buildSummary();
            goToStep(3);
        }, 1500);
    }

    // 更新已完成列表
    const doneList = document.getElementById('read-devices-list');
    const doneItems = document.getElementById('devices-done-list');
    if (state.completedDevices.length > 0) {
        doneList.style.display = '';
        doneItems.innerHTML = state.completedDevices.map(d => `
            <div class="done-device-item">
                <div class="done-device-info">
                    <span class="done-badge">测点 #${d.point_number}</span>
                    <span>SN: ${d.serial_number}</span>
                    <span style="color:var(--text-muted)">${d.record_count} 条数据</span>
                </div>
                <span style="color:var(--success)">✅</span>
            </div>
        `).join('');
    }
}

async function scanDevice() {
    const btn = document.getElementById('btn-scan');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading-spinner"></span> 扫描中...';
    setStatus('正在检测设备...');

    try {
        const result = await api(`/api/sessions/${state.sessionId}/detect`, { method: 'POST' });

        btn.disabled = false;
        btn.innerHTML = '🔍 扫描设备';

        if (result.detected) {
            state.detectedDevice = result.device;
            const dev = result.device;

            // 显示设备信息
            document.getElementById('dev-name').textContent = dev.name;
            document.getElementById('dev-sn').value = dev.serial_number || '';
            document.getElementById('dev-count').textContent = `${dev.record_count} 条`;
            document.getElementById('dev-files').textContent = dev.csv_files.join(', ') || '-';

            // 数据预览
            const table = document.getElementById('preview-table');
            const thead = table.querySelector('thead tr');
            const tbody = table.querySelector('tbody');

            if (dev.preview && dev.preview.length > 0) {
                // 构建表头
                const keys = ['date', 'time', 'temperature', 'humidity', 'alarm'];
                const labels = { date: '日期', time: '时间', temperature: '温度(°C)', humidity: '湿度(%)', alarm: '报警' };
                thead.innerHTML = keys.map(k => `<th>${labels[k]}</th>`).join('');

                // 构建数据行
                tbody.innerHTML = dev.preview.map(rec =>
                    `<tr>${keys.map(k => `<td>${rec[k] || '-'}</td>`).join('')}</tr>`
                ).join('');
            } else {
                thead.innerHTML = '<th>暂无数据预览</th>';
                tbody.innerHTML = '';
            }

            document.getElementById('device-result').style.display = '';
            document.getElementById('btn-confirm-read').style.display = '';
            // 设备识别成功但没读到数据：引导用户改用 vi2 导入（Testo 设备数据常需 ComSoft 导出）
            const vi2Hint = document.getElementById('vi2-hint');
            const vi2Box = document.getElementById('vi2-import-box');
            if (dev.record_count && dev.record_count > 0) {
                document.getElementById('current-point-status').textContent = '检测到设备！请确认后点击「确认读取」';
                if (vi2Hint) vi2Hint.style.display = 'none';
            } else {
                document.getElementById('current-point-status').textContent =
                    '检测到设备，但未读取出温度数据。请尝试下方的「导入 .vi2 数据文件」';
                if (vi2Hint) vi2Hint.style.display = '';
                if (vi2Box) vi2Box.classList.add('highlight');
            }

            toast(`检测到设备: ${dev.name}`, 'success');
            setStatus(`检测到设备: ${dev.name}，共 ${dev.record_count} 条数据`);
        } else {
            state.detectedDevice = null;
            document.getElementById('device-result').style.display = 'none';
            document.getElementById('btn-confirm-read').style.display = 'none';
            document.getElementById('current-point-status').textContent = result.message;
            toast(result.message, 'warning');
            setStatus('未检测到设备');
        }
    } catch (e) {
        btn.disabled = false;
        btn.innerHTML = '🔍 扫描设备';
        toast(`扫描失败: ${e.message}`, 'error');
        setStatus('扫描失败');
    }
}

async function confirmRead() {
    if (!state.detectedDevice) {
        toast('未检测到设备', 'warning');
        return;
    }

    const btn = document.getElementById('btn-confirm-read');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading-spinner"></span> 读取中...';
    setStatus('正在读取数据并保存...');

    try {
        const result = await api(`/api/sessions/${state.sessionId}/read`, {
            method: 'POST',
            body: {
                device_path: state.detectedDevice.path,
                serial_number: document.getElementById('dev-sn').value.trim() || state.detectedDevice.serial_number || '',
            },
        });

        // 更新状态
        state.points[state.currentPointIndex].serial_number = result.serial_number;
        state.points[state.currentPointIndex].status = 'completed';
        state.completedDevices.push({
            point_number: result.point_number,
            serial_number: result.serial_number,
            record_count: result.record_count,
        });
        state.totalRecords += result.record_count;
        state.currentPointIndex++;

        toast(result.message, 'success');
        setStatus(result.message);

        // 更新 UI
        btn.disabled = false;
        btn.innerHTML = '✅ 确认读取此设备';
        state.detectedDevice = null;
        updateReadUI();

    } catch (e) {
        btn.disabled = false;
        btn.innerHTML = '✅ 确认读取此设备';
        toast(`读取失败: ${e.message}`, 'error');
        setStatus('读取失败');
    }
}

// ─── Step 2: 导入 .vi2 文件 ─────────────────────────────────────────────
async function importVi2() {
    const input = document.getElementById('vi2-file-input');
    const file = input.files && input.files[0];
    const msg = document.getElementById('vi2-msg');
    const btn = document.getElementById('btn-import-vi2');
    if (!file) {
        toast('请先选择 .vi2 文件', 'warning');
        return;
    }
    if (!state.sessionId) {
        toast('请先在第一步创建会话', 'warning');
        return;
    }

    btn.disabled = true;
    btn.innerHTML = '<span class="loading-spinner"></span> 解析中...';
    if (msg) msg.textContent = '正在解析 .vi2 文件...';

    const form = new FormData();
    form.append('file', file);

    try {
        const result = await api(`/api/sessions/${state.sessionId}/import-vi2`, {
            method: 'POST',
            body: form,
            isForm: true,
        });

        state.points[state.currentPointIndex].serial_number = result.serial_number;
        state.points[state.currentPointIndex].status = 'completed';
        state.completedDevices.push({
            point_number: result.point_number,
            serial_number: result.serial_number,
            record_count: result.record_count,
        });
        state.totalRecords += result.record_count;
        state.currentPointIndex++;

        if (msg) {
            msg.style.color = '#2e7d32';
            msg.textContent = `${result.message}｜采样间隔 ${result.sample_minutes} 分钟｜起始 ${result.start_time}`;
        }
        toast(result.message, 'success');
        setStatus(result.message);

        // 清空选择，方便导入下一个
        input.value = '';
        btn.disabled = false;
        btn.innerHTML = '📥 导入并读取数据';
        updateReadUI();
    } catch (e) {
        btn.disabled = false;
        btn.innerHTML = '📥 导入并读取数据';
        if (msg) { msg.style.color = '#c62828'; msg.textContent = `导入失败: ${e.message}`; }
        toast(`导入失败: ${e.message}`, 'error');
        setStatus('导入失败');
    }
}

// ─── Step 3: 数据汇总 ─────────────────────────────────────────────────────
async function buildSummary() {
    try {
        const session = await api(`/api/sessions/${state.sessionId}`);
        const grid = document.getElementById('summary-grid');
        const stats = document.getElementById('summary-stats');

        const pts = session.points || [];
        const totalCount = pts.reduce((sum, p) => sum + (p.record_count || 0), 0);

        grid.innerHTML = pts.map((p, i) => `
            <div class="summary-item">
                <label class="summary-check">
                    <input type="checkbox" class="point-check" data-idx="${i}" ${(p.record_count||0)>0 ? 'checked' : 'disabled'}>
                </label>
                <div class="s-point">测点 #${p.point_number}</div>
                <div class="s-sn">SN: ${p.serial_number || '未设置'}</div>
                <div class="s-count">${p.record_count || 0} 条数据</div>
                <button class="btn btn-ghost btn-sm" onclick="viewPointDetail('${p.point_number}')">查看明细</button>
            </div>
        `).join('');

        const sheetCount = 3 + (pts.filter(p => (p.record_count || 0) > 0).length);
        stats.innerHTML = `
            <div class="summary-stat-item">
                <div class="summary-stat-value">${pts.length}</div>
                <div class="summary-stat-label">设备数量</div>
            </div>
            <div class="summary-stat-item">
                <div class="summary-stat-value">${totalCount}</div>
                <div class="summary-stat-label">数据总条数</div>
            </div>
            <div class="summary-stat-item">
                <div class="summary-stat-value">${sheetCount}</div>
                <div class="summary-stat-label">Excel 工作表</div>
            </div>
        `;

        // 勾选事件：更新导出按钮状态
        document.querySelectorAll('.point-check').forEach(cb => {
            cb.addEventListener('change', updateExportState);
        });

        // 更新 Step 4 数据
        document.getElementById('export-total-points').textContent = pts.length;
        document.getElementById('export-total-records').textContent = totalCount;
        document.getElementById('export-sheets').textContent = sheetCount;
        updateExportState();

        // 隐藏明细
        document.getElementById('summary-detail').style.display = 'none';

        goToStep(3);
        setStatus('数据汇总完成，可勾选设备后导出');
    } catch (e) {
        toast(`加载汇总失败: ${e.message}`, 'error');
    }
}

// 根据勾选情况更新导出相关按钮状态
function updateExportState() {
    const checked = document.querySelectorAll('.point-check:checked').length;
    const hasData = checked > 0;
    document.getElementById('btn-export').disabled = !hasData;
    document.getElementById('btn-to-export').disabled = !hasData;
    const hint = document.getElementById('export-hint');
    if (hint) {
        hint.textContent = hasData
            ? `已选择 ${checked} 台设备用于导出`
            : '尚未选择设备，请勾选有数据的设备';
    }
}

// 查看某测点的数据明细
async function viewPointDetail(pointNumber) {
    const detail = document.getElementById('summary-detail');
    const table = document.getElementById('summary-detail-table');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');

    document.getElementById('summary-detail-title').textContent = `测点 #${pointNumber} 数据明细（最多前200条）`;
    detail.style.display = '';

    thead.innerHTML = '<th>序号</th><th>日期</th><th>时间</th><th>温度(°C)</th><th>湿度(%)</th><th>报警</th>';
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text-muted)">加载中...</td></tr>';

    try {
        const records = await api(`/api/sessions/${state.sessionId}/points/${pointNumber}/data`);
        if (!records || records.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text-muted)">暂无数据</td></tr>';
            return;
        }
        tbody.innerHTML = records.map(r => `
            <tr>
                <td>${r.record_index || ''}</td>
                <td>${r.date_val || ''}</td>
                <td>${r.time_val || ''}</td>
                <td>${r.temperature ?? '-'}</td>
                <td>${r.humidity ?? '-'}</td>
                <td>${r.alarm || ''}</td>
            </tr>
        `).join('');
    } catch (e) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--danger)">加载失败: ${e.message}</td></tr>`;
    }
}

// ─── Step 4: 导出 ─────────────────────────────────────────────────────────
async function exportExcel() {
    const btn = document.getElementById('btn-export');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading-spinner"></span> 正在生成...';
    setStatus('正在生成 Excel 文件...');

    try {
        // 收集勾选的测点编号
        const checks = document.querySelectorAll('.point-check:checked');
        const selected = Array.from(checks).map(cb => cb.dataset.idx);
        const res = await api(`/api/sessions/${state.sessionId}/export`, {
            method: 'POST',
            body: JSON.stringify({ selected_indexes: selected }),
            headers: { 'Content-Type': 'application/json' }
        });
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Testo184_汇总_${new Date().toISOString().slice(0, 10)}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);

        btn.innerHTML = '📥 导出 Excel 文件';
        btn.disabled = false;
        document.getElementById('export-success').style.display = '';
        toast('Excel 文件已下载！', 'success');
        setStatus('导出完成');

        setTimeout(() => {
            document.getElementById('export-success').style.display = 'none';
        }, 5000);
    } catch (e) {
        btn.innerHTML = '📥 导出 Excel 文件';
        btn.disabled = false;
        toast(`导出失败: ${e.message}`, 'error');
        setStatus('导出失败');
    }
}

function newSession() {
    if (confirm('确定要新建会话吗？当前数据将保留在历史记录中。')) {
        state.sessionId = null;
        state.points = [];
        state.currentPointIndex = 0;
        state.completedDevices = [];
        state.totalRecords = 0;
        state.detectedDevice = null;
        document.getElementById('export-success').style.display = 'none';
        document.getElementById('session-name').value = '';
        document.getElementById('session-info').textContent = '';
        generatePoints();
        goToStep(1);
        setStatus('就绪 - 请配置测点后开始');
    }
}

// ─── 历史会话 ──────────────────────────────────────────────────────────────
async function showHistory() {
    document.getElementById('modal-history').style.display = '';
    const list = document.getElementById('history-list');
    list.innerHTML = '<p style="color:var(--text-muted);text-align:center">加载中...</p>';

    try {
        const sessions = await api('/api/sessions');
        if (sessions.length === 0) {
            list.innerHTML = '<p style="color:var(--text-muted);text-align:center">暂无历史会话</p>';
            return;
        }
        list.innerHTML = sessions.map(s => `
            <div class="history-item">
                <div class="history-item-info">
                    <h4>${s.name}</h4>
                    <p>${s.completed_points}/${s.total_points} 设备 · ${s.total_records} 条数据 · ${s.created_at || ''}</p>
                </div>
                <div class="history-item-actions">
                    <button class="btn btn-ghost btn-sm" onclick="loadSession('${s.id}')">查看</button>
                    <button class="btn btn-danger btn-sm" onclick="deleteSession('${s.id}')">删除</button>
                </div>
            </div>
        `).join('');
    } catch (e) {
        list.innerHTML = `<p style="color:var(--danger)">加载失败: ${e.message}</p>`;
    }
}

async function loadSession(sessionId) {
    try {
        const session = await api(`/api/sessions/${sessionId}`);
        state.sessionId = session.id;
        state.sessionName = session.name;
        state.points = session.points;
        state.totalRecords = session.total_records;
        state.currentPointIndex = session.completed_points;
        state.completedDevices = session.points
            .filter(p => p.status === 'completed')
            .map(p => ({ point_number: p.point_number, serial_number: p.serial_number, record_count: p.record_count }));

        document.getElementById('session-info').textContent = `会话: ${session.name}`;

        if (session.completed_points >= session.total_points) {
            buildSummary();
        } else {
            updateReadUI();
            goToStep(2);
        }

        document.getElementById('modal-history').style.display = 'none';
        toast(`已加载会话: ${session.name}`, 'info');
    } catch (e) {
        toast(`加载失败: ${e.message}`, 'error');
    }
}

async function deleteSession(sessionId) {
    if (!confirm('确定要删除此会话吗？所有数据将被清除。')) return;
    try {
        await api(`/api/sessions/${sessionId}`, { method: 'DELETE' });
        toast('会话已删除', 'success');
        showHistory(); // 刷新列表
    } catch (e) {
        toast(`删除失败: ${e.message}`, 'error');
    }
}

// ─── 事件绑定 ──────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    // Step 1
    document.getElementById('btn-gen-points').addEventListener('click', generatePoints);
    document.getElementById('btn-start-reading').addEventListener('click', startReading);

    // Step 2
    document.getElementById('btn-scan').addEventListener('click', scanDevice);
    document.getElementById('btn-confirm-read').addEventListener('click', confirmRead);
    document.getElementById('btn-import-vi2').addEventListener('click', importVi2);

    // Step 4
    document.getElementById('btn-export').addEventListener('click', exportExcel);
    document.getElementById('btn-new-session').addEventListener('click', newSession);

    // Step 3 → Step 4 衔接
    document.getElementById('btn-to-export').addEventListener('click', () => {
        // 勾选数量
        const checked = document.querySelectorAll('.point-check:checked').length;
        if (checked === 0) {
            toast('请先勾选要导出的设备（有数据的测点会自动勾选）', 'warning');
            return;
        }
        const hint = document.getElementById('export-hint');
        if (hint) hint.textContent = `将导出所勾选的 ${checked} 台设备`;
        goToStep(4);
        setStatus('请点击导出 Excel 文件');
    });
    document.getElementById('btn-close-detail').addEventListener('click', () => {
        document.getElementById('summary-detail').style.display = 'none';
    });

    // History
    document.getElementById('btn-history').addEventListener('click', showHistory);
    document.getElementById('btn-close-history').addEventListener('click', () => {
        document.getElementById('modal-history').style.display = 'none';
    });

    // 初始化测点
    generatePoints();
    setStatus('就绪 - 请配置测点后开始');
});
