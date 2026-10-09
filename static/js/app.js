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
    if (/\/sessions\/[^/]+\/export$/.test(url)) return res; // 旧版 session 文件流下载，不解析 JSON
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

// ─── v3.2.0 统一模式导航：read=读设备 / vi2=导入vi2 / export=分析导出 ───
function switchTo(mode) {
    state.currentMode = mode;
    const map = { read: 'panel-read', vi2: 'panel-vi2', export: 'panel-export' };
    document.querySelectorAll('.step-item').forEach(el => {
        el.classList.remove('active', 'done');
        if (el.getAttribute('data-mode') === mode) el.classList.add('active');
    });
    document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
    const pid = map[mode];
    if (pid && document.getElementById(pid)) document.getElementById(pid).classList.add('active');
    if (mode === 'export') refreshAll();
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
            // 文件处理明细（scan_log：每个文件怎么处理的、结果如何）
            const filesEl = document.getElementById('dev-files');
            const log = dev.scan_log || [];
            if (log.length) {
                filesEl.innerHTML = log.map(l => {
                    let color = 'var(--text-muted)', mark = '—';
                    if (l.status === 'ok') { color = 'var(--success)'; mark = '✓'; }
                    else if (l.status === 'error') { color = '#c0392b'; mark = '✗'; }
                    const cnt = l.status === 'ok' ? ` ${l.records} 条` : '';
                    return `<div style="color:${color};line-height:1.7">${mark} ${l.file} [${l.kind}]${cnt}${l.detail ? ' · ' + l.detail : ''}</div>`;
                }).join('');
            } else {
                filesEl.textContent = dev.csv_files.join(', ') || '-';
            }

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
                    '检测到设备，但未读取出温度数据。看左侧「文件处理明细」了解每个文件的情况；设备数据文件也可直接在下方「直接导入数据文件」处选择导入。';
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

// ─── Step 2: 导入 .vi2 文件（支持一次多选） ─────────────────────────────
async function importVi2() {
    const input = document.getElementById('vi2-file-input');
    const files = input.files;
    const msg = document.getElementById('vi2-msg');
    const btn = document.getElementById('btn-import-vi2');
    if (!files || files.length === 0) {
        toast('请先选择 .vi2 文件', 'warning');
        return;
    }
    if (!state.sessionId) {
        toast('请先在第一步创建会话', 'warning');
        return;
    }

    btn.disabled = true;
    btn.innerHTML = '<span class="loading-spinner"></span> 解析中...';
    if (msg) { msg.style.color = '#666'; msg.textContent = `正在解析 ${files.length} 个 .vi2 文件...`; }

    const form = new FormData();
    for (const f of files) form.append('file', f);

    try {
        const result = await api(`/api/sessions/${state.sessionId}/import-vi2`, {
            method: 'POST',
            body: form,
            isForm: true,
        });

        if (result.ok) {
            for (const r of (result.results || [])) {
                if (r.ok) {
                    state.completedDevices.push({
                        point_number: r.point_number,
                        serial_number: r.serial_number,
                        record_count: r.record_count,
                    });
                    state.totalRecords += r.record_count;
                    toast(r.message, 'success');
                } else {
                    toast(`${r.file}: ${r.error}`, 'error');
                }
            }
            if (msg) {
                msg.style.color = '#2e7d32';
                msg.textContent = '✅ ' + result.message;
            }
            setStatus(result.message);
            input.value = '';
            updateReadUI();
            loadSession();
        } else {
            if (msg) { msg.style.color = '#c62828'; msg.textContent = `❌ ${result.message || '导入失败'}`; }
            toast(result.message || '导入失败', 'error');
        }
    } catch (e) {
        if (msg) { msg.style.color = '#c62828'; msg.textContent = `导入失败: ${e.message}`; }
        toast(`导入失败: ${e.message}`, 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '📥 导入并读取数据';
    }
}

// ─── Step 2: Testo 软件联动（目录监控，新 .vi2/.csv 自动导入） ──────────────
let watchTimer = null;

async function prefillExportFolder() {
    const dirEl = document.getElementById('watch-dir');
    if (!dirEl || dirEl.value.trim()) return;
    try {
        const r = await api('/api/export-folder');
        if (r && r.path) dirEl.value = r.path;
    } catch (e) { /* 静默：用户可手动填写 */ }
}

async function openExportFolder() {
    try {
        const r = await api('/api/export-folder/open', { method: 'POST' });
        if (r && r.ok) {
            toast('已打开交接文件夹，在 ComSoft 里保存/导出数据时选择这个位置', 'success');
            const dirEl = document.getElementById('watch-dir');
            if (dirEl && !dirEl.value.trim() && r.path) dirEl.value = r.path;
        }
    } catch (e) {
        toast(`打开失败: ${e.message}`, 'error');
    }
}

async function detectComsoft() {
    const list = document.getElementById('comsoft-list');
    const btn = document.getElementById('btn-detect-comsoft');
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="loading-spinner"></span> 检测中...'; }
    try {
        const r = await api('/api/comsoft/detect');
        if (r.platform !== 'win32') {
            if (list) list.innerHTML = '<span style="color:#c62828">自动调用仅支持 Windows（ComSoft 官方软件只有 Windows 版）。Mac 请手动用虚拟机/其他 Windows 电脑读取，把 .vi2 或 CSV 存到监控文件夹，本工具会自动导入。</span>';
            return;
        }
        if (!r.softwares || r.softwares.length === 0) {
            if (list) list.innerHTML = '<span style="color:#c62828">未在电脑上找到 ComSoft 软件。请确认 ComSoft Professional（专业版，支持 testo 184 全量读取）已安装，或手动打开。读取后把 .vi2 / CSV 保存到监控文件夹即可。</span>';
            return;
        }
        const items = r.softwares.map(sw =>
            `<div style="display:flex;align-items:center;gap:10px;padding:6px 0;flex-wrap:wrap">
                <span style="color:#2e7d32">✔ ${sw.name}</span>
                ${sw.exe ? `<button class="btn btn-sm" style="padding:4px 12px" onclick="launchComsoft('${sw.exe.replace(/\\/g, '\\\\')}')">▶ 启动</button>` : ''}
                <span style="color:#999;font-size:12px">${sw.location || ''}</span>
            </div>`).join('');
        if (list) list.innerHTML = '<b>检测到以下 Testo 官方软件：</b><br>' + items;
        // 同时启动第一个
        if (r.softwares[0].exe) await launchComsoft(r.softwares[0].exe);
    } catch (e) {
        if (list) list.innerHTML = `<span style="color:#c62828">检测失败: ${e.message}</span>`;
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = '🔍 一键调用官方软件'; }
    }
}

async function launchComsoft(exe) {
    try {
        const r = await api('/api/comsoft/launch', { method: 'POST', body: { exe } });
        if (r.ok) {
            toast('已启动官方软件：连接并读取温度计后，把数据保存为 .vi2 或导出 CSV 到监控文件夹，本工具会自动导入', 'success');
        }
    } catch (e) {
        toast(`启动失败: ${e.message}`, 'error');
    }
}

async function startWatch() {
    const dirEl = document.getElementById('watch-dir');
    const dir = dirEl ? dirEl.value.trim() : '';
    const st = document.getElementById('watch-status');
    if (!dir) { toast('请先填写 Testo 软件保存 .vi2 的文件夹路径', 'warning'); return; }
    if (!state.sessionId) { toast('请先在第一步创建会话', 'warning'); return; }
    const btn = document.getElementById('btn-watch');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading-spinner"></span> 启动中...';
    try {
        const r = await api(`/api/sessions/${state.sessionId}/watch-folder`, {
            method: 'POST', body: { path: dir }
        });
        if (st) {
            st.style.color = '#2e7d32';
            st.textContent = `🔄 正在监控 ${dir} —— 在 Testo 软件里读取设备并保存 .vi2 到该文件夹，数据会自动导入`;
        }
        toast('已开始监控文件夹', 'success');
        btn.innerHTML = '🔄 监控中...';
        if (watchTimer) clearInterval(watchTimer);
        watchTimer = setInterval(pollWatchStatus, 3000);
    } catch (e) {
        if (st) { st.style.color = '#c62828'; st.textContent = `启动监控失败: ${e.message}`; }
        toast(e.message || '启动监控失败', 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '🔄 开始监控';
    }
}

async function pollWatchStatus() {
    if (!state.sessionId) return;
    try {
        const r = await api(`/api/sessions/${state.sessionId}/watch-status`);
        const diag = document.getElementById('watch-diag');
        if (diag) {
            if (r.new_files_seen && r.new_files_seen.length) {
                diag.innerHTML = r.new_files_seen.map(f =>
                    `⚠️ 发现新文件 <b>${f.file}</b>（${(f.size/1024).toFixed(1)} KB）—— ${f.note}`
                ).join('<br>');
            }
        }
        if (r.new_imports && r.new_imports.length) {
            for (const im of r.new_imports) {
                if (im.ok) {
                    state.completedDevices.push({
                        point_number: im.point_number,
                        serial_number: im.serial_number,
                        record_count: im.record_count,
                    });
                    state.totalRecords += im.record_count;
                    toast(`自动导入 ${im.file}: ${im.record_count} 条 (SN ${im.serial_number})`, 'success');
                } else {
                    toast(`自动导入失败 ${im.file}: ${im.error}`, 'error');
                }
            }
            loadSession();
        }
    } catch (e) { /* 轮询静默 */ }
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
    const $ = (id) => document.getElementById(id);

    // === 统一模式导航（read / vi2 / export）===
    document.querySelectorAll('.step-item').forEach(el => {
        el.addEventListener('click', (e) => {
            e.preventDefault();
            const mode = el.getAttribute('data-mode');
            if (mode) switchTo(mode);
        });
    });

    // === 方式一：读取温度计设备 ===
    // ⚙ 设置 Comfort Software 路径：展开面板，支持"自动检测"或"手动输入完整路径"
    $('btn-cc4-setup').addEventListener('click', () => {
        const panel = document.getElementById('cc4-panel');
        panel.style.display = (panel.style.display === 'none') ? 'block' : 'none';
        cc4RefreshPanel();
    });
    $('btn-cc4-autodetect').addEventListener('click', cc4AutoDetect);
    $('btn-cc4-save').addEventListener('click', cc4SaveInput);
    document.getElementById('cc4-path-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') cc4SaveInput(); });
    $('btn-batch-start').addEventListener('click', batchStart);
    $('btn-batch-detect').addEventListener('click', batchDetect);
    $('btn-batch-clear').addEventListener('click', batchClear);

    // === 方式二：批量上传 vi2 ===
    const vdz = $('vi2-dropzone');
    const vfile = $('vi2-file-input');
    vdz.addEventListener('click', () => vfile.click());
    vfile.addEventListener('change', (e) => { vi2UploadFiles(e.target.files); e.target.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => vdz.addEventListener(ev, (e) => {
        e.preventDefault(); vdz.style.borderColor = '#27ae60'; vdz.style.background = '#eef9f0';
    }));
    ['dragleave', 'drop'].forEach(ev => vdz.addEventListener(ev, (e) => {
        e.preventDefault(); vdz.style.borderColor = '#aab8c8'; vdz.style.background = '#f7fafc';
    }));
    vdz.addEventListener('drop', (e) => { if (e.dataTransfer && e.dataTransfer.files) vi2UploadFiles(e.dataTransfer.files); });

    // === 统一数据池：分析并导出 ===
    $('btn-all-export').addEventListener('click', allExport);
    $('btn-all-refresh').addEventListener('click', refreshAll);
    $('btn-all-clear').addEventListener('click', allClear);

    // === 各方式「进入下一环节：分析导出」 ===
    if ($('btn-read-next')) $('btn-read-next').addEventListener('click', () => switchTo('export'));
    if ($('btn-vi2-next')) $('btn-vi2-next').addEventListener('click', () => switchTo('export'));

    // 初始化：显示方式一，刷新统一列表
    switchTo('read');
    refreshAll();
    setStatus('就绪 - 选择上方任一方式读取数据，再到「分析并导出Excel」');
});

/* ===== v3.0.0 多温度计批量插拔读取 ===== */
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function escJs(s) { return String(s == null ? '' : s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

async function detectCc4() {
    const info = document.getElementById('batch-cc4-info');
    try {
        const r = await fetch('/api/comsoft/detect').then(x => x.json());
        const saved = r.saved || '';
        const sw = (r.softwares || []).filter(s => s.exe || s.location);
        if (saved) {
            info.innerHTML = '💾 已记住 cc4 路径：<b>' + esc(saved) + '</b>' +
                (sw.length ? '<br>检测到其他候选如下，可点选切换：' : '');
        } else {
            info.innerHTML = '';
        }
        if (!sw.length) {
            if (!saved) info.innerHTML += '⚠️ 未检测到 cc4.exe / Comfort Software。<br>可点击「📁 选择 cc4.exe」手动指定文件，或直接在电脑上打开软件读取设备。';
        } else {
            const rows = sw.map(s => {
                const path = s.exe || (s.location || '');
                const tag = path === saved ? '（当前）' : '';
                return '<div style="margin:3px 0"><a href="javascript:void(0)" onclick="saveCc4Path(\'' + escJs(path) + '\')" style="text-decoration:underline">✔ 选用</a> ' +
                    esc(s.name || (path.split(/[\\/]/).pop() || '软件')) + tag + '<br><span style="color:#888;font-size:12px">' + esc(path) + '</span></div>';
            }).join('');
            info.innerHTML += '✅ 检测到 ' + sw.length + ' 个候选：<br>' + rows;
        }
    } catch (e) { info.textContent = '检测失败：' + e.message; }
}
async function saveCc4Path(p) {
    const info = document.getElementById('batch-cc4-info');
    try {
        const r = await fetch('/api/comsoft/set-path', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ exe: p }) }).then(x => x.json());
        if (r.ok) { info.innerHTML = '✅ 已选用 cc4：<b>' + esc(r.saved) + '</b>'; }
        else info.textContent = r.error || '保存失败';
    } catch (e) { info.textContent = '保存失败：' + e.message; }
}
async function launchCc4() {
    const info = document.getElementById('batch-cc4-info');
    try {
        const r = await fetch('/api/comsoft/launch', { method: 'POST', headers: {'Content-Type':'application/json'}, body: '{}' }).then(x => x.json());
        if (r.ok) info.innerHTML = '✅ 已启动：' + esc(r.launched || 'Comfort Software') + '<br>在软件里读取设备后，数据会自动被本工具识别并汇总。';
        else info.textContent = r.error || '启动失败';
    } catch (e) { info.textContent = '启动失败：' + e.message; }
}
async function cc4RefreshPanel() {
    const info = document.getElementById('batch-cc4-info');
    try {
        const r = await fetch('/api/comsoft/get-path').then(x => x.json());
        if (r && r.path) {
            document.getElementById('cc4-path-input').value = r.path;
            info.innerHTML = '当前已保存路径：<b>' + esc(r.path) + '</b>';
        } else {
            info.textContent = '尚未设置。请「自动检测」，或手动输入 cc4.exe 的完整路径后点「保存」。';
        }
    } catch (e) { info.textContent = '读取已保存路径失败：' + e.message; }
}

async function cc4AutoDetect() {
    const info = document.getElementById('batch-cc4-info');
    const input = document.getElementById('cc4-path-input');
    info.innerHTML = '🔎 正在自动扫描已安装的 Testo/ComSoft…';
    try {
        const r = await fetch('/api/comsoft/detect').then(x => x.json());
        let found = null;
        const softs = r.softwares || [];
        // 优先 cc4.exe（ComSoft），其次 testo 相关 exe
        found = softs.find(s => s.exe && /cc4\.exe$/i.test(s.exe))
             || softs.find(s => s.exe && /(comsoft|testo)/i.test(s.exe))
             || softs[0];
        if (found && found.exe) {
            input.value = found.exe;
            // 自动保存
            const res = await fetch('/api/comsoft/set-path', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ exe: found.exe }) }).then(x => x.json());
            if (res.ok) info.innerHTML = '✅ 已自动检测并保存：<b>' + esc(res.saved) + '</b>';
            else info.textContent = '检测到但保存失败：' + (res.error || '');
        } else {
            info.innerHTML = '未自动检测到 cc4.exe。请手动输入完整路径后点「保存」。<br><span style="font-size:12px;color:#888">默认位置通常是 <code>D:\\Testo\\Comfort Software\\cc4.exe</code> 或 <code>C:\\Program Files (x86)\\Testo\\Comfort Software\\cc4.exe</code></span>';
        }
    } catch (e) { info.textContent = '自动检测失败：' + e.message; }
}

async function cc4SaveInput() {
    const info = document.getElementById('batch-cc4-info');
    const p = document.getElementById('cc4-path-input').value.trim();
    if (!p) { info.textContent = '请输入 cc4.exe 的完整路径，或点「🔎 自动检测」。'; return; }
    // 用户常粘贴带引号的路径，去掉
    const clean = p.replace(/^"+|"+$/g, '');
    info.innerHTML = '保存中…';
    try {
        const r = await fetch('/api/comsoft/set-path', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ exe: clean }) }).then(x => x.json());
        if (r.ok) info.innerHTML = '✅ 已保存 Comfort Software 路径：<b>' + esc(r.saved) + '</b>';
        else info.innerHTML = '❌ 保存失败：' + esc(r.error || '') + '<br><span style="font-size:12px;color:#888">请确认路径正确（包含文件名 cc4.exe 且文件确实存在），例如 <code>D:\\Testo\\Comfort Software\\cc4.exe</code></span>';
    } catch (e) { info.textContent = '保存失败：' + e.message; }
}

async function batchStart() {
    document.getElementById('batch-wizard').style.display = 'block';
    document.getElementById('btn-batch-start').disabled = true;
    document.getElementById('btn-batch-clear').style.display = '';
    document.getElementById('batch-current').innerHTML = '';
    // 先探测 ComSoft 路径是否已设置
    let cc4 = '';
    try {
        const p = await fetch('/api/comsoft/get-path').then(x => x.json());
        if (p && p.path) cc4 = p.path;
    } catch (e) { /* 忽略 */ }
    if (!cc4) {
        setBatchStepHint('⚠️ 请先点击上方「⚙ 设置 Comfort Software 路径」，选择你的 cc4.exe（通常在 D:\\Testo\\Comfort Software\\cc4.exe），之后再点开始。');
        document.getElementById('btn-batch-start').disabled = false;
        return;
    }
    try {
        const r = await fetch('/api/collector/start', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({}) }).then(x => x.json());
        if (!r.ok) {
            setBatchStepHint('❌ 采集启动失败：' + esc(r.error || '未知错误'));
            document.getElementById('btn-batch-start').disabled = false;
            return;
        }
    } catch (e) {
        setBatchStepHint('❌ 采集启动失败：' + esc(e.message));
        document.getElementById('btn-batch-start').disabled = false;
        return;
    }
    setBatchStepHint('🤖 采集已启动 —— 现在请<b>依次插拔每一台温度计</b>，工具会自动读取原始数据并收集。<br>插完所有温度计后，点下方「进入下一环节」即可分析导出。');
    // 显示采集运行状态 + 自动刷新已读入列表（每次轮询都自动扫描接力文件夹新增 vi2）
    ensureCollectorPoll();
    batchAutoRefreshTimer = setInterval(() => {
        refreshBatchList();
        pollCollectorStatus();
    }, 4000);
}
let batchAutoRefreshTimer = null;
async function batchDetect() {
    const box = document.getElementById('batch-current');
    box.innerHTML = '检测中…';
    const r = await fetch('/api/batch/autoscan', { method: 'POST' }).then(x => x.json());
    if (r.ok && r.count) {
        const added = (r.added && r.added.length) ? '（新增 <b>' + r.added.length + '</b> 台）' : '（本次无新增，共已读入 <b>' + r.count + '</b> 台）';
        box.innerHTML = '<div style="color:#1b7f3b">✅ 已自动读入：共 <b>' + r.count + '</b> 台温度计 ' + added +
            '<br><span style="font-size:12px;color:#666">继续插拔下一台，工具会自动收集。</span></div>';
    } else if (r.error) {
        box.innerHTML = '<div style="color:#c0392b">检测失败：' + esc(r.error) + '</div>';
    } else {
        box.innerHTML = '<div style="color:#c0392b">未检测到新增设备：' +
            '<br>请确认温度计已插好或采集器已完成导出，稍后再点检测。</div>';
    }
    await refreshBatchList();
}
async function refreshBatchList() {
    // 一键全自动：先自动扫描接力文件夹/已插设备，把新增测点读入批次（幂等去重）
    try {
        await fetch('/api/batch/autoscan', { method: 'POST' }).then(x => x.json());
    } catch (e) { /* 忽略 */ }
    const r = await fetch('/api/batch/status').then(x => x.json());
    setBatchList(r.devices || []);
    const n = (r.devices || []).length;
    document.getElementById('read-point-count').textContent = n;
    if (n > 0) {
        document.getElementById('read-next-bar').style.display = '';
        document.getElementById('btn-batch-clear').style.display = '';
    }
}
async function batchExport() {
    const box = document.getElementById('batch-current');
    box.innerHTML = '正在生成汇总 Excel…';
    const r = await fetch('/api/batch/export', { method: 'POST' }).then(x => x.json());
    if (!r.ok) { box.innerHTML = '<div style="color:#c0392b">导出失败：' + esc(r.error || '') + '</div>'; return; }
    const a = '/api/batch/download/' + encodeURIComponent(r.file);
    box.innerHTML = '<div style="color:#1e7e34;line-height:1.9">✅ 已生成汇总 Excel（<b>' + r.device_count + '</b> 台设备）。<br>' +
        '<a class="btn btn-primary" href="' + a + '" download="Testo184_Batch.xlsx">📥 下载 汇总 Excel</a></div>';
}
async function batchClear() {
    await fetch('/api/batch/clear', { method: 'POST' });
    document.getElementById('batch-wizard').style.display = 'none';
    document.getElementById('batch-list').innerHTML = '';
    document.getElementById('batch-current').innerHTML = '';
    document.getElementById('btn-batch-start').disabled = false;
}
function setBatchStepHint(html) { document.getElementById('batch-step-hint').innerHTML = html; }

// ─── 采集引擎状态（自动插拔读取）────────────────────────────────────────
let collectorPollTimer = null;
function ensureCollectorPoll() {
    if (collectorPollTimer) return;
    collectorPollTimer = setInterval(pollCollectorStatus, 3000);
}
async function pollCollectorStatus() {
    try {
        const r = await fetch('/api/collector/status').then(x => x.json());
        // 采集专用状态面板：展示运行状态 + 最近日志
        const panel = document.getElementById('batch-collector-state');
        if (panel) {
            const st = r.running
                ? '<span style="color:#1b7f3b">● 采集引擎运行中</span>（插拔温度计即自动读取导出）'
                : '<span style="color:#c0392b">● 采集引擎未运行</span>';
            const logs = (r.log || []).slice(-6);
            let logHtml = '';
            if (logs.length) {
                logHtml = '<div style="margin-top:6px;font-family:monospace;font-size:12px;line-height:1.6;color:#555;background:#fbfcfd;border:1px solid #eef1f4;border-radius:5px;padding:8px 10px;max-height:130px;overflow:auto">' +
                    logs.map(x => esc(x)).join('<br>') + '</div>';
            }
            panel.style.display = 'block';
            panel.innerHTML = '<div style="font-size:13px">' + st + '</div>' + logHtml;
        }
        // next-bar 也显示状态
        const bar = document.getElementById('read-next-bar');
        if (bar) {
            const inner = bar.querySelector('.next-bar-inner');
            if (inner) {
                const st = r.running ? `<span style="color:#1b7f3b">● 采集引擎运行中</span>` : `<span style="color:#c0392b">● 采集引擎未运行</span>`;
                const last = (r.log || []).slice(-1)[0] || '';
                inner.innerHTML = st + (last ? `<br><span style="font-size:12px;color:#666">${esc(last)}</span>` : '');
            }
        }
    } catch (e) { /* 忽略轮询错误 */ }
}

// ─── v3.1.0 批量上传 .vi2 → 一键导出 Excel ────────────────────────────────
async function vi2UploadFiles(fileList) {
    if (!fileList || !fileList.length) return;
    const dz = document.getElementById('vi2-dropzone');
    const prog = document.getElementById('vi2-progress');
    const bar = document.getElementById('vi2-progress-bar');
    const txt = document.getElementById('vi2-progress-text');
    prog.style.display = 'block';
    txt.textContent = `正在上传解析 ${fileList.length} 个文件…(0/1)`;
    bar.style.width = '0%';

    const form = new FormData();
    for (const f of fileList) {
        if (f.name && f.name.toLowerCase().endsWith('.vi2')) form.append('file', f);
    }
    if (![...form.keys()].length) {
        prog.style.display = 'none';
        toast('请选择 .vi2 文件', 'error');
        return;
    }
    try {
        const r = await api('/api/vi2/upload', { method: 'POST', body: form, isForm: true });
        const res = r;
        bar.style.width = '100%';
        txt.textContent = '解析完成。';
        if (res.added && res.added.length) {
            toast(`已成功解析 ${res.added.length} 个文件`, 'success');
            // 上传成功 → 立即跳转统一导出页并刷新
            switchTo('export');
        } else {
            toast('没有新设备被加入', 'error');
        }
        if (res.errors && res.errors.length) {
            const msg = document.getElementById('vi2-msg');
            if (msg) msg.innerHTML = '⚠️ ' + res.errors.length + ' 个文件未加入：' + res.errors.slice(0, 3).map(e => esc(e.file) + ' - ' + esc(e.error)).join('；');
        }
    } catch (e) {
        prog.style.display = 'none';
        toast('上传失败: ' + e.message, 'error');
    }
}

/* ===== v3.2.0 统一数据池 refreshAll（合并 设备读取 + vi2导入） ===== */
async function refreshAll() {
    try {
        const res = await api('/api/all/list', { method: 'GET' });
        const devices = res.devices || [];
        updateNextBars(devices);
        const box = document.getElementById('all-list');
        const count = document.getElementById('exp-count');
        const rec = document.getElementById('exp-records');
        const src = document.getElementById('exp-source');
        if (count) count.textContent = devices.length;
        let total = 0; const srcs = new Set();
        devices.forEach(d => { total += (d.records || d.record_count || 0); if (d.source) srcs.add(d.source); });
        if (rec) rec.textContent = total;
        if (src) src.textContent = srcs.size ? [...srcs].join(' / ') : '—';
        if (!devices.length) {
            box.innerHTML = '<div style="padding:20px;background:#f7f9fb;border-radius:8px;color:#666;text-align:center">暂无数据。<br>请先通过 <b>方式一读取设备</b> 或 <b>方式二导入 .vi2</b> 添加数据，再回到这里分析导出。</div>';
            return;
        }
        box.innerHTML = '<div style="font-size:13px;color:#666;margin-bottom:6px">已收集设备：</div>' +
            devices.map((d, i) =>
                '<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;margin:5px 0;background:#f3f7fb;border-radius:8px;font-size:13px;flex-wrap:wrap;gap:6px">' +
                '<span><b>#' + (i + 1) + '</b>　SN <b>' + esc(d.serial_number || d.sn || '') + '</b>　—　' + (d.records || d.record_count || 0) + ' 条</span>' +
                '<span style="color:#888;font-size:12px">' + esc(d.source || '') + (d.file ? ' · ' + esc(d.file) : '') + '</span>' +
                '</div>').join('');
    } catch (e) {
        const box = document.getElementById('all-list');
        if (box) box.innerHTML = '<span style="color:#c0392b">刷新失败: ' + esc(e.message) + '</span>';
    }
}

/* ===== v3.2.0 更新「已形成测点数据」提示条 ===== */
function updateNextBars(devices) {
    const total = (devices || []).length;
    // 方式一：读取设备
    const rb = document.getElementById('read-next-bar');
    if (rb) {
        const n = document.getElementById('read-point-count');
        if (n) n.textContent = total;
        rb.style.display = total > 0 ? '' : 'none';
    }
    // 方式二：导入 vi2
    const vb = document.getElementById('vi2-next-bar');
    if (vb) {
        const n = document.getElementById('vi2-point-count');
        if (n) n.textContent = total;
        vb.style.display = total > 0 ? '' : 'none';
    }
}

/* ===== v3.2.0 统一导出 allExport ===== */
async function allExport() {
    const box = document.getElementById('all-export-note');
    const btn = document.getElementById('btn-all-export');
    if (!btn) return;
    btn.disabled = true;
    btn.textContent = '⏳ 正在生成 Excel…';
    try {
        const res = await api('/api/all/export', { method: 'POST', body: {} });
        if (!res.ok) { box.innerHTML = '<span style="color:#c0392b">导出失败：' + esc(res.error || '') + '</span>'; return; }
        box.innerHTML = '✅ 已生成 <b>' + res.count + '</b> 台设备、共 <b>' + res.records + '</b> 条数据的汇总包：<br><br>' +
            '<a class="btn btn-primary btn-lg" href="' + res.url + '" download="' + esc(res.filename || 'Testo184_汇总.zip') + '" style="display:inline-block">📥 下载汇总包 (ZIP)</a>' +
            '<div style="font-size:12px;color:#666;margin-top:6px">内含：每个测点一个「测点N-日期.xlsx」+ 一个「汇总-演示-日期.xlsx」。<br>文件名：' + esc(res.filename || 'Testo184_汇总.zip') + '</div>';
    } catch (e) {
        box.innerHTML = '<span style="color:#c0392b">导出失败: ' + esc(e.message) + '</span>';
    }
    btn.disabled = false;
    btn.textContent = '📥 分析并导出 Excel';
}

/* ===== v3.2.0 统一清空 allClear ===== */
async function allClear() {
    if (!confirm('确定清空全部已读取/已导入的数据吗？')) return;
    try {
        await api('/api/all/clear', { method: 'POST', body: {} });
        document.getElementById('batch-wizard').style.display = 'none';
        document.getElementById('batch-list').innerHTML = '';
        document.getElementById('batch-current').innerHTML = '';
        document.getElementById('btn-batch-start').disabled = false;
        document.getElementById('vi2-progress').style.display = 'none';
        document.getElementById('vi2-msg').innerHTML = '';
        document.getElementById('all-export-note').innerHTML = '';
        await refreshAll();
        toast('已清空全部数据', 'success');
    } catch (e) {
        toast('清空失败: ' + e.message, 'error');
    }
}

function setBatchList(devices) {
    const el = document.getElementById('batch-list');
    if (!devices || !devices.length) { el.innerHTML = '<div style="color:#666;font-size:13px">已读取设备列表为空</div>'; return; }
    el.innerHTML = '<div style="font-size:13px;color:#666;margin-bottom:6px">已读取设备：</div>' +
        devices.map((d, i) =>
            '<div style="padding:6px 10px;margin:4px 0;background:#f3f7fb;border-radius:6px;font-size:13px">' +
            (i + 1) + '. <b>SN ' + esc(d.sn) + '</b> — ' + d.record_count + ' 条（' + esc(d.source || '') + '）</div>').join('');
}
