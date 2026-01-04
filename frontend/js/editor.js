/**
 * 共享表格 编辑器 - 核心协作逻辑
 */

// 全局变量
let spreadsheet = null;
let websocket = null;
let sheetKey = null;
let currentUserId = null;
let onlineUsers = [];
let isUpdatingFromRemote = false;  // 防止循环更新
let historyCount = 20;  // 历史记录显示条数
let syncInterval = 3000;  // 定时同步间隔（毫秒）
let pendingCellUpdate = null;  // 待发送的单元格更新（退出编辑状态时发送）
let lastSelectedCell = null;  // 上次选中的单元格位置
let isEditTextareaUpdating = false;  // 防止编辑窗口循环更新
let currentEditCell = { row: -1, col: -1 };  // 当前编辑的单元格位置
let lastDimensions = { cols: {}, rows: {} };  // 上次记录的列宽行高

// 用户颜色映射
const userColors = [
    '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4',
    '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F'
];

function getUserColor(userId) {
    const hash = userId.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    return userColors[hash % userColors.length];
}

// 获取用户标识（浏览器限制，使用指纹+localStorage持久化）
async function getMacAddress() {
    // 先检查localStorage中是否已有标识
    let storedMac = localStorage.getItem('共享表格_user_id');
    if (storedMac) {
        return storedMac;
    }

    // 浏览器无法直接获取MAC地址，使用指纹作为替代
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    ctx.textBaseline = 'top';
    ctx.font = '14px Arial';
    ctx.fillText('共享表格', 2, 2);
    const fingerprint = canvas.toDataURL().slice(-50);

    // 生成类似MAC的标识
    let hash = 0;
    for (let i = 0; i < fingerprint.length; i++) {
        hash = ((hash << 5) - hash) + fingerprint.charCodeAt(i);
        hash = hash & hash;
    }

    const mac = Math.abs(hash).toString(16).padStart(12, '0').slice(0, 12);
    storedMac = mac.match(/.{2}/g).join(':').toUpperCase();

    // 保存到localStorage
    localStorage.setItem('共享表格_user_id', storedMac);
    return storedMac;
}

// 初始化
window.addEventListener('load', async function() {
    // 从URL路径获取密钥（/sheet/{key}）
    const pathMatch = window.location.pathname.match(/^\/sheet\/([^\/]+)$/);

    if (pathMatch) {
        sheetKey = decodeURIComponent(pathMatch[1]);

        // 验证密钥并获取表格信息
        try {
            const response = await fetch(`/api/auth?key=${encodeURIComponent(sheetKey)}`);
            const result = await response.json();

            if (result.success) {
                sessionStorage.setItem('sheet_key', sheetKey);
                sessionStorage.setItem('sheet_name', result.sheet_name);
            } else {
                showToast('密钥无效或表格不存在', 'error');
                return;
            }
        } catch (error) {
            console.error('验证密钥失败:', error);
            showToast('无法连接服务器', 'error');
            return;
        }
    } else {
        // 尝试从sessionStorage获取（兼容旧方式）
        sheetKey = sessionStorage.getItem('sheet_key');
        if (!sheetKey) {
            showToast('无效的访问链接', 'error');
            return;
        }
    }

    const sheetName = sessionStorage.getItem('sheet_name');
    document.getElementById('sheetName').textContent = sheetName || '未命名表格';

    try {
        // 加载服务器配置
        const configResponse = await fetch('/api/config');
        const config = await configResponse.json();
        historyCount = config.history_count || 20;
        syncInterval = config.sync_interval || 3000;

        // 加载表格数据
        const sheetData = await loadSheetData();

        // 初始化电子表格
        initSpreadsheet(sheetData);

        // 连接WebSocket
        await connectWebSocket();

        // 加载历史记录
        await loadHistory();

        // 初始化历史面板
        initHistoryPanel();

        // 初始化编辑窗口
        initEditPanel();

        // 初始化列宽行高跟踪
        initializeDimensionsTracking();

        // 默认选中A1单元格
        setTimeout(() => {
            if (spreadsheet) {
                // 触发选中A1单元格
                updateEditPanel(0, 0);
            }
        }, 500);

        // 启动定时同步
        if (syncInterval > 0) {
            startPeriodicSync();
        }

        showToast('表格加载成功', 'success');
    } catch (error) {
        console.error('初始化失败:', error);
        console.error('错误堆栈:', error.stack);
        console.error('错误名称:', error.name);
        console.error('错误消息:', error.message);
        showToast('加载失败: ' + error.message + ' (详细信息请查看控制台)', 'error');
    }

    // 绑定事件
    document.getElementById('exportBtn').addEventListener('click', exportSheet);
    document.getElementById('logoutBtn').addEventListener('click', logout);
});

// 加载表格数据
async function loadSheetData() {
    const response = await fetch(`/api/sheet/${sheetKey}`);
    if (!response.ok) {
        throw new Error('无法加载表格数据');
    }
    return await response.json();
}

// 初始化电子表格
function initSpreadsheet(sheetData) {
    try {
        console.log('初始化电子表格，数据:', sheetData);
        const container = document.getElementById('spreadsheet-container');

        // 安全检查
        if (!sheetData || typeof sheetData !== 'object') {
            sheetData = {};
        }

        // 转换数据格式为x-spreadsheet格式
        console.log('开始转换数据格式...');
        const xsData = convertToXsFormat(sheetData);
        console.log('转换后的数据:', xsData);

        // 创建电子表格实例
        console.log('创建电子表格实例...');
        spreadsheet = x_spreadsheet(container, {
            mode: 'edit',
            showToolbar: true,
            showGrid: true,
            showContextmenu: true,
            view: {
                height: () => container.clientHeight,
                width: () => container.clientWidth
            },
            row: {
                len: (sheetData && sheetData.rowCount) || 100,
                height: 25
            },
            col: {
                len: (sheetData && sheetData.columnCount) || 26,
                width: 100,
                indexWidth: 60,
                minWidth: 60
            }
        });

        // 加载数据
        console.log('加载数据到表格...');
        console.log('加载的数据结构:', JSON.stringify(xsData, null, 2));
        spreadsheet.loadData(xsData);
        console.log('数据加载完成');

        // 调试：检查加载后的单元格数据
        setTimeout(() => {
            console.log('===== 调试：检查单元格数据 =====');
            for (let ri = 0; ri <= 8; ri++) {  // 检查 0-8 行 (A1-F9)
                for (let ci = 0; ci <= 5; ci++) {  // 检查 0-5 列 (A-F)
                    const cell = spreadsheet.cell(ri, ci);
                    const colLetter = String.fromCharCode(65 + ci);
                    console.log(`${colLetter}${ri + 1}:`, cell);
                }
            }
            console.log('================================');
        }, 1000);

        // 监听单元格变化 - 只保存，不立即发送（退出编辑状态时才发送）
        spreadsheet.on('cell-edited', (text, ri, ci) => {
            if (!isUpdatingFromRemote) {
                // 只保存待发送的更新，不立即发送
                pendingCellUpdate = { row: ri, col: ci, value: text };

                // 同步到编辑窗口
                if (currentEditCell.row === ri && currentEditCell.col === ci && !isEditTextareaUpdating) {
                    isEditTextareaUpdating = true;
                    const textarea = document.getElementById('editTextarea');
                    if (textarea) {
                        textarea.value = text;
                    }
                    isEditTextareaUpdating = false;
                }
            }
        });

        // 监听选区变化 - 切换单元格时发送待处理的更新（退出编辑状态）
        spreadsheet.on('cell-selected', (cell, ri, ci) => {
            // 如果有待发送的更新，发送它（用户已退出编辑状态）
            if (pendingCellUpdate) {
                sendCellUpdate(pendingCellUpdate.row, pendingCellUpdate.col, pendingCellUpdate.value);
                pendingCellUpdate = null;
            }
            // 记录当前选中的单元格及其当前值（用于Delete检测）
            const cellData = spreadsheet.cell(ri, ci);
            lastSelectedCell = {
                row: ri,
                col: ci,
                originalValue: cellData ? (cellData.text || '') : ''
            };
            // 更新编辑窗口
            updateEditPanel(ri, ci);
            sendSelectionChange(ri, ci);
        });

        // 监听键盘事件 - 处理Delete键删除（在document上监听以确保捕获）
        document.addEventListener('keydown', (e) => {
            if ((e.key === 'Delete' || e.key === 'Backspace') && lastSelectedCell && !isUpdatingFromRemote) {
                // 检查焦点是否在输入框内（编辑模式），如果是则不处理
                const activeEl = document.activeElement;
                if (activeEl && (activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA' || activeEl.isContentEditable)) {
                    return;
                }

                // 如果选中的单元格原本有内容，发送清空更新
                if (lastSelectedCell.originalValue) {
                    // 延迟一点发送，确保x-spreadsheet已处理删除
                    setTimeout(() => {
                        sendCellUpdate(lastSelectedCell.row, lastSelectedCell.col, '');
                        // 清空原始值，避免重复发送
                        lastSelectedCell.originalValue = '';
                    }, 100);
                }
            }
        });

        console.log('电子表格初始化完成');
    } catch (error) {
        console.error('initSpreadsheet 错误:', error);
        console.error('错误堆栈:', error.stack);
        throw error;
    }
}

// 转换数据格式
function convertToXsFormat(sheetData) {
    try {
        console.log('开始转换数据格式...', sheetData);

        const rows = {};
        const styles = [];  // 样式表

        // 安全检查：确保sheetData和cellData存在
        if (!sheetData) {
            sheetData = {};
        }
        const cellData = sheetData.cellData || {};

        // 确保cellData是对象
        if (typeof cellData !== 'object' || cellData === null) {
            console.warn('cellData不是有效对象，使用空对象');
            return [{
                name: sheetData.name || 'Sheet1',
                rows: {},
                merges: [],
                cols: {},
                styles: []
            }];
        }

        // 先处理行高数据
        const rowData = sheetData.rowData || {};
        console.log('处理行高数据:', rowData);
        for (const [rowIdx, height] of Object.entries(rowData)) {
            const idx = parseInt(rowIdx);
            if (!rows[idx]) {
                rows[idx] = { cells: {}, height: height };
            }
        }

        // 处理单元格数据
        console.log('处理单元格数据，单元格数量:', Object.keys(cellData).length);
        for (const [key, cell] of Object.entries(cellData)) {
            try {
                const [ri, ci] = key.split('_').map(Number);

                // 确保行存在
                if (!rows[ri]) {
                    rows[ri] = { cells: {} };
                }

                // 确保cells对象存在
                if (!rows[ri].cells) {
                    rows[ri].cells = {};
                }

                // 调试信息：输出单元格数据
                console.log(`处理单元格 ${key}:`, cell);
                console.log(`  - 行列: (${ri}, ${ci})`);
                console.log(`  - 值: ${cell.v}`);
                console.log(`  - 边框数据:`, cell.bd);

                // 转换样式并添加到样式表
                const style = convertCellStyle(cell);
                console.log(`  - 转换后的样式:`, style);

                // 构建单元格数据
                const cellObj = {
                    text: cell.v !== undefined ? String(cell.v) : ''
                };

                // 重要：如果单元格有边框或其他样式，即使没有内容也要保留单元格对象
                // 使用空对象而不是完全移除，确保 x-spreadsheet 能渲染边框
                if (style && Object.keys(style).length > 0) {
                    // 查找是否已有相同样式
                    let styleIndex = styles.findIndex(s => JSON.stringify(s) === JSON.stringify(style));
                    if (styleIndex === -1) {
                        styleIndex = styles.length;
                        styles.push(style);
                    }
                    cellObj.style = styleIndex;
                    console.log(`  - 应用样式索引: ${styleIndex}`);
                }

                rows[ri].cells[ci] = cellObj;
                console.log(`  - 最终单元格对象:`, cellObj);
            } catch (cellError) {
                console.error(`处理单元格 ${key} 时出错:`, cellError);
            }
        }

        // 准备列宽数据
        const cols = {};
        const columnData = sheetData.columnData || {};
        console.log('处理列宽数据:', columnData);
        for (const [colIdx, width] of Object.entries(columnData)) {
            cols[colIdx] = { width: width };
        }

        const result = [{
            name: sheetData.name || 'Sheet1',
            rows: rows,
            cols: cols,
            styles: styles,
            merges: sheetData.mergeData ? sheetData.mergeData.map(m =>
                `${String.fromCharCode(65 + m.startColumn)}${m.startRow + 1}:${String.fromCharCode(65 + m.endColumn)}${m.endRow + 1}`
            ) : []
        }];

        console.log('数据转换完成，结果:', result);
        return result;
    } catch (error) {
        console.error('convertToXsFormat 错误:', error);
        console.error('错误堆栈:', error.stack);
        throw error;
    }
}

// 转换单元格样式到x-spreadsheet格式
function convertCellStyle(cell) {
    const style = {};

    // 字体加粗
    if (cell.bl) {
        style.font = style.font || {};
        style.font.bold = true;
    }

    // 斜体
    if (cell.it) {
        style.font = style.font || {};
        style.font.italic = true;
    }

    // 字体大小
    if (cell.fs) {
        style.font = style.font || {};
        style.font.size = cell.fs;
    }

    // 字体名称
    if (cell.ff) {
        style.font = style.font || {};
        style.font.name = cell.ff;
    }

    // 字体颜色
    if (cell.cl && cell.cl.rgb) {
        style.color = '#' + cell.cl.rgb.slice(-6);
    }

    // 背景色
    if (cell.bg && cell.bg.rgb) {
        style.bgcolor = '#' + cell.bg.rgb.slice(-6);
    }

    // 对齐方式
    if (cell.ht) {
        // 水平对齐: left, center, right
        style.align = cell.ht;
    }

    if (cell.vt) {
        // 垂直对齐: top, middle, bottom
        style.valign = cell.vt;
    }

    // 边框
    if (cell.bd) {
        console.log('    处理边框数据:', cell.bd);
        style.border = {};
        const bd = cell.bd;

        // 转换每条边
        ['t', 'b', 'l', 'r'].forEach(side => {
            if (bd[side]) {
                const borderStyle = bd[side].s === 1 ? 'thin' : 'medium';
                const borderColor = bd[side].cl && bd[side].cl.rgb ?
                    '#' + bd[side].cl.rgb.slice(-6) : '#000000';

                const sideName = {
                    't': 'top',
                    'b': 'bottom',
                    'l': 'left',
                    'r': 'right'
                }[side];

                style.border[sideName] = [borderStyle, borderColor];
                console.log(`      边框 ${side} -> ${sideName}: ${borderStyle}, ${borderColor}`);
            }
        });
        console.log('    最终边框样式:', style.border);
    }

    // 文字换行
    if (cell.tb === '2') {
        style.textwrap = true;
    }

    // 下划线
    if (cell.ul) {
        style.underline = true;
    }

    // 删除线
    if (cell.st) {
        style.strike = true;
    }

    return style;
}

// 转换样式
function convertStyle(style) {
    if (!style || typeof style !== 'object') {
        return undefined;
    }

    const result = {};

    try {
        if (style.bl) result.font = { ...result.font, bold: true };
        if (style.it) result.font = { ...result.font, italic: true };
        if (style.fs) result.font = { ...result.font, size: style.fs };
        if (style.bg && style.bg.rgb && typeof style.bg.rgb === 'string') {
            const rgb = style.bg.rgb;
            if (/^[0-9A-Fa-f]{6,8}$/.test(rgb)) {
                result.bgcolor = '#' + rgb.slice(-6);
            }
        }
        if (style.cl && style.cl.rgb && typeof style.cl.rgb === 'string') {
            const rgb = style.cl.rgb;
            if (/^[0-9A-Fa-f]{6,8}$/.test(rgb)) {
                result.color = '#' + rgb.slice(-6);
            }
        }
    } catch (e) {
        console.warn('转换样式出错:', e);
    }

    return Object.keys(result).length > 0 ? result : undefined;
}

// 连接WebSocket
async function connectWebSocket() {
    const mac = await getMacAddress();
    const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
    const wsUrl = `${protocol}//${window.location.host}/ws/${sheetKey}?mac=${encodeURIComponent(mac)}`;

    websocket = new WebSocket(wsUrl);

    websocket.onopen = () => {
        console.log('WebSocket连接已建立');
        startHeartbeat();
    };

    websocket.onmessage = (event) => {
        handleWebSocketMessage(JSON.parse(event.data));
    };

    websocket.onclose = (event) => {
        console.log('WebSocket连接已关闭', event.code, event.reason);
        showToast('连接已断开，正在重连...', 'info');
        setTimeout(connectWebSocket, 3000);
    };

    websocket.onerror = (error) => {
        console.error('WebSocket错误:', error);
    };
}

// 心跳保持连接
function startHeartbeat() {
    setInterval(() => {
        if (websocket && websocket.readyState === WebSocket.OPEN) {
            websocket.send(JSON.stringify({ type: 'ping' }));
        }
    }, 30000);
}

// 处理WebSocket消息
function handleWebSocketMessage(message) {
    switch (message.type) {
        case 'connected':
            currentUserId = message.user_id;
            updateOnlineUsers(message.online_users);
            showToast(`你的标识: ${message.display_name}`, 'info');
            break;

        case 'user_join':
            updateOnlineUsers(message.online_users);
            showToast(`${message.display_name} 加入了协作`, 'info');
            break;

        case 'user_leave':
            updateOnlineUsers(message.online_users);
            showToast(`${message.display_name} 离开了`, 'info');
            break;

        case 'cell_update':
            applyRemoteCellUpdate(message);
            break;

        case 'batch_update':
            applyRemoteBatchUpdate(message);
            break;

        case 'selection_change':
            showCollaboratorSelection(message);
            break;

        case 'pong':
            // 心跳响应
            break;

        case 'history_update':
            updateHistoryPanel(message.history);
            break;

        case 'dimension_update':
            applyRemoteDimensionUpdate(message);
            break;
    }
}

// 发送单元格更新
function sendCellUpdate(row, col, value, style = null) {
    if (websocket && websocket.readyState === WebSocket.OPEN) {
        websocket.send(JSON.stringify({
            type: 'cell_update',
            row: row,
            col: col,
            value: value,
            style: style
        }));
    }
}

// 发送选区变化
function sendSelectionChange(row, col) {
    if (websocket && websocket.readyState === WebSocket.OPEN) {
        websocket.send(JSON.stringify({
            type: 'selection_change',
            selection: { row, col }
        }));
    }
}

// 应用远程单元格更新
function applyRemoteCellUpdate(message) {
    if (!spreadsheet) return;

    isUpdatingFromRemote = true;

    try {
        const { row, col, value, display_name } = message;

        // 使用x-spreadsheet的API更新单元格
        spreadsheet.cellText(row, col, value !== null ? String(value) : '');

        // 强制重新渲染以显示更新
        spreadsheet.reRender();

        // 显示更新提示
        showCellUpdateIndicator(row, col, display_name);
    } catch (error) {
        console.error('应用远程更新失败:', error);
    } finally {
        isUpdatingFromRemote = false;
    }
}

// 应用远程批量更新
function applyRemoteBatchUpdate(message) {
    if (!spreadsheet) return;

    isUpdatingFromRemote = true;

    try {
        const { updates } = message;
        for (const update of updates) {
            spreadsheet.cellText(update.row, update.col, update.value !== null ? String(update.value) : '');
        }
        // 强制重新渲染以显示更新
        spreadsheet.reRender();
    } catch (error) {
        console.error('应用批量更新失败:', error);
    } finally {
        isUpdatingFromRemote = false;
    }
}

// 显示单元格更新指示器
function showCellUpdateIndicator(row, col, userName) {
    // 简单的闪烁效果（通过临时改变背景色实现）
    // x-spreadsheet 的API有限，这里只做简单提示
    console.log(`${userName} 更新了单元格 (${row}, ${col})`);
}

// 显示协作者选区
function showCollaboratorSelection(message) {
    const { user_id, display_name, selection } = message;
    if (user_id === currentUserId) return;

    // 移除旧的选区指示器
    const oldIndicator = document.querySelector(`.collaborator-selection[data-user="${user_id}"]`);
    if (oldIndicator) {
        oldIndicator.remove();
    }

    // x-spreadsheet 的内部结构较复杂，这里简化处理
    // 实际项目中可能需要更深入的集成
}

// 更新在线用户列表
function updateOnlineUsers(users) {
    onlineUsers = users;
    document.getElementById('onlineCount').textContent = users.length;

    const listContainer = document.getElementById('onlineUsersList');
    listContainer.innerHTML = users.map(user => `
        <div class="online-user-item">
            <span class="dot" style="background: ${getUserColor(user.user_id)}"></span>
            <span>${user.display_name}</span>
        </div>
    `).join('');
}

// 导出表格
async function exportSheet() {
    try {
        const response = await fetch(`/api/sheet/${sheetKey}/export`);
        if (!response.ok) {
            throw new Error('导出失败');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${sessionStorage.getItem('sheet_name') || 'sheet'}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        showToast('导出成功', 'success');
    } catch (error) {
        showToast('导出失败: ' + error.message, 'error');
    }
}

// 退出/关闭页面
function logout() {
    if (websocket) {
        websocket.close();
    }
    sessionStorage.removeItem('sheet_key');
    sessionStorage.removeItem('sheet_name');
    window.close();
    // 如果window.close()不起作用（不是通过脚本打开的窗口），显示提示
    showToast('请关闭此标签页', 'info');
}

// 显示Toast通知
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// 页面关闭时清理
window.addEventListener('beforeunload', () => {
    // 发送待处理的更新（用户关闭页面前保存）
    if (pendingCellUpdate && websocket && websocket.readyState === WebSocket.OPEN) {
        websocket.send(JSON.stringify({
            type: 'cell_update',
            row: pendingCellUpdate.row,
            col: pendingCellUpdate.col,
            value: pendingCellUpdate.value,
            style: null
        }));
        pendingCellUpdate = null;
    }
    if (websocket) {
        websocket.close();
    }
});

// ==================== 历史记录功能 ====================

// 加载历史记录
async function loadHistory() {
    try {
        const response = await fetch(`/api/sheet/${sheetKey}/history`);
        const data = await response.json();
        updateHistoryPanel(data.history || []);
    } catch (error) {
        console.error('加载历史记录失败:', error);
    }
}

// 初始化历史面板
function initHistoryPanel() {
    const panel = document.getElementById('historyPanel');
    const toggleBtn = document.getElementById('historyToggle');

    // 默认最小化
    panel.classList.add('collapsed');
    toggleBtn.textContent = '+';

    // 折叠/展开切换
    toggleBtn.addEventListener('click', () => {
        panel.classList.toggle('collapsed');
        toggleBtn.textContent = panel.classList.contains('collapsed') ? '+' : '−';
    });

    // 可拖拽
    makeDraggable(panel);
}

// 更新历史记录面板
function updateHistoryPanel(history) {
    const listEl = document.getElementById('historyList');

    if (!history || history.length === 0) {
        listEl.innerHTML = '<li class="history-empty">暂无修改记录</li>';
        return;
    }

    listEl.innerHTML = history.slice(0, historyCount).map(item => `
        <li class="history-item">
            <div class="time">${item.timestamp || ''}</div>
            <div>
                <span class="user">${escapeHtml(item.user || '')}</span>
                <span class="action">${escapeHtml(item.action || '')}</span>
                <span class="cell">${escapeHtml(item.cell || '')}</span>
            </div>
            ${item.value ? `<div class="value">${escapeHtml(item.value)}</div>` : ''}
        </li>
    `).join('');
}

// HTML转义
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 使元素可拖拽
function makeDraggable(element) {
    // 查找header元素（支持不同类名）
    const header = element.querySelector('.history-header, .edit-header');
    if (!header) {
        console.warn('makeDraggable: 找不到header元素');
        return;
    }

    let isDragging = false;
    let offsetX, offsetY;

    header.addEventListener('mousedown', (e) => {
        if (e.target.tagName === 'BUTTON') return;
        isDragging = true;
        offsetX = e.clientX - element.offsetLeft;
        offsetY = e.clientY - element.offsetTop;
        element.style.transition = 'none';
    });

    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        let x = e.clientX - offsetX;
        let y = e.clientY - offsetY;

        // 边界检查
        x = Math.max(0, Math.min(x, window.innerWidth - element.offsetWidth));
        y = Math.max(0, Math.min(y, window.innerHeight - element.offsetHeight));

        element.style.left = x + 'px';
        element.style.top = y + 'px';
        element.style.right = 'auto';
        element.style.bottom = 'auto';
    });

    document.addEventListener('mouseup', () => {
        isDragging = false;
        element.style.transition = '';
    });
}

// ==================== 定时同步功能 ====================

// 启动定时同步
function startPeriodicSync() {
    setInterval(async () => {
        // 如果有待发送的更新，先发送
        if (pendingCellUpdate) {
            sendCellUpdate(pendingCellUpdate.row, pendingCellUpdate.col, pendingCellUpdate.value);
            pendingCellUpdate = null;
        }

        // 重新加载表格数据并同步
        try {
            const response = await fetch(`/api/sheet/${sheetKey}`);
            if (response.ok) {
                const sheetData = await response.json();
                syncSheetData(sheetData);
            }
        } catch (error) {
            console.error('定时同步失败:', error);
        }
    }, syncInterval);
}

// 同步表格数据（只更新有差异的单元格）
function syncSheetData(sheetData) {
    if (!spreadsheet || !sheetData) return;

    isUpdatingFromRemote = true;

    try {
        const cellData = sheetData.cellData || {};
        let hasChanges = false;

        for (const [key, cell] of Object.entries(cellData)) {
            const [ri, ci] = key.split('_').map(Number);
            const currentCell = spreadsheet.cell(ri, ci);
            const currentValue = currentCell ? (currentCell.text || '') : '';
            const newValue = cell.v !== undefined ? String(cell.v) : '';

            // 如果值不同，更新单元格
            if (currentValue !== newValue) {
                spreadsheet.cellText(ri, ci, newValue);
                hasChanges = true;
            }
        }

        // 如果有变化，重新渲染
        if (hasChanges) {
            spreadsheet.reRender();
        }
    } catch (error) {
        console.error('同步数据失败:', error);
    } finally {
        isUpdatingFromRemote = false;
    }
}

// ==================== 编辑窗口功能 ====================

// 初始化编辑窗口
function initEditPanel() {
    const panel = document.getElementById('editPanel');
    const toggleBtn = document.getElementById('editToggle');
    const textarea = document.getElementById('editTextarea');

    // 折叠/展开切换
    toggleBtn.addEventListener('click', () => {
        panel.classList.toggle('collapsed');
        toggleBtn.textContent = panel.classList.contains('collapsed') ? '+' : '−';
    });

    // 可拖拽
    makeDraggable(panel);

    // 监听编辑窗口的输入事件
    textarea.addEventListener('input', () => {
        if (isEditTextareaUpdating || currentEditCell.row < 0 || currentEditCell.col < 0) {
            return;
        }

        // 将编辑窗口的内容同步到单元格
        const value = textarea.value;
        if (spreadsheet && !isUpdatingFromRemote) {
            isEditTextareaUpdating = true;
            spreadsheet.cellText(currentEditCell.row, currentEditCell.col, value);
            spreadsheet.reRender();
            // 保存待发送的更新
            pendingCellUpdate = {
                row: currentEditCell.row,
                col: currentEditCell.col,
                value: value
            };
            isEditTextareaUpdating = false;
        }
    });

    // 监听键盘事件 - 回车插入换行而不是结束编辑
    textarea.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            // 阻止默认行为，插入换行
            e.preventDefault();
            const start = textarea.selectionStart;
            const end = textarea.selectionEnd;
            const value = textarea.value;

            // 在光标位置插入换行符
            textarea.value = value.substring(0, start) + '\n' + value.substring(end);

            // 将光标移动到换行符后面
            textarea.selectionStart = textarea.selectionEnd = start + 1;

            // 触发input事件以同步到单元格
            textarea.dispatchEvent(new Event('input'));
        }
    });
}

// 更新编辑窗口
function updateEditPanel(row, col) {
    if (isEditTextareaUpdating) return;

    currentEditCell.row = row;
    currentEditCell.col = col;

    const cellInfo = document.getElementById('editCellInfo');
    const textarea = document.getElementById('editTextarea');

    // 获取列字母（A, B, C...）
    const colLetter = String.fromCharCode(65 + col);
    cellInfo.textContent = `${colLetter}${row + 1}`;

    // 获取单元格值并显示在编辑窗口
    if (spreadsheet) {
        isEditTextareaUpdating = true;
        const cellData = spreadsheet.cell(row, col);
        textarea.value = cellData ? (cellData.text || '') : '';
        isEditTextareaUpdating = false;
    }
}

// ==================== 列宽行高同步功能 ====================

// 检查并同步列宽行高
function checkAndSyncDimensions() {
    if (!spreadsheet || isUpdatingFromRemote) return;

    try {
        const data = spreadsheet.getData();
        if (!data || !data[0]) return;

        const currentSheet = data[0];
        const currentCols = currentSheet.cols || {};
        const currentRows = currentSheet.rows || {};

        // 检查列宽变化
        const colChanges = {};
        for (const colIdx in currentCols) {
            if (currentCols[colIdx] && currentCols[colIdx].width !== undefined) {
                const currentWidth = currentCols[colIdx].width;
                const lastWidth = lastDimensions.cols[colIdx];
                if (currentWidth !== lastWidth) {
                    colChanges[colIdx] = currentWidth;
                }
            }
        }

        // 检查行高变化
        const rowChanges = {};
        for (const rowIdx in currentRows) {
            if (currentRows[rowIdx] && currentRows[rowIdx].height !== undefined) {
                const currentHeight = currentRows[rowIdx].height;
                const lastHeight = lastDimensions.rows[rowIdx];
                if (currentHeight !== lastHeight) {
                    rowChanges[rowIdx] = currentHeight;
                }
            }
        }

        // 如果有变化，发送更新
        if (Object.keys(colChanges).length > 0 || Object.keys(rowChanges).length > 0) {
            sendDimensionUpdate(colChanges, rowChanges);
            // 更新记录
            Object.assign(lastDimensions.cols, colChanges);
            Object.assign(lastDimensions.rows, rowChanges);
        }
    } catch (error) {
        console.error('检查列宽行高失败:', error);
    }
}

// 发送列宽行高更新
function sendDimensionUpdate(colWidths, rowHeights) {
    if (websocket && websocket.readyState === WebSocket.OPEN) {
        websocket.send(JSON.stringify({
            type: 'dimension_update',
            col_widths: colWidths,
            row_heights: rowHeights
        }));
    }
}

// 应用远程列宽行高更新
function applyRemoteDimensionUpdate(message) {
    if (!spreadsheet) return;

    isUpdatingFromRemote = true;

    try {
        const { col_widths, row_heights } = message;
        const data = spreadsheet.getData();
        if (!data || !data[0]) return;

        const currentSheet = data[0];

        // 应用列宽更新
        if (col_widths) {
            if (!currentSheet.cols) currentSheet.cols = {};
            for (const colIdx in col_widths) {
                if (!currentSheet.cols[colIdx]) currentSheet.cols[colIdx] = {};
                currentSheet.cols[colIdx].width = col_widths[colIdx];
                lastDimensions.cols[colIdx] = col_widths[colIdx];
            }
        }

        // 应用行高更新
        if (row_heights) {
            if (!currentSheet.rows) currentSheet.rows = {};
            for (const rowIdx in row_heights) {
                if (!currentSheet.rows[rowIdx]) currentSheet.rows[rowIdx] = {};
                currentSheet.rows[rowIdx].height = row_heights[rowIdx];
                lastDimensions.rows[rowIdx] = row_heights[rowIdx];
            }
        }

        // 重新加载数据以应用变化
        spreadsheet.loadData(data);
        spreadsheet.reRender();
    } catch (error) {
        console.error('应用列宽行高更新失败:', error);
    } finally {
        isUpdatingFromRemote = false;
    }
}

// 初始化列宽行高记录
function initializeDimensionsTracking() {
    if (!spreadsheet) return;

    try {
        const data = spreadsheet.getData();
        if (!data || !data[0]) return;

        const currentSheet = data[0];

        // 记录列宽
        lastDimensions.cols = {};
        const cols = currentSheet.cols || {};
        for (const colIdx in cols) {
            if (cols[colIdx] && cols[colIdx].width !== undefined) {
                lastDimensions.cols[colIdx] = cols[colIdx].width;
            }
        }

        // 记录每行的高度
        lastDimensions.rows = {};
        const rows = currentSheet.rows || {};
        for (const rowIdx in rows) {
            if (rows[rowIdx] && rows[rowIdx].height !== undefined) {
                lastDimensions.rows[rowIdx] = rows[rowIdx].height;
            }
        }

        // 启动定期检查
        setInterval(checkAndSyncDimensions, 2000);  // 每2秒检查一次
    } catch (error) {
        console.error('初始化列宽行高跟踪失败:', error);
    }
}
