// ==================== IndexedDB 双存储 ====================
const DB_NAME = 'LocalFileDB';
const DB_VERSION = 2;
const ITEM_STORE = 'excelFiles';
const REGION_STORE = 'regionFiles';

let db = null;
let allItemsCache = [];      // 存储所有物品 { id, name, imageUrl, rowData, keys, extraAttrs, skillAttrs, fileName }
let allRegionsCache = {};    // { regionName: { "附加属性": [attrs], "技能属性": [attrs] } }

// ---------- 数据库操作 ----------
function initDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onerror = () => reject('数据库打开失败');
        request.onsuccess = (e) => {
            db = e.target.result;
            resolve(db);
        };
        request.onupgradeneeded = (e) => {
            const dbRef = e.target.result;
            if (!dbRef.objectStoreNames.contains(ITEM_STORE)) {
                dbRef.createObjectStore(ITEM_STORE, { keyPath: 'id', autoIncrement: true });
            }
            if (!dbRef.objectStoreNames.contains(REGION_STORE)) {
                dbRef.createObjectStore(REGION_STORE, { keyPath: 'id', autoIncrement: true });
            }
        };
    });
}

function getAllFiles(storeName) {
    return new Promise((resolve, reject) => {
        if (!db) return reject('DB not ready');
        const tx = db.transaction([storeName], 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.getAll();
        req.onsuccess = () => resolve(req.result);
        req.onerror = (e) => reject(e);
    });
}

function saveFileToDB(storeName, file, blob) {
    return new Promise((resolve, reject) => {
        if (!db) return reject('DB not ready');
        const tx = db.transaction([storeName], 'readwrite');
        const store = tx.objectStore(storeName);
        const record = {
            name: file.name,
            type: file.type,
            size: file.size,
            timestamp: Date.now(),
            blob: blob
        };
        const req = store.add(record);
        req.onsuccess = (e) => resolve(e.target.result);
        req.onerror = (e) => reject(e);
    });
}

function deleteFileById(storeName, id) {
    return new Promise((resolve, reject) => {
        if (!db) return reject('DB not ready');
        const tx = db.transaction([storeName], 'readwrite');
        const store = tx.objectStore(storeName);
        const req = store.delete(id);
        req.onsuccess = () => resolve();
        req.onerror = (e) => reject(e);
    });
}

function getFileBlobById(storeName, id) {
    return new Promise((resolve, reject) => {
        if (!db) return reject('DB not ready');
        const tx = db.transaction([storeName], 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.get(id);
        req.onsuccess = () => {
            if (req.result && req.result.blob) resolve(req.result.blob);
            else reject('File not found');
        };
        req.onerror = (e) => reject(e);
    });
}

// ==================== 解析工具 ====================
function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>]/g, function(m) {
        if (m === '&') return '&amp;';
        if (m === '<') return '&lt;';
        if (m === '>') return '&gt;';
        return m;
    });
}

function extractName(row, keys) {
    const candidates = ['武器', '名称', 'name', '标题', 'title', 'Name'];
    for (let c of candidates) {
        if (row[c] && String(row[c]).trim()) return String(row[c]).trim();
    }
    if (keys.length >= 2) {
        const second = keys[1];
        if (row[second] && String(row[second]).trim()) return String(row[second]).trim();
    }
    return '未命名';
}

function parseItemExcelBlob(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            resolve(json);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob);
    });
}

function parseRegionExcelBlob(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:G30');
            const rows = [];
            for (let R = range.s.r; R <= range.e.r; R++) {
                const row = [];
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    const cell = sheet[cellAddress];
                    let val = cell ? (cell.v !== undefined ? String(cell.v).trim() : '') : '';
                    // 清除不可见字符（零宽空格等）
                    val = val.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
                    row.push(val);
                }
                rows.push(row);
            }
            if (rows.length < 2) return reject('表格数据不足');
            const headers = rows[0];
            const regionNames = [];
            for (let c = 1; c < headers.length; c++) {
                if (headers[c] && headers[c].trim()) regionNames.push(headers[c].trim());
            }
            if (regionNames.length === 0) return reject('未找到地名行');
            let currentCategory = '';
            const regionData = {};
            regionNames.forEach(name => { regionData[name] = {}; });
            for (let r = 1; r < rows.length; r++) {
                const row = rows[r];
                const firstCol = row[0] ? row[0].trim() : '';
                if (firstCol) {
                    currentCategory = firstCol;
                    regionNames.forEach(name => {
                        if (!regionData[name][currentCategory]) regionData[name][currentCategory] = [];
                    });
                }
                if (!currentCategory) continue;
                for (let c = 1; c < row.length; c++) {
                    let cellVal = row[c] ? row[c].trim() : '';
                    if (cellVal === '') continue;
                    // 再次清理
                    cellVal = cellVal.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
                    const regionName = regionNames[c-1];
                    if (regionName && regionData[regionName]) {
                        if (!regionData[regionName][currentCategory]) {
                            regionData[regionName][currentCategory] = [];
                        }
                        if (!regionData[regionName][currentCategory].includes(cellVal)) {
                            regionData[regionName][currentCategory].push(cellVal);
                        }
                    }
                }
            }refreshAllItems
            resolve(regionData);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob);
    });
}

// ==================== 物品合并展示 ====================
async function refreshAllItems() {
    const container = document.getElementById('cardListContainer');

    if (!container) return;
    try {
        const files = await getAllFiles(ITEM_STORE);
        if (!files.length) {
            container.innerHTML = '<div class="placeholder-text">暂无物品文件，请点击“上传物品表”</div>';
            document.getElementById('currentFileNameLabel').innerHTML = '📇 全部物品 (0)';
            allItemsCache = [];
            return;
        }
        let allItems = [];
        for (let file of files) {
            const blob = await getFileBlobById(ITEM_STORE, file.id);
            const jsonData = await parseItemExcelBlob(blob);
            if (jsonData && jsonData.length) {
                const rawKeys = Object.keys(jsonData[0]);
                const keys = rawKeys.filter(k => !k.startsWith('_EMPTY'));
                // 查找附加属性和技能属性列（支持列名包含“附加属性”、“技能属性”）
                const attrColKey = keys.find(k => k.includes('附加属性') || k === '附加属性');
                const skillColKey = keys.find(k => k.includes('技能属性') || k === '技能属性');
                const mainColKey = keys.find(k => k.includes('主属性') || k === '主属性');
                
                jsonData.forEach(row => {
                    const imageUrlKey = keys[0];
                    let imageUrl = row[imageUrlKey] ? String(row[imageUrlKey]).trim() : '';
                    const displayName = extractName(row, keys);
                    
                    // 提取附加属性数组（支持逗号、空格、顿号等分隔）
                    let extraAttrs = [];
                    if (attrColKey && row[attrColKey]) {
                        let raw = String(row[attrColKey]).trim();
                        raw = raw.replace(/[\u200B-\u200D\uFEFF]/g, ''); // 清除不可见字符
                        extraAttrs = raw.split(/[,，、\s\n]+/).filter(a => a !== '');
                    }
                    let skillAttrs = [];
                    if (skillColKey && row[skillColKey]) {
                        let raw = String(row[skillColKey]).trim();
                        raw = raw.replace(/[\u200B-\u200D\uFEFF]/g, '');
                        skillAttrs = raw.split(/[,，、\s\n]+/).filter(s => s !== '');
                    }
                    // 找到主属性列
                    const mainAttrKey = keys.find(k => k.includes('主属性') || k === '主属性');
                    let mainAttr = '';
                    if (mainAttrKey && row[mainAttrKey]) {
                        mainAttr = String(row[mainAttrKey]).trim();
                    }

                    allItems.push({
                        fileId: file.id,
                        fileName: file.name,
                        imageUrl: imageUrl,
                        name: displayName,
                        rowData: row,
                        keys: keys,
                        extraAttrs: extraAttrs,
                        skillAttrs: skillAttrs,
                        mainAttr: mainAttr     //新增
                    });
                });
            }
        }
        allItemsCache = allItems;
        document.getElementById('currentFileNameLabel').innerHTML = `📇 全部物品 (${allItems.length} 项)`;
        if (allItems.length === 0) {
            container.innerHTML = '<div class="placeholder-text">暂无物品数据</div>';
            return;
        }
        let cardsHtml = '';
        allItems.forEach((item, idx) => {
            const validUrl = item.imageUrl && (item.imageUrl.startsWith('http') || item.imageUrl.startsWith('data:image'));
            cardsHtml += `
                <div class="list-card" data-index="${idx}">
                    <div class="list-card-img">
                        ${validUrl ? `<img src="${escapeHtml(item.imageUrl)}" onerror="this.onerror=null; this.parentElement.innerHTML='📷';">` : '📷'}
                    </div>
                    <div class="list-card-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</div>
                </div>
            `;
        });
        container.innerHTML = cardsHtml;
        // 绑定点击事件
        document.querySelectorAll('.list-card').forEach(card => {
            card.addEventListener('click', (e) => {
                const idx = parseInt(card.getAttribute('data-index'));
                const item = allItemsCache[idx];
                if (item) {
                    showItemDetail(item);
                    document.querySelectorAll('.list-card').forEach(c => c.classList.remove('active'));
                    card.classList.add('active');
                }
            });
        });
        if (allItems.length > 0) {
            const first = document.querySelector('.list-card[data-index="0"]');
            if (first) first.click();
        }
    } catch (err) {
        console.error('refreshAllItems error:', err);
        container.innerHTML = '<div class="placeholder-text">加载失败，请检查控制台</div>';
    }
}
// 获取所有可能的主属性组合（5选3）
function getAllMainAttrCombinations() {
    const uniqueMain = [...new Set(allItemsCache.map(item => item.mainAttr).filter(Boolean))];
    if (uniqueMain.length === 0) return [];
    if (uniqueMain.length <= 3) return [uniqueMain];  // 少于等于3种时，全选
    const combinations = [];
    const n = uniqueMain.length;
    for (let i = 0; i < n; i++) {
        for (let j = i+1; j < n; j++) {
            for (let k = j+1; k < n; k++) {
                combinations.push([uniqueMain[i], uniqueMain[j], uniqueMain[k]]);
            }
        }
    }
    return combinations;
}

// 计算对于给定地域，最佳锁定组合及覆盖物品数
function computeBestCoverageForRegion(regionName, currentItem) {
    const region = allRegionsCache[regionName];
    if (!region) return [];

    let regionAttrs = [];
    let regionSkills = [];
    for (const [catName, attrs] of Object.entries(region)) {
        if (catName.includes('附加属性')) regionAttrs = attrs.map(a => String(a).trim());
        if (catName.includes('技能属性')) regionSkills = attrs.map(a => String(a).trim());
    }

    function isItemFullyContained(item) {
        const itemAttrs = item.extraAttrs || [];
        const itemSkills = item.skillAttrs || [];
        const attrsOk = itemAttrs.length === 0 || itemAttrs.every(attr => regionAttrs.includes(attr));
        const skillsOk = itemSkills.length === 0 || itemSkills.every(skill => regionSkills.includes(skill));
        return attrsOk && skillsOk;
    }

    const allPossibleLockValues = [...new Set([...regionAttrs, ...regionSkills])];
    if (allPossibleLockValues.length === 0) return [];

    const mainCombos = getAllMainAttrCombinations();
    let bestCount = -1;
    let bestCombos = [];

    for (const mainSet of mainCombos) {
        for (const lockVal of allPossibleLockValues) {
            let covered = [];
            for (const item of allItemsCache) {
                if (!isItemFullyContained(item)) continue;
                if (!item.mainAttr) continue;
                if (!mainSet.includes(item.mainAttr)) continue;
                const hasAttr = (item.extraAttrs && item.extraAttrs.includes(lockVal)) ||
                                (item.skillAttrs && item.skillAttrs.includes(lockVal));
                if (!hasAttr) continue;
                covered.push(item);
            }
            // 必须包含当前物品
            const containsCurrent = covered.some(c => c.name === currentItem.name && c.fileId === currentItem.fileId);
            if (!containsCurrent) continue;
            if (covered.length > bestCount) {
                bestCount = covered.length;
                bestCombos = [{ mainSelected: mainSet, lockValue: lockVal, count: covered.length, coveredItems: covered }];
            } else if (covered.length === bestCount && bestCount >= 0) {
                const exists = bestCombos.some(c =>
                    JSON.stringify(c.mainSelected) === JSON.stringify(mainSet) && c.lockValue === lockVal
                );
                if (!exists) {
                    bestCombos.push({ mainSelected: mainSet, lockValue: lockVal, count: covered.length, coveredItems: covered });
                }
            }
        }
    }

    // 如果没有包含当前物品的组合，则降级为不要求包含当前物品（仅兜底）
    if (bestCombos.length === 0) {
        for (const mainSet of mainCombos) {
            for (const lockVal of allPossibleLockValues) {
                let covered = [];
                for (const item of allItemsCache) {
                    if (!isItemFullyContained(item)) continue;
                    if (!item.mainAttr) continue;
                    if (!mainSet.includes(item.mainAttr)) continue;
                    const hasAttr = (item.extraAttrs && item.extraAttrs.includes(lockVal)) ||
                                    (item.skillAttrs && item.skillAttrs.includes(lockVal));
                    if (hasAttr) covered.push(item);
                }
                if (covered.length > bestCount) {
                    bestCount = covered.length;
                    bestCombos = [{ mainSelected: mainSet, lockValue: lockVal, count: covered.length, coveredItems: covered }];
                } else if (covered.length === bestCount && bestCount >= 0) {
                    const exists = bestCombos.some(c =>
                        JSON.stringify(c.mainSelected) === JSON.stringify(mainSet) && c.lockValue === lockVal
                    );
                    if (!exists) bestCombos.push({ mainSelected: mainSet, lockValue: lockVal, count: covered.length, coveredItems: covered });
                }
            }
        }
    }

    return bestCombos;
}

function showCoveredItemsModal(items, title) {
    // 创建遮罩层
    let modal = document.getElementById('coveredItemsModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'coveredItemsModal';
        modal.style.position = 'fixed';
        modal.style.top = '0';
        modal.style.left = '0';
        modal.style.width = '100%';
        modal.style.height = '100%';
        modal.style.backgroundColor = 'rgba(0,0,0,0.5)';
        modal.style.display = 'flex';
        modal.style.alignItems = 'center';
        modal.style.justifyContent = 'center';
        modal.style.zIndex = '1000';
        modal.style.padding = '20px';
        document.body.appendChild(modal);
    }
    const modalContent = document.createElement('div');
    modalContent.style.backgroundColor = '#fff';
    modalContent.style.borderRadius = '20px';
    modalContent.style.maxWidth = '800px';
    modalContent.style.width = '90%';
    modalContent.style.maxHeight = '80%';
    modalContent.style.overflow = 'auto';
    modalContent.style.padding = '20px';
    modalContent.style.position = 'relative';
    modalContent.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
            <h3 style="margin:0;">${escapeHtml(title)}</h3>
            <button id="closeModalBtn" style="background:none; border:none; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(100px, 1fr)); gap: 12px;">
            ${items.map(item => `
                <div style="text-align: center;">
                    <div style="width: 80px; height: 80px; margin: 0 auto; background: #f1f5f9; border-radius: 12px; overflow: hidden;">
                        ${item.imageUrl && (item.imageUrl.startsWith('http') || item.imageUrl.startsWith('data:image')) ? `<img src="${escapeHtml(item.imageUrl)}" style="width:100%;height:100%;object-fit:cover;">` : '📷'}
                    </div>
                    <div style="font-size: 0.7rem; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(item.name)}</div>
                </div>
            `).join('')}
        </div>
    `;
    modal.innerHTML = '';
    modal.appendChild(modalContent);
    modal.style.display = 'flex';
    const closeBtn = modalContent.querySelector('#closeModalBtn');
    closeBtn.onclick = () => {
        modal.style.display = 'none';
    };
    modal.onclick = (e) => {
        if (e.target === modal) modal.style.display = 'none';
    };
}


// 显示物品详情并匹配地域
function showItemDetail(item) {
    const row = item.rowData;
    const keys = item.keys;
    const imageUrl = item.imageUrl;
    const validUrl = imageUrl && (imageUrl.startsWith('http') || imageUrl.startsWith('data:image'));
    
    const itemAttrs = item.extraAttrs || [];
    const itemSkills = item.skillAttrs || [];
    
    // 查找所有满足包含关系的地域（物品属性 ⊆ 地域属性）
    const matchedRegions = [];
    for (const [regionName, categories] of Object.entries(allRegionsCache)) {
        let regionAttrs = [];
        let regionSkills = [];
        for (const [catName, attrs] of Object.entries(categories)) {
            if (catName.includes('附加属性')) {
                regionAttrs = attrs.map(a => String(a).trim());
            }
            if (catName.includes('技能属性')) {
                regionSkills = attrs.map(a => String(a).trim());
            }
        }
        const allAttrsContained = itemAttrs.length === 0 || itemAttrs.every(attr => regionAttrs.includes(attr));
        const allSkillsContained = itemSkills.length === 0 || itemSkills.every(skill => regionSkills.includes(skill));
        if (allAttrsContained && allSkillsContained) {
            matchedRegions.push(regionName);
        }
    }

    // 为每个匹配的地域计算最优覆盖建议
    const coverageMap = {};
    for (const regionName of matchedRegions) {
        coverageMap[regionName] = computeBestCoverageForRegion(regionName, item);
    }

    // 开始构建详情HTML
    let detailsHtml = `
        <div class="detail-card">
            <div class="detail-image">
                ${validUrl ? `<img src="${escapeHtml(imageUrl)}" alt="图片" onerror="this.onerror=null; this.src='';">` : '<div style="background:#f1f5f9; padding:30px; text-align:center;">📷 无图片</div>'}
            </div>
    `;

    // 显示所有属性（跳过第一列图片URL，并过滤空列）
    for (let i = 1; i < keys.length; i++) {
        const key = keys[i];
        if (key.startsWith('_EMPTY')) continue;
        let value = row[key] !== undefined && row[key] !== null ? String(row[key]) : '';
        // 如果值为空，跳过这一行（避免显示空字段）
        if (value === '') continue;
        detailsHtml += `
            <div class="detail-field">
                <div class="field-label">${escapeHtml(key)}</div>
                <div class="field-value">${escapeHtml(value)}</div>
            </div>
        `;
    }

    // 显示匹配的地域及最优覆盖建议
    if (matchedRegions.length > 0) {
        detailsHtml += `
            <div class="detail-field" style="margin-top: 12px; border-top: 2px solid #e9edf2; padding-top: 12px;">
                <div class="field-label" style="color: #10b981;">🏷️ 完全包含该物品属性的地域</div>
                <div class="field-value">${matchedRegions.map(r => escapeHtml(r)).join('、')}</div>
            </div>
            <div style="margin-top: 16px; background: #fefce8; border-radius: 16px; padding: 12px;">
                <div style="font-weight: 600; margin-bottom: 8px;">📊 最优覆盖建议（针对当前地域）</div>
        `;
        for (const regionName of matchedRegions) {
            const bestList = coverageMap[regionName];
            if (bestList && bestList.length > 0) {
                // 过滤掉覆盖数量为1的方案
                const filteredBestList = bestList.filter(best => best.count >= 2);
                if (filteredBestList.length === 0) {
                    detailsHtml += `
                        <div style="margin-bottom: 12px; border-left: 3px solid #f97316; padding-left: 10px;">
                            <div><strong>${escapeHtml(regionName)}</strong></div>
                            <div style="font-size: 0.75rem; color: #475569;">⚠️ 仅能覆盖当前物品，无更多组合</div>
                        </div>
                    `;
                    continue;
                }
                for (let idx = 0; idx < filteredBestList.length; idx++) {
                    const best = filteredBestList[idx];
                    detailsHtml += `
                        <div style="margin-bottom: 16px; border-left: 3px solid #10b981; padding-left: 10px;">
                            <div><strong>${escapeHtml(regionName)}${filteredBestList.length > 1 ? ` 方案 ${idx+1}` : ''}</strong></div>
                            <div style="font-size: 0.75rem; color: #475569; margin: 4px 0 8px 0;">
                                主属性选：${best.mainSelected.join('、')} &nbsp;|&nbsp;
                                锁定附加/技能属性：${escapeHtml(best.lockValue)}<br>
                                ✅ 可覆盖 <strong>${best.count}</strong> 个物品
                            </div>
                            <div class="covered-items-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 8px; max-height: 300px; overflow-y: auto; padding: 4px;">
                    `;
                    for (const coveredItem of best.coveredItems) {
                        const validCoveredUrl = coveredItem.imageUrl && (coveredItem.imageUrl.startsWith('http') || coveredItem.imageUrl.startsWith('data:image'));
                        detailsHtml += `
                            <div class="covered-item-card" data-item-name="${escapeHtml(coveredItem.name)}" data-file-id="${coveredItem.fileId}" style="text-align: center; cursor: pointer;">
                                <div style="width: 64px; height: 64px; margin: 0 auto; background: #f1f5f9; border-radius: 10px; overflow: hidden;">
                                    ${validCoveredUrl ? `<img src="${escapeHtml(coveredItem.imageUrl)}" style="width:100%;height:100%;object-fit:cover;">` : '📷'}
                                </div>
                                <div style="font-size: 0.65rem; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(coveredItem.name)}</div>
                            </div>
                        `;
                    }
                    detailsHtml += `</div></div>`;
                }
            } else {
                detailsHtml += `
                    <div style="margin-bottom: 12px; border-left: 3px solid #f97316; padding-left: 10px;">
                        <div><strong>${escapeHtml(regionName)}</strong></div>
                        <div style="font-size: 0.75rem; color: #475569;">⚠️ 无可用的锁定属性组合</div>
                    </div>
                `;
            }
        }
        detailsHtml += `</div>`;
    } else {
        detailsHtml += `
            <div class="detail-field" style="margin-top: 12px; border-top: 2px solid #e9edf2; padding-top: 12px;">
                <div class="field-label" style="color: #94a3b8;">🏷️ 完全包含该物品属性的地域</div>
                <div class="field-value">无</div>
            </div>
        `;
    }
    detailsHtml += `</div>`;  // 关闭 detail-card

    document.getElementById('detailContent').innerHTML = detailsHtml;
    document.getElementById('detailHint').innerText = `当前查看: ${item.name} (来自 ${item.fileName})`;

    // 为所有动态生成的覆盖物品卡片绑定浮窗和点击事件
    setTimeout(() => {
        const cards = document.querySelectorAll('.covered-item-card');
        cards.forEach(card => {
            const itemName = card.getAttribute('data-item-name');
            const fileId = parseInt(card.getAttribute('data-file-id'));
            const matchedItem = allItemsCache.find(i => i.name === itemName && i.fileId === fileId);
            if (matchedItem) {
                bindTooltipToCard(card, matchedItem);
                card.addEventListener('click', (e) => {
                    e.stopPropagation();
                    showItemDetail(matchedItem);
                    const leftCard = document.querySelector(`.list-card[data-index="${allItemsCache.indexOf(matchedItem)}"]`);
                    if (leftCard) {
                        document.querySelectorAll('.list-card').forEach(c => c.classList.remove('active'));
                        leftCard.classList.add('active');
                    }
                });
            }
        });
    }, 50);
}

// ==================== 浮窗功能（全局统一） ====================
let currentTooltip = null;

function showItemTooltip(item, event) {
    if (currentTooltip) currentTooltip.remove();
    const div = document.createElement('div');
    div.className = 'item-tooltip';
    const imageHtml = (item.imageUrl && (item.imageUrl.startsWith('http') || item.imageUrl.startsWith('data:image')))
        ? `<img src="${escapeHtml(item.imageUrl)}" alt="">`
        : `<div style="width:48px;height:48px;background:#f1f5f9;border-radius:12px;display:flex;align-items:center;justify-content:center;">📷</div>`;
    div.innerHTML = `
        <div class="tooltip-header">
            ${imageHtml}
            <div>
                <div class="tooltip-name">${escapeHtml(item.name)}</div>
                <div class="tooltip-main">⭐ 主属性：${escapeHtml(item.mainAttr || '无')}</div>
            </div>
        </div>
        <div class="tooltip-attr">
            <strong>附加属性</strong>：${(item.extraAttrs || []).join('、') || '无'}
        </div>
        <div class="tooltip-attr">
            <strong>技能属性</strong>：${(item.skillAttrs || []).join('、') || '无'}
        </div>
    `;
    document.body.appendChild(div);
    currentTooltip = div;
    // 定位在鼠标右下方
    div.style.left = (event.clientX + 15) + 'px';
    div.style.top = (event.clientY + 15) + 'px';
}

function hideTooltip() {
    if (currentTooltip) {
        currentTooltip.remove();
        currentTooltip = null;
    }
}

// 为任意容器内的 .item-card 类绑定浮窗（注意：我们的卡片类为 .list-card，但覆盖物品网格中的卡片类需要统一）
// 为了通用，我们编写一个函数，给任何带有 data-item-index 或 data-item 的元素绑定
function bindTooltipToCard(cardElement, itemData) {
    if (!cardElement || !itemData) return;
    cardElement.addEventListener('mouseenter', (e) => {
        showItemTooltip(itemData, e);
    });
    cardElement.addEventListener('mouseleave', hideTooltip);
}

// 批量绑定（用于动态生成的卡片列表）
function bindTooltipToCards(containerSelector, itemDataArray, cardSelector = '.list-card') {
    const container = document.querySelector(containerSelector);
    if (!container) return;
    const cards = container.querySelectorAll(cardSelector);
    cards.forEach((card, idx) => {
        const item = itemDataArray[idx];
        if (item) bindTooltipToCard(card, item);
    });
}

// ==================== 地域合并展示 ====================
async function refreshAllRegions() {
    const container = document.getElementById('regionFileListContainer');
    if (!container) return;
    try {
        const files = await getAllFiles(REGION_STORE);
        if (!files.length) {
            container.innerHTML = '<div class="placeholder-text">暂无地域文件，请点击“上传地域表”</div>';
            allRegionsCache = {};
            return;
        }
        const mergedRegionData = {};
        for (let file of files) {
            const blob = await getFileBlobById(REGION_STORE, file.id);
            const regionData = await parseRegionExcelBlob(blob);
            for (let [regionName, categories] of Object.entries(regionData)) {
                if (!mergedRegionData[regionName]) mergedRegionData[regionName] = {};
                for (let [cat, attrs] of Object.entries(categories)) {
                    if (!mergedRegionData[regionName][cat]) mergedRegionData[regionName][cat] = [];
                    for (let attr of attrs) {
                        if (!mergedRegionData[regionName][cat].includes(attr)) {
                            mergedRegionData[regionName][cat].push(attr);
                        }
                    }
                }
            }
        }
        allRegionsCache = mergedRegionData;
        const regionNames = Object.keys(mergedRegionData);
        if (regionNames.length === 0) {
            container.innerHTML = '<div class="placeholder-text">无有效地名</div>';
            return;
        }
        let html = '';
        regionNames.forEach(name => {
            html += `
                <div class="region-item" data-region="${escapeHtml(name)}">
                    <div class="region-name">📍 ${escapeHtml(name)}</div>
                </div>
            `;
        });
        container.innerHTML = html;
        document.querySelectorAll('.region-item[data-region]').forEach(item => {
            item.addEventListener('click', () => {
                const regionName = item.getAttribute('data-region');
                showRegionDetail(regionName);
            });
        });
    } catch (err) {
        console.error(err);
        container.innerHTML = '<div class="placeholder-text">加载失败</div>';
    }
}

function showRegionDetail(regionName) {
    if (!allRegionsCache[regionName]) return;
    const data = allRegionsCache[regionName];

    let regionAttrs = [];
    let regionSkills = [];
    for (const [catName, attrs] of Object.entries(data)) {
        if (catName.includes('附加属性')) regionAttrs = attrs.map(a => String(a).trim());
        if (catName.includes('技能属性')) regionSkills = attrs.map(a => String(a).trim());
    }

    const matchedItems = [];
    for (const item of allItemsCache) {
        const itemAttrs = item.extraAttrs || [];
        const itemSkills = item.skillAttrs || [];
        const attrsOk = itemAttrs.length === 0 || itemAttrs.every(attr => regionAttrs.includes(attr));
        const skillsOk = itemSkills.length === 0 || itemSkills.every(skill => regionSkills.includes(skill));
        if (attrsOk && skillsOk) matchedItems.push(item);
    }

    let html = `<div class="region-detail-card"><div class="region-detail-title">📍 ${escapeHtml(regionName)}</div>`;
    for (let [category, attrs] of Object.entries(data)) {
        if (attrs && attrs.length) {
            html += `<div class="category-block"><div class="category-title">${escapeHtml(category)}</div><div class="attribute-list">${attrs.map(attr => `<span class="attribute-tag">${escapeHtml(attr)}</span>`).join('')}</div></div>`;
        }
    }
    html += `</div>`;

    if (matchedItems.length > 0) {
        html += `<div style="margin-top: 20px; background: #f0fdf4; border-radius: 16px; padding: 12px;">
            <div style="font-weight: 600; margin-bottom: 8px;">📦 该地域包含的物品 (${matchedItems.length})</div>
            <div id="regionItemsGrid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 8px; max-height: 300px; overflow-y: auto; padding: 4px;">`;
        for (let idx = 0; idx < matchedItems.length; idx++) {
            const item = matchedItems[idx];
            const validUrl = item.imageUrl && (item.imageUrl.startsWith('http') || item.imageUrl.startsWith('data:image'));
            html += `
                <div class="region-item-card" data-item-index="${idx}" style="text-align: center; cursor: pointer;">
                    <div style="width: 64px; height: 64px; margin: 0 auto; background: #f1f5f9; border-radius: 10px; overflow: hidden;">
                        ${validUrl ? `<img src="${escapeHtml(item.imageUrl)}" style="width:100%;height:100%;object-fit:cover;">` : '📷'}
                    </div>
                    <div style="font-size: 0.65rem; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(item.name)}</div>
                </div>
            `;
        }
        html += `</div></div>`;
    } else {
        html += `<div style="margin-top: 20px; background: #f1f5f9; border-radius: 16px; padding: 12px; text-align: center; color: #64748b;">📦 该地域暂无包含任何物品</div>`;
    }

    document.getElementById('detailContent').innerHTML = html;
    document.getElementById('detailHint').innerText = `当前查看: ${regionName}`;

    // 为地域物品网格中的每个卡片绑定浮窗和点击事件
    const regionItemsContainer = document.getElementById('regionItemsGrid');
    if (regionItemsContainer) {
        const cards = regionItemsContainer.querySelectorAll('.region-item-card');
        cards.forEach((card, idx) => {
            const item = matchedItems[idx];
            if (item) {
                bindTooltipToCard(card, item);
                card.addEventListener('click', (e) => {
                    e.stopPropagation();
                    showItemDetail(item);
                    // 高亮左侧对应的卡片（可选）
                    const leftCard = document.querySelector(`.list-card[data-index="${allItemsCache.indexOf(item)}"]`);
                    if (leftCard) {
                        document.querySelectorAll('.list-card').forEach(c => c.classList.remove('active'));
                        leftCard.classList.add('active');
                    }
                });
            }
        });
    }
}
// ==================== 文件删除列表管理 ====================
async function refreshFileLists() {
    // 物品文件列表
    const itemContainer = document.getElementById('fileListContainer');
    if (itemContainer) {
        try {
            const files = await getAllFiles(ITEM_STORE);
            if (!files.length) {
                itemContainer.innerHTML = '<div class="file-list-empty">暂无物品文件</div>';
            } else {
                files.sort((a,b) => b.timestamp - a.timestamp);
                let html = '';
                for (let f of files) {
                    const date = new Date(f.timestamp).toLocaleString();
                    const sizeKB = (f.size / 1024).toFixed(1);
                    html += `
                        <div class="file-item" data-id="${f.id}">
                            <div class="file-info" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</div>
                            <button class="btn-delete" data-id="${f.id}" data-store="item">🗑️</button>
                        </div>
                    `;
                }
                itemContainer.innerHTML = html;
                document.querySelectorAll('#fileListContainer .btn-delete').forEach(btn => {
                    btn.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        const id = Number(btn.getAttribute('data-id'));
                        if (confirm('删除物品文件后，其中的物品将从列表中移除，确定？')) {
                            await deleteFileById(ITEM_STORE, id);
                            await refreshFileLists();
                            await refreshAllItems();
                            updateStatus('已删除物品文件');
                        }
                    });
                });
            }
        } catch (err) {
            itemContainer.innerHTML = '<div class="file-list-empty">加载失败</div>';
        }
    }
    
    // 地域文件列表
    const regionContainer = document.getElementById('regionFileDeleteContainer');
    if (regionContainer) {
        try {
            const files = await getAllFiles(REGION_STORE);
            if (!files.length) {
                regionContainer.innerHTML = '<div class="file-list-empty">暂无地域文件</div>';
            } else {
                files.sort((a,b) => b.timestamp - a.timestamp);
                let html = '';
                for (let f of files) {
                    const date = new Date(f.timestamp).toLocaleString();
                    const sizeKB = (f.size / 1024).toFixed(1);
                    html += `
                        <div class="file-item" data-id="${f.id}">
                            <div class="file-info" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</div>
                            <button class="btn-delete" data-id="${f.id}" data-store="region">🗑️</button>
                        </div>
                    `;
                }
                regionContainer.innerHTML = html;
                document.querySelectorAll('#regionFileDeleteContainer .btn-delete').forEach(btn => {
                    btn.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        const id = Number(btn.getAttribute('data-id'));
                        if (confirm('删除地域文件后，其中的地名将从列表中移除，确定？')) {
                            await deleteFileById(REGION_STORE, id);
                            await refreshFileLists();
                            await refreshAllRegions();
                            updateStatus('已删除地域文件');
                        }
                    });
                });
            }
        } catch (err) {
            regionContainer.innerHTML = '<div class="file-list-empty">加载失败</div>';
        }
    }
}

// ==================== 上传处理 ====================
async function handleItemUpload(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'xlsm', 'csv'].includes(ext)) {
        alert('请上传 Excel 文件');
        return;
    }
    updateStatus(`正在处理物品表 ${file.name} ...`);
    try {
        const blob = file.slice(0, file.size, file.type);
        await saveFileToDB(ITEM_STORE, file, blob);
        await refreshFileLists();
        await refreshAllItems();
        updateStatus(`已存储物品: ${file.name}`);
    } catch (err) {
        console.error(err);
        alert('上传失败: ' + err.message);
        updateStatus('上传失败');
    }
}

async function handleRegionUpload(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'xlsm', 'csv'].includes(ext)) {
        alert('请上传 Excel 文件');
        return;
    }
    updateStatus(`正在处理地域表 ${file.name} ...`);
    try {
        const blob = file.slice(0, file.size, file.type);
        await saveFileToDB(REGION_STORE, file, blob);
        await refreshFileLists();
        await refreshAllRegions();
        updateStatus(`已存储地域表: ${file.name}`);
    } catch (err) {
        console.error(err);
        alert('地域表解析失败: ' + err.message);
        updateStatus('上传失败');
    }
}

function updateStatus(msg) {
    const span = document.getElementById('statusText');
    if (span) span.innerText = msg;
}

// ==================== 初始化 ====================
(async function init() {
    await initDB();
    await refreshFileLists();
    await refreshAllItems();
    await refreshAllRegions();
    updateStatus('就绪，可上传物品表或地域表');
    
    // 绑定上传按钮
    const uploadItemBtn = document.getElementById('uploadExcelBtn');
    const uploadRegionBtn = document.getElementById('uploadRegionBtn');
    const itemInput = document.getElementById('excelInput');
    const regionInput = document.getElementById('regionInput');
    
    if (uploadItemBtn && itemInput) {
        uploadItemBtn.onclick = () => itemInput.click();
        itemInput.onchange = async (e) => {
            if (e.target.files.length) {
                await handleItemUpload(e.target.files[0]);
                itemInput.value = '';
            }
        };
    } else {
        console.error('未找到上传物品按钮或文件输入框');
    }
    
    if (uploadRegionBtn && regionInput) {
        uploadRegionBtn.onclick = () => regionInput.click();
        regionInput.onchange = async (e) => {
            if (e.target.files.length) {
                await handleRegionUpload(e.target.files[0]);
                regionInput.value = '';
            }
        };
    } else {
        console.error('未找到上传地域按钮或文件输入框');
    }
    
    const refreshBtn = document.getElementById('refreshFileListBtn');
    if (refreshBtn) refreshBtn.onclick = async () => {
        await refreshFileLists();
        await refreshAllItems();
        await refreshAllRegions();
        updateStatus('已刷新');
    };
})();
