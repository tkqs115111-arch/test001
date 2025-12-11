// =========================================================
//  PART 1: 設定與 API
// =========================================================
const SPREADSHEET_ID = '1tJjquBs-Wyav4VEg7XF-BnTAGoWhE-5RFwwhU16GuwQ'; 
const TARGET_SHEETS = ['Windows', 'RHEL', 'Oracle', 'ESXi', 'FW'];

let allProducts = [];
let groups = [ { id: 'g1', name: '群組 1', items: [] } ];
let activeGroupId = 'g1';
let currentView = 'search';

// =========================================================
//  PART 2: 初始化
// =========================================================
async function initData() {
    renderGroupsSidebar();
    
    const sheetsPromises = TARGET_SHEETS.map(name => fetchSheetData(name));
    let sheetsData = [];
    try { sheetsData = await Promise.all(sheetsPromises); } 
    catch (e) { document.getElementById('productContainer').innerHTML = '<div class="no-results">資料載入失敗。</div>'; return; }

    let aggregatedMap = {};
    sheetsData.forEach((sheet, index) => {
        if (!Array.isArray(sheet)) return;
        const currentSheetName = TARGET_SHEETS[index];
        const isFwSheet = (currentSheetName === 'FW');
        let lastComponent = '', lastVendor = '';    

        sheet.forEach(item => {
            const desc = item.description || item.Description || item['Model Name'];
            let rawComp = item.component || item.Component; 
            let rawVendor = item.vendor || item.Vendor;
            if (rawComp) lastComponent = rawComp; else rawComp = lastComponent;
            if (rawVendor) lastVendor = rawVendor; else rawVendor = lastVendor;

            if (!desc) return;
            const modelKey = desc.trim();

            if (!aggregatedMap[modelKey]) {
                aggregatedMap[modelKey] = {
                    id: item.swid || item.SWID || 'N/A', 
                    model: desc, 
                    brand: rawVendor || 'Generic',
                    type: rawComp || 'Component',    
                    fw: 'N/A',
                    drivers: [] 
                };
            }

            if (isFwSheet) {
                if (item['FW Version'] || item.FW) aggregatedMap[modelKey].fw = item['FW Version'] || item.FW;
                if (item.swid || item.SWID) aggregatedMap[modelKey].id = item.swid || item.SWID;
            } else {
                aggregatedMap[modelKey].drivers.push({
                    os: item.os || item.OS || currentSheetName, 
                    ver: item.driver || item.Driver || item.Version || 'N/A'
                });
            }
        });
    });

    allProducts = Object.values(aggregatedMap);
    document.getElementById('result-count').innerText = `載入完成，共 ${allProducts.length} 筆資料。`;
    renderSidebarMenu();
    renderProducts(allProducts, 'search');
}

async function fetchSheetData(sheetName) {
    try {
        const response = await fetch(`https://opensheet.elk.sh/${SPREADSHEET_ID}/${sheetName}`);
        return response.ok ? await response.json() : [];
    } catch (error) { return []; }
}

// =========================================================
//  PART 3: 渲染列表
// =========================================================
function renderProducts(data, viewType) {
    const container = document.getElementById('productContainer');
    container.innerHTML = '';

    if (data.length === 0) {
        container.innerHTML = '<div class="no-results">找不到資料</div>';
        return;
    }

    data.forEach((product) => {
        let displayDriver = "N/A";
        let displayOS = "OS";
        if (product.drivers.length > 0) {
            displayDriver = product.drivers[0].ver; 
            displayOS = product.drivers[0].os;
        }

        const currentGroup = groups.find(g => g.id === activeGroupId);
        const isAdded = currentGroup.items.some(i => i.model === product.model);
        
        let btnClass = isAdded ? 'active' : '';
        let btnIcon = isAdded ? '<i class="fas fa-check"></i>' : '<i class="fas fa-plus"></i>';
        let btnAction = `addToGroup('${product.model}')`;
        
        if (viewType === 'group') {
            btnClass = 'remove';
            btnIcon = '<i class="fas fa-minus"></i>';
            btnAction = `removeFromGroup('${product.model}')`;
        } else if (isAdded) {
            btnAction = ''; 
        }

        const mockCmd = generateCommand(product);

        const html = `
        <div class="hw-row-card">
            <div class="row-main-content">
                <div class="row-header">
                    <div class="row-model-title" title="${product.model}">${product.model}</div>
                    <div class="row-brand-badge">${product.brand}</div>
                </div>

                <div class="row-body">
                    <div class="data-group">
                        <div class="data-item">
                            <span class="data-label">FW Version</span>
                            <span class="data-val val-fw">${product.fw}</span>
                        </div>
                        <div class="data-item">
                            <span class="data-label">Driver (${displayOS})</span>
                            <span class="data-val val-driver">${displayDriver}</span>
                        </div>
                    </div>

                    <div class="action-group">
                        <div class="btn-expand-text" onclick="toggleDetails(this)">
                            詳細 <i class="fas fa-chevron-down"></i>
                        </div>
                        <div class="btn-circle ${btnClass}" onclick="${btnAction}" title="加入/移除">
                            ${btnIcon}
                        </div>
                    </div>
                </div>
            </div>

            <div class="row-details-panel">
                <div style="margin-bottom:5px; font-weight:bold; color:#666;">升刷指令:</div>
                <code class="cmd-code">${mockCmd}</code>
                <div style="margin-top:10px; font-size:12px; color:#999;">
                    SWID: <span style="font-family:monospace;">${product.id}</span>
                </div>
            </div>
        </div>`;
        
        container.innerHTML += html;
    });
}

function generateCommand(product) {
    const brand = product.brand.toLowerCase();
    const id = product.id !== 'N/A' ? product.id : 'DEVICE_ID';
    if (brand.includes('intel')) return `nvmupdate64e -l log.txt -c nvmupdate.cfg -id ${id}`;
    if (brand.includes('mellanox') || brand.includes('nvidia')) return `mstflint -d 00:03.0 -i ${id}.bin burn`;
    return `fw_update_tool --device "${product.model}" --firmware ${product.fw}.bin`;
}

function toggleDetails(btn) {
    const card = btn.closest('.hw-row-card');
    const panel = card.querySelector('.row-details-panel');
    const icon = btn.querySelector('i');
    
    if (panel.style.display === 'block') {
        panel.style.display = 'none';
        icon.classList.replace('fa-chevron-up', 'fa-chevron-down');
    } else {
        panel.style.display = 'block';
        icon.classList.replace('fa-chevron-down', 'fa-chevron-up');
    }
}

// =========================================================
//  PART 4: 功能邏輯 (包含導出)
// =========================================================

function renderSidebarMenu() {
    const menu = document.getElementById('sidebarMenu');
    menu.innerHTML = '';
    const components = [...new Set(allProducts.map(p => p.type))].filter(Boolean).sort();
    
    components.forEach(comp => {
        const vendors = [...new Set(allProducts.filter(p => p.type === comp).map(p => p.brand))].sort();
        let vendorHtml = '';
        
        vendors.forEach(vendor => {
            const models = allProducts.filter(p => p.type === comp && p.brand === vendor).map(p => p.model);
            const modelHtml = models.map(m => 
                `<li class="menu-item menu-model" onclick="filterByModel('${m}'); event.stopPropagation();">
                    ${m}
                 </li>`
            ).join('');

            vendorHtml += `
            <li>
                <div class="menu-item menu-vendor" onclick="toggleSubMenu(this)">
                    ${vendor} <i class="fas fa-caret-right arrow"></i>
                </div>
                <ul class="submenu">${modelHtml}</ul>
            </li>`;
        });

        menu.innerHTML += `
        <li>
            <div class="menu-item menu-category" onclick="toggleSubMenu(this)">
                ${comp} <i class="fas fa-caret-right arrow"></i>
            </div>
            <ul class="submenu">${vendorHtml}</ul>
        </li>`;
    });
}

function toggleSubMenu(el) {
    el.nextElementSibling.classList.toggle('open');
    el.parentElement.classList.toggle('open');
}

function renderGroupsSidebar() {
    const wrapper = document.getElementById('groups-wrapper');
    wrapper.innerHTML = '';

    groups.forEach(g => {
        const isActive = (g.id === activeGroupId);
        
        // ★ 修改處：這裡將 18 改成了 35，並增加了判斷長度的邏輯(只有超過才顯示...) ★
        let itemsHtml = g.items.length === 0 
            ? '<div style="color:#ccc;font-style:italic;padding:5px;text-align:center;">無卡片</div>' 
            : g.items.map(i => {
                // 設定最大顯示字數
                const maxLen = 35; 
                const displayName = i.model.length > maxLen ? i.model.substring(0, maxLen) + '...' : i.model;
                return `<div style="border-bottom:1px solid #eee;padding:2px;">${displayName}</div>`;
            }).join('');
        
        wrapper.innerHTML += `
        <div class="group-box ${isActive ? 'active' : ''}" onclick="setActiveGroup('${g.id}', event)">
            <div class="group-header">
                <input class="group-name-input" value="${g.name}" onchange="updateGroupName('${g.id}',this.value)" onclick="event.stopPropagation()">
                <i class="fas fa-trash-alt" style="color:#d93025;cursor:pointer;font-size:12px;" onclick="deleteGroup('${g.id}', event)" title="刪除"></i>
            </div>
            
            <div class="group-items-list">${itemsHtml}</div>
            
            <div class="group-actions">
                <i class="fas fa-file-export btn-icon btn-export-icon" title="導出 CSV" onclick="exportGroupToCSV('${g.id}', event)"></i>
                <i class="fas fa-eye btn-icon btn-view-icon" title="查看詳情" onclick="loadGroupView('${g.id}'); event.stopPropagation()"></i>
            </div>
        </div>`;
    });
    const activeGroup = groups.find(g => g.id === activeGroupId);
    if(activeGroup) document.getElementById('active-group-name').innerText = activeGroup.name;
}

// ★ 新增：導出 CSV 功能 ★
function exportGroupToCSV(gid, event) {
    event.stopPropagation(); // 防止切換群組
    const group = groups.find(g => g.id === gid);
    
    if (!group || group.items.length === 0) {
        alert("此群組沒有資料可導出！");
        return;
    }

    // CSV Header (加入 \uFEFF 解決 Excel 中文亂碼)
    let csvContent = "\uFEFF類別,廠牌,型號,FW版本,Driver版本(OS),SWID,升刷指令\n";

    group.items.forEach(item => {
        let drvInfo = "N/A";
        if(item.drivers.length > 0) {
            drvInfo = `${item.drivers[0].ver} (${item.drivers[0].os})`;
        }
        
        // 處理指令中的逗號，避免 CSV 格式跑掉
        const cmd = generateCommand(item).replace(/,/g, " ");

        csvContent += `${item.type},${item.brand},${item.model},${item.fw},${drvInfo},${item.id},${cmd}\n`;
    });

    // 建立下載連結
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${group.name}_export.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function createNewGroup() {
    groups.push({ id: 'g' + Date.now(), name: '新配置', items: [] });
    renderGroupsSidebar();
}

function setActiveGroup(gid, evt) {
    if(evt && evt.target.tagName === 'INPUT') return;
    activeGroupId = gid;
    renderGroupsSidebar();
    if(currentView === 'search') applyFilters();
}

function deleteGroup(gid, evt) {
    evt.stopPropagation();
    if(groups.length <= 1) return alert("至少保留一個");
    if(confirm("刪除?")) {
        groups = groups.filter(g => g.id !== gid);
        if(gid === activeGroupId) activeGroupId = groups[0].id;
        renderGroupsSidebar();
        if(currentView === 'search') applyFilters(); else loadGroupView(activeGroupId);
    }
}

function updateGroupName(gid, val) { groups.find(x => x.id === gid).name = val; document.getElementById('active-group-name').innerText = val; }

function addToGroup(name) {
    const g = groups.find(x => x.id === activeGroupId);
    const p = allProducts.find(x => x.model === name);
    if(p && !g.items.find(x => x.model === name)) { g.items.push(p); renderGroupsSidebar(); applyFilters(); }
}

function removeFromGroup(name) {
    const g = groups.find(x => x.id === activeGroupId);
    g.items = g.items.filter(x => x.model !== name);
    renderGroupsSidebar();
    renderProducts(g.items, 'group');
}

function loadGroupView(gid) {
    currentView = 'group'; activeGroupId = gid;
    renderGroupsSidebar();
    renderProducts(groups.find(g => g.id === gid).items, 'group');
}

function filterByModel(m) { document.getElementById('searchInput').value = m; applyFilters(); }

function applyFilters() {
    currentView = 'search';
    const kw = document.getElementById('searchInput').value.toLowerCase();
    renderProducts(allProducts.filter(p => p.model.toLowerCase().includes(kw) || p.brand.toLowerCase().includes(kw)), 'search');
}

function clearFilters() { document.getElementById('searchInput').value = ''; applyFilters(); }

document.getElementById("searchInput").addEventListener("keypress", e => { if(e.key==="Enter") applyFilters(); });

window.onload = initData;