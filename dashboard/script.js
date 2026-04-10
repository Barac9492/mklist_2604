let allData = [];
let filteredData = [];
let currentSort = { column: 'score', asc: false };

let rawWatchlist = JSON.parse(localStorage.getItem('mk_watchlist') || '{}');
let watchlist = {};
if (Array.isArray(rawWatchlist)) {
    rawWatchlist.forEach(name => {
        watchlist[name] = { status: 'Scouted', note: '' };
    });
    localStorage.setItem('mk_watchlist', JSON.stringify(watchlist));
} else {
    watchlist = rawWatchlist;
}

let isWatchlistView = false;

// Load data
async function loadData() {
    try {
        const response = await fetch('../data/startups_all_weeks.json');
        if (!response.ok) throw new Error('Data fetch failed');
        allData = await response.json();
        filteredData = [...allData];
        
        initDashboard();
    } catch (error) {
        console.error('Error loading data:', error);
        document.querySelector('.dashboard-content').innerHTML = `
            <div style="background: rgba(255,0,0,0.1); border: 1px solid rgba(255,0,0,0.3); padding:20px; border-radius:12px; color:#ff4d4d;">
                <h3>데이터 로드 오류</h3>
                <p>${error.message}</p>
                <p style="margin-top:10px; font-size:14px;">파일 프로토콜(file://)로 바로 열었을 경우 브라우저 보안 정책(CORS)에 의해 JSON 파일을 불러올 수 없습니다. 로컬 웹서버를 구동해주세요 (예: <code>python -m http.server</code>).</p>
            </div>`;
    }
}

function initDashboard() {
    updateKPIs();
    renderCharts();
    populateCategoryFilter();
    applyFilters(); // This will trigger sort, heatmap, and renderTable
    renderMomentum();
    setupEventListeners();
}

/* Helper: convert period string "M/D~M/D" to a sortable number */
function periodToNum(p) {
    // Extract start month/day from "M/D~M/D" format
    const m = p.match(/^(\d{1,2})\/(\d{1,2})/);
    return m ? parseInt(m[1]) * 100 + parseInt(m[2]) : 0;
}

function updateKPIs() {
    // Animate numbers up for premium feel
    animateValue('kpi-total', 0, allData.length, 1000);
    
    // Most popular category
    const countMap = {};
    allData.forEach(item => {
        if(item.category && item.category !== 'Other') {
            countMap[item.category] = (countMap[item.category] || 0) + 1;
        }
    });
    const topCat = Object.keys(countMap).sort((a, b) => countMap[b] - countMap[a])[0];
    document.getElementById('kpi-top-category').textContent = topCat || '-';
    
    // Weeks count
    const weeks = new Set(allData.map(item => item.period));
    animateValue('kpi-weeks-count', 0, weeks.size, 500);
    
    // Find and display latest period
    const sortedPeriods = [...weeks].sort((a, b) => periodToNum(b) - periodToNum(a));
    const latestPeriod = sortedPeriods[0];
    const latestCount = allData.filter(d => d.period === latestPeriod).length;
    const latestEl = document.getElementById('kpi-latest-period');
    if (latestEl) latestEl.textContent = `${latestPeriod} (${latestCount}개)`;
}

function animateValue(id, start, end, duration) {
    const obj = document.getElementById(id);
    let startTimestamp = null;
    const step = (timestamp) => {
        if (!startTimestamp) startTimestamp = timestamp;
        const progress = Math.min((timestamp - startTimestamp) / duration, 1);
        obj.innerHTML = Math.floor(progress * (end - start) + start);
        if (progress < 1) {
            window.requestAnimationFrame(step);
        } else {
            obj.innerHTML = end;
        }
    };
    window.requestAnimationFrame(step);
}

function renderCharts() {
    Chart.defaults.color = '#adb5bd';
    Chart.defaults.font.family = "'Inter', sans-serif";
    
    // 1. Trend Chart
    // Sort weeks chronologically by simply extracting the first part of date (assuming sorted naturally or handle it)
    const weeksSet = [...new Set(allData.map(item => item.period))];
    const trendCounts = weeksSet.map(week => allData.filter(i => i.period === week).length);
    
    const ctxTrend = document.getElementById('trendChart').getContext('2d');
    
    // Gradient for line chart
    const gradient = ctxTrend.createLinearGradient(0, 0, 0, 400);
    gradient.addColorStop(0, 'rgba(0, 210, 255, 0.4)');
    gradient.addColorStop(1, 'rgba(0, 210, 255, 0.0)');
    
    new Chart(ctxTrend, {
        type: 'line',
        data: {
            labels: weeksSet,
            datasets: [{
                label: '발굴된 스타트업 수',
                data: trendCounts,
                borderColor: '#00d2ff',
                backgroundColor: gradient,
                borderWidth: 3,
                pointBackgroundColor: '#0f111a',
                pointBorderColor: '#00d2ff',
                pointBorderWidth: 2,
                pointRadius: 4,
                pointHoverRadius: 6,
                fill: true,
                tension: 0.4 // Smooth curves
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(25, 28, 41, 0.9)',
                    titleColor: '#fff',
                    bodyColor: '#adb5bd',
                    borderColor: 'rgba(255,255,255,0.1)',
                    borderWidth: 1,
                    padding: 10,
                    displayColors: false
                }
            },
            scales: {
                y: { 
                    beginAtZero: true, 
                    grid: { color: 'rgba(255,255,255,0.05)', drawBorder: false }, 
                    ticks: { padding: 10 } 
                },
                x: { 
                    grid: { display: false, drawBorder: false }, 
                    ticks: { padding: 10 } 
                }
            },
            interaction: {
                intersect: false,
                mode: 'index',
            },
        }
    });

    // 2. Category Donut Chart
    const countMap = {};
    allData.forEach(item => {
        if(item.category !== 'Other' && item.category !== '') {
            countMap[item.category] = (countMap[item.category] || 0) + 1;
        }
    });
    
    const catLabels = Object.keys(countMap).sort((a,b) => countMap[b] - countMap[a]).slice(0, 5);
    const catData = catLabels.map(l => countMap[l]);
    
    const ctxCat = document.getElementById('categoryChart').getContext('2d');
    new Chart(ctxCat, {
        type: 'doughnut',
        data: {
            labels: catLabels,
            datasets: [{
                data: catData,
                backgroundColor: ['#00d2ff', '#3a86ff', '#8338ec', '#ff006e', '#fb5607'],
                borderWidth: 2,
                borderColor: '#0f111a',
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { 
                    position: 'right', 
                    labels: { 
                        color: '#adb5bd', 
                        font: {family: 'Inter', size: 12},
                        padding: 20,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    } 
                },
                tooltip: {
                    backgroundColor: 'rgba(25, 28, 41, 0.9)',
                    borderColor: 'rgba(255,255,255,0.1)',
                    borderWidth: 1,
                    padding: 12
                }
            },
            cutout: '75%'
        }
    });

    // 3. Region Horizontal Bar Chart (Top 5 excluding Seoul)
    const regionMap = {};
    allData.forEach(item => {
        if(item.region && item.region !== '서울' && item.region !== '') {
            regionMap[item.region] = (regionMap[item.region] || 0) + 1;
        }
    });
    
    const regLabels = Object.keys(regionMap).sort((a,b) => regionMap[b] - regionMap[a]).slice(0, 5);
    const regData = regLabels.map(l => regionMap[l]);
    
    const ctxReg = document.getElementById('regionChart').getContext('2d');
    new Chart(ctxReg, {
        type: 'bar',
        data: {
            labels: regLabels,
            datasets: [{
                data: regData,
                backgroundColor: 'rgba(58, 134, 255, 0.6)',
                borderColor: '#3a86ff',
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y', // makes it horizontal
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(25, 28, 41, 0.9)',
                    borderColor: 'rgba(255,255,255,0.1)',
                    borderWidth: 1,
                    padding: 12
                }
            },
            scales: {
                x: { 
                    beginAtZero: true, 
                    grid: { color: 'rgba(255,255,255,0.05)' } 
                },
                y: { 
                    grid: { display: false } 
                }
            }
        }
    });
}

function populateCategoryFilter() {
    const selector = document.getElementById('categoryFilter');
    const categories = [...new Set(allData.map(i => i.category))].filter(Boolean).sort();
    categories.forEach(cat => {
        const opt = document.createElement('option');
        opt.value = cat;
        opt.textContent = cat;
        selector.appendChild(opt);
    });
    
    // Populate region filter
    const regionSel = document.getElementById('regionFilter');
    if (regionSel) {
        const regions = [...new Set(allData.map(i => i.region))].filter(Boolean).sort();
        regions.forEach(r => {
            const opt = document.createElement('option');
            opt.value = r;
            opt.textContent = r;
            regionSel.appendChild(opt);
        });
    }
    
    // Populate period filter (sorted correctly)
    const periodSel = document.getElementById('periodFilter');
    if (periodSel) {
        const periods = [...new Set(allData.map(i => i.period))].sort((a, b) => periodToNum(b) - periodToNum(a));
        periods.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p;
            opt.textContent = p;
            periodSel.appendChild(opt);
        });
    }
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:30px; color:#adb5bd;">검색 결과가 없습니다.</td></tr>';
        return;
    }
    
    // Determine latest period for "NEW" badge
    const allPeriods = [...new Set(allData.map(d => d.period))].sort((a, b) => periodToNum(b) - periodToNum(a));
    const latestPeriod = allPeriods[0];
    
    filteredData.forEach(item => {
        const tr = document.createElement('tr');
        
        const isHigh = item.score >= 35;
        const scoreClass = isHigh ? 'score-high' : 'score-med';
        const capDisplay = item.capital_million_krw ? `${item.capital_million_krw.toLocaleString()}M` : '-';
        
        const gradeBadge = `<span class="grade-badge grade-${item.investment_grade ? item.investment_grade.toLowerCase() : 'c'}">${item.investment_grade || 'C'}</span>`;
        
        // NEW badge for the most recent week
        const isNew = item.period === latestPeriod;
        const newBadge = isNew ? `<span style="background:linear-gradient(135deg,#00d2ff,#3a86ff); color:#fff; font-size:9px; font-weight:800; padding:2px 5px; border-radius:3px; margin-left:6px; vertical-align:middle; letter-spacing:0.5px;">NEW</span>` : '';
        
        let eliteBadge = '';
        if (item.talent_signals && item.talent_signals.length > 0) {
            const tooltipText = "사유: " + item.talent_signals.join(", ");
            eliteBadge = `<span class="badge-elite" title="${tooltipText}">🎖️ Elite</span>`;
        }

        const isStarred = item.name in watchlist;
        const safeName = item.name.replace(/'/g, "\\'");
        const starBtn = `<button class="star-btn ${isStarred ? 'active' : ''}" onclick="toggleWatchlistItem('${safeName}')">${isStarred ? '★' : '☆'}</button>`;
        
        let crmControls = '';
        if (isStarred) {
            const wData = watchlist[item.name];
            crmControls = `
                <div class="crm-controls" style="margin-top:6px; background:rgba(0,0,0,0.2); padding:6px; border-radius:6px; font-size:12px; display:flex; gap:6px; align-items:center;">
                    <select onchange="updateWatchlistStatus('${safeName}', this.value)" style="background:#1a1d2e; color:#fff; border:1px solid #333; border-radius:4px; padding:2px 4px; font-size:11px;">
                        <option value="Scouted" ${wData.status === 'Scouted' ? 'selected' : ''}>🔍 Scouted</option>
                        <option value="Contacted" ${wData.status === 'Contacted' ? 'selected' : ''}>✉️ Contacted</option>
                        <option value="Meeting Set" ${wData.status === 'Meeting Set' ? 'selected' : ''}>🤝 Meeting Set</option>
                        <option value="Passed" ${wData.status === 'Passed' ? 'selected' : ''}>❌ Passed</option>
                    </select>
                    <button onclick="editWatchlistNote('${safeName}')" style="background:#1a1d2e; color:#00d2ff; border:1px solid #333; border-radius:4px; padding:2px 6px; cursor:pointer; font-size:11px; outline:none;">📝 Memo</button>
                    ${item.outreach_draft ? `<button onclick="showEmailDraft('${encodeURIComponent(item.outreach_draft)}')" style="background:#1a1d2e; color:#f1a5ff; border:1px solid #333; border-radius:4px; padding:2px 6px; cursor:pointer; font-size:11px; outline:none;">✉️ Email</button>` : ''}
                    ${item.lp_teaser ? `<button onclick="showLPTeaser('${encodeURIComponent(item.lp_teaser)}')" style="background:#1a1d2e; color:#ffdd57; border:1px solid #333; border-radius:4px; padding:2px 6px; cursor:pointer; font-size:11px; outline:none;">🌐 LP (EN)</button>` : ''}
                    <button onclick="showLookalikes('${safeName}')" style="background:#1a1d2e; color:#fff; border:1px solid #333; border-radius:4px; padding:2px 6px; cursor:pointer; font-size:11px; outline:none;">👯 Twins</button>
                    <span style="color:#adb5bd; font-family:serif; font-style:italic; max-width:150px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${wData.note}">${wData.note || 'No notes...'}</span>
                </div>
            `;
        }
        
        const intelLinks = `<span class="intel-links">
            <a href="https://www.linkedin.com/search/results/all/?keywords=${encodeURIComponent(item.name + ' ' + (item.ceo || ''))}" target="_blank" class="intel-btn li" title="LinkedIn 검색">in</a>
            <a href="https://search.naver.com/search.naver?query=${encodeURIComponent(item.name + ' 스타트업')}" target="_blank" class="intel-btn nv" title="Naver 검색">N</a>
            <a href="http://kpat.kipris.or.kr/kpat/searchLogina.do?next=MainSearch#page1?keyword=${encodeURIComponent(item.name)}" onclick="navigator.clipboard.writeText('${item.name}')" target="_blank" class="intel-btn kp" title="특허 검색 (클릭 시 이름 복사됨)" style="background:#4CAF50; color:#fff; border-color:#388E3C;">⚖️</a>
            <a href="https://www.google.com/search?q=${encodeURIComponent('"' + item.ceo + '" AND ("횡령" OR "배임" OR "파산" OR "임금체불" OR "사기")')}" target="_blank" class="intel-btn red" title="Red-Flag 창업자 검증" style="background:transparent; border-color:#ff4444; color:#ff4444;">🚨</a>
        </span>`;
        
        let aiTag = '';
        if (item.llm_tag) {
            aiTag = `<br><span class="tag-ai" title="OpenAI 자동 분류">✨ ${item.llm_tag}</span>`;
        }
        
        let osintTag = '';
        if (item.osint_snippet) {
            const safeSnippet = item.osint_snippet.replace(/"/g, '&quot;');
            osintTag = `<br><span class="tag-osint" title="${safeSnippet}">🔎 Intel: ${item.osint_snippet.substring(0, 55)}...</span>`;
        }
        
        tr.innerHTML = `
            <td>${starBtn}</td>
            <td>${gradeBadge}</td>
            <td><span class="badge-score ${scoreClass}">${item.score}</span></td>
            <td>
                <div class="name-wrapper">
                    <a href="https://www.google.com/search?q=${encodeURIComponent(item.name + ' 스타트업 홈페이지')}" target="_blank" class="rank-biz-name" title="Google 웹사이트 검색">${item.name}</a>
                    ${newBadge}
                    ${eliteBadge}
                    ${intelLinks}
                </div>
                ${crmControls}
            </td>
            <td><span class="tag-category">${item.category || '-'}</span>${aiTag}${osintTag}</td>
            <td style="color:#adb5bd">${capDisplay}</td>
            <td><span class="biz-desc" title="${item.business}">${item.business}</span></td>
            <td style="color:#adb5bd;font-size:13px;">${item.period}</td>
        `;
        tbody.appendChild(tr);
    });
}

function applyFilters() {
    const searchStr = document.getElementById('searchInput').value.toLowerCase();
    const catVal = document.getElementById('categoryFilter').value;
    const gradeVal = document.getElementById('gradeFilter') ? document.getElementById('gradeFilter').value : 'all';
    const regionVal = document.getElementById('regionFilter') ? document.getElementById('regionFilter').value : 'all';
    const periodVal = document.getElementById('periodFilter') ? document.getElementById('periodFilter').value : 'all';
    
    filteredData = allData.filter(item => {
        const matchSearch = item.name.toLowerCase().includes(searchStr) || item.business.toLowerCase().includes(searchStr) || (item.ceo || '').toLowerCase().includes(searchStr);
        const matchCat = (catVal === 'all') || (item.category === catVal);
        const matchGrade = (gradeVal === 'all') || (item.investment_grade === gradeVal);
        const matchRegion = (regionVal === 'all') || (item.region === regionVal);
        const matchPeriod = (periodVal === 'all') || (item.period === periodVal);
        const matchWatchlist = isWatchlistView ? (item.name in watchlist) : true;
        
        return matchSearch && matchCat && matchGrade && matchRegion && matchPeriod && matchWatchlist;
    });
    
    sortData();
    renderHeatmap();
}

function sortData() {
    const { column, asc } = currentSort;
    
    filteredData.sort((a, b) => {
        let valA = a[column];
        let valB = b[column];

        if (column === 'grade_val') {
            valA = a.score;
            valB = b.score;
        }
        
        if (typeof valA === 'string') valA = valA.toLowerCase();
        if (typeof valB === 'string') valB = valB.toLowerCase();
        
        // Handle nulls/missing for numbers
        if (column === 'capital_million_krw') {
            valA = valA || 0;
            valB = valB || 0;
        }

        if (valA < valB) return asc ? -1 : 1;
        if (valA > valB) return asc ? 1 : -1;
        return 0;
    });
    
    renderTable();
}

// CRM Logic Functions
function toggleWatchlistItem(name) {
    if (name in watchlist) {
        delete watchlist[name];
    } else {
        watchlist[name] = { status: 'Scouted', note: '' };
    }
    localStorage.setItem('mk_watchlist', JSON.stringify(watchlist));
    
    if (isWatchlistView) {
        applyFilters(); 
    } else {
        renderTable(); 
    }
}

function updateWatchlistStatus(name, status) {
    if (name in watchlist) {
        watchlist[name].status = status;
        localStorage.setItem('mk_watchlist', JSON.stringify(watchlist));
    }
}

function editWatchlistNote(name) {
    if (name in watchlist) {
        const currentNote = watchlist[name].note;
        const newNote = prompt(`"${name}"에 대한 메모를 입력하세요:`, currentNote);
        if (newNote !== null) {
            watchlist[name].note = newNote;
            localStorage.setItem('mk_watchlist', JSON.stringify(watchlist));
            renderTable();
        }
    }
}

function generateWeeklyMemo() {
    if (allData.length === 0) return alert("데이터가 없습니다.");
    // Find the latest period string by sorting descending and picking the first
    const periods = [...new Set(allData.map(d => d.period))].sort((a,b) => periodToNum(b) - periodToNum(a));
    const latestPeriod = periods[0];
    
    const topPicks = allData.filter(d => d.period === latestPeriod && d.score >= 35).sort((a,b) => b.score - a.score);
    
    if (topPicks.length === 0) {
        return alert(`${latestPeriod} 주차에는 점수 35점 이상(S/A급) 후보가 없습니다.`);
    }

    let memo = `🎯 주간 신규 설립 스타트업 요약 리포트 (${latestPeriod})\n`;
    memo += `총 ${topPicks.length}개의 주목할만한 (S/A급) 타겟이 발견되었습니다.\n\n`;

    topPicks.forEach((item, idx) => {
        const grade = item.score >= 45 ? 'S' : 'A';
        memo += `${idx + 1}. ${item.name} [${grade}급 - Score: ${item.score}]\n`;
        memo += `   - 카테고리: ${item.category || '-'}${item.llm_tag ? ` (✨ ${item.llm_tag})` : ''}\n`;
        memo += `   - 대표자: ${item.ceo}\n`;
        if (item.talent_signals && item.talent_signals.length > 0) {
            memo += `   - 주요 시그널: ${item.talent_signals.join(', ')}\n`;
        }
        memo += `   - 비즈니스 모델: ${item.business}\n\n`;
    });

    memo += `---\n이 리포트는 MK Scanner 엔진에서 자동으로 추출되었습니다.`;

    // Try to copy to clipboard
    navigator.clipboard.writeText(memo).then(() => {
        alert(`주간 리포트가 클립보드에 복사되었습니다! 슬랙이나 이메일에 붙여넣기 하세요.\n\n[미리보기]\n${memo.substring(0, 300)}...`);
    }).catch(err => {
        alert("클립보드 복사에 실패했습니다. 아래 텍스트를 복사하세요:\n\n" + memo);
    });
}

function showEmailDraft(encodedDraft) {
    const draft = decodeURIComponent(encodedDraft);
    const modalTitle = document.querySelector('#emailModal h3');
    modalTitle.textContent = "✉️ Cold-Start 파이프라인 (아웃바운드 초안)";
    
    document.getElementById('emailDraftContent').value = draft;
    document.getElementById('emailDraftContent').style.display = 'block';
    
    const visualDiv = document.getElementById('twinVisualDiv');
    if (visualDiv) visualDiv.style.display = 'none';
    
    const copyBtn = document.getElementById('copyEmailBtn');
    if (copyBtn) copyBtn.style.display = 'inline-block';
    
    document.getElementById('emailModal').style.display = 'flex';
}

function showLPTeaser(encodedDraft) {
    const draft = decodeURIComponent(encodedDraft);
    const modalTitle = document.querySelector('#emailModal h3');
    modalTitle.textContent = "🌐 Global LP Syndication (Teaser)";
    
    document.getElementById('emailDraftContent').value = draft;
    document.getElementById('emailDraftContent').style.display = 'block';
    
    const visualDiv = document.getElementById('twinVisualDiv');
    if (visualDiv) visualDiv.style.display = 'none';
    
    const copyBtn = document.getElementById('copyEmailBtn');
    if (copyBtn) copyBtn.style.display = 'inline-block';
    
    document.getElementById('emailModal').style.display = 'flex';
}

document.getElementById('closeEmailModal')?.addEventListener('click', () => {
    document.getElementById('emailModal').style.display = 'none';
});

document.getElementById('copyEmailBtn')?.addEventListener('click', () => {
    const text = document.getElementById('emailDraftContent').value;
    navigator.clipboard.writeText(text).then(() => {
        alert("내용이 클립보드에 복사되었습니다!");
    }).catch(err => {
        alert("복사 실패. 직접 영역을 지정하여 복사해주세요.");
    });
});

function exportToCSV() {
    if (filteredData.length === 0) return alert("데이터가 없습니다.");
    
    const headers = ["Grade", "Score", "Name", "Category", "Capital(M)", "CEO", "Business", "Region", "Period", "TalentSignals"];
    const rows = filteredData.map(item => [
        item.investment_grade || 'C',
        item.score,
        `"${item.name.replace(/"/g, '""')}"`,
        item.category,
        item.capital_million_krw || 0,
        `"${(item.ceo || '').replace(/"/g, '""')}"`,
        `"${item.business.replace(/"/g, '""')}"`,
        item.region,
        item.period,
        `"${(item.talent_signals || []).join(", ").replace(/"/g, '""')}"`
    ]);
    
    const csvContent = [headers.join(","), ...rows.map(e => e.join(","))].join("\n");
    const blob = new Blob(["\uFEFF"+csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `MK_Scanner_Export_${new Date().toISOString().slice(0,10)}.csv`;
    link.click();
}

function renderHeatmap() {
    const container = document.getElementById('heatmapContainer');
    if (!container) return;
    
    const ignoreWords = ['및', '개발', '공급', '제조', '판매', '기반', '소프트웨어', '시스템', '서비스', '서비스업', '개발업', '공급업', '제공업', '제조업', '도소매업', '관련', '솔루션', '컨설팅업', '플랫폼', '운영업', '전문'];
    const wordCounts = {};
    
    filteredData.forEach(item => {
        const words = item.business.split(/[\s,()/\·]+/);
        words.forEach(w => {
            if (w.length > 1 && !ignoreWords.includes(w)) {
                wordCounts[w] = (wordCounts[w] || 0) + 1;
            }
        });
    });
    
    const sortedWords = Object.keys(wordCounts).map(w => ({word: w, count: wordCounts[w]}))
        .sort((a, b) => b.count - a.count).slice(0, 15);
    
    container.innerHTML = '';
    sortedWords.forEach((item, index) => {
        let level = 0;
        if (index < 3) level = 3;
        else if (index < 7) level = 2;
        else if (index < 12) level = 1;
        
        const tag = document.createElement('span');
        tag.className = `heatmap-tag heat-level-${level}`;
        tag.textContent = `${item.word} (${item.count})`;
        container.appendChild(tag);
    });
}

function renderMomentum() {
    const momentumContainer = document.getElementById('momentumContent');
    if (!momentumContainer || allData.length === 0) return;

    const periods = [...new Set(allData.map(d => d.period))].sort((a,b) => periodToNum(b) - periodToNum(a));
    if (periods.length < 2) {
        momentumContainer.innerHTML = "데이터가 부족하여 주간 비교를 수행할 수 없습니다.";
        return;
    }

    const currentPeriod = periods[0];
    const previousPeriod = periods[1];

    const countWords = (dataArray) => {
        const ignoreWords = ['및', '개발', '공급', '제조', '판매', '기반', '소프트웨어', '시스템', '서비스', '서비스업', '개발업', '공급업', '제공업', '제조업', '도소매업', '관련', '솔루션', '컨설팅업', '플랫폼', '운영업', '전문'];
        const wordCounts = {};
        dataArray.forEach(item => {
            const words = item.business.split(/[\s,()/\·]+/);
            words.forEach(w => {
                if (w.length > 1 && !ignoreWords.includes(w)) {
                    wordCounts[w] = (wordCounts[w] || 0) + 1;
                }
            });
        });
        return wordCounts;
    };

    const currentCounts = countWords(allData.filter(d => d.period === currentPeriod));
    const previousCounts = countWords(allData.filter(d => d.period === previousPeriod));

    const momentumTarget = [];
    for (const [word, curCount] of Object.entries(currentCounts)) {
        if (curCount >= 2) { 
            const prevCount = previousCounts[word] || 0;
            const diff = curCount - prevCount;
            if (diff > 0) {
                momentumTarget.push({ word, diff, curCount, prevCount });
            }
        }
    }
    
    momentumTarget.sort((a,b) => b.diff - a.diff);
    const topMomentum = momentumTarget.slice(0, 5);

    if (topMomentum.length === 0) {
        momentumContainer.innerHTML = "이전 주차 대비 급상승한 키워드가 없습니다.";
        return;
    }

    let html = `<div style="margin-bottom:8px; color:#adb5bd;">(${previousPeriod} ➔ ${currentPeriod} 비교)</div>`;
    topMomentum.forEach(m => {
        html += `<div style="display:flex; justify-content:space-between; margin-bottom:6px; border-bottom:1px solid rgba(255,255,255,0.05); padding-bottom:4px;">
            <strong style="color:#fff;">${m.word}</strong>
            <span style="color:#00d2ff; font-weight:bold;">+${m.diff} 건 <span style="font-size:10px; color:#666; font-weight:normal;">(총 ${m.curCount}건)</span></span>
        </div>`;
    });
    momentumContainer.innerHTML = html;
}

function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', applyFilters);
    document.getElementById('categoryFilter').addEventListener('change', applyFilters);
    if(document.getElementById('gradeFilter')) document.getElementById('gradeFilter').addEventListener('change', applyFilters);
    if(document.getElementById('regionFilter')) document.getElementById('regionFilter').addEventListener('change', applyFilters);
    if(document.getElementById('periodFilter')) document.getElementById('periodFilter').addEventListener('change', applyFilters);
    
    document.getElementById('toggleWatchlistBtn')?.addEventListener('click', (e) => {
        isWatchlistView = !isWatchlistView;
        e.currentTarget.classList.toggle('active', isWatchlistView);
        applyFilters();
    });
    
    document.getElementById('exportCsvBtn')?.addEventListener('click', exportToCSV);
    const generateMemoBtn = document.getElementById('generateMemoBtn');
    if (generateMemoBtn) generateMemoBtn.addEventListener('click', generateWeeklyMemo);
    
    document.querySelectorAll('.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const column = th.dataset.sort;
            if (currentSort.column === column) {
                currentSort.asc = !currentSort.asc; // Toggle sort
            } else {
                currentSort.column = column;
                currentSort.asc = column === 'name'; // Names asc by default, others desc
            }
            sortData();
            
            // visually update header
            document.querySelectorAll('.sortable').forEach(el => {
                el.style.color = '';
                el.innerHTML = el.innerHTML.replace(' ↑', ' ↕').replace(' ↓', ' ↕');
            });
            th.style.color = '#00d2ff';
            const icon = currentSort.asc ? ' ↑' : ' ↓';
            th.innerHTML = th.innerHTML.replace(' ↕', icon);
        });
    });
    
    // Setup initial sort state visual
    const initialTh = document.querySelector('th[data-sort="score"]');
    if (initialTh) {
        initialTh.style.color = '#00d2ff';
        initialTh.innerHTML = initialTh.innerHTML.replace(' ↕', ' ↓');
    }
    sortData();
}

// ── TF-IDF Lookalike Engine ──────────────────────────────────────────────────

function buildTfIdfVectors(data) {
    const STOPWORDS = new Set(['및', '개발', '공급', '제조', '판매', '기반', '소프트웨어',
        '시스템', '서비스', '서비스업', '개발업', '공급업', '제공업', '제조업', '도소매업',
        '관련', '솔루션', '컨설팅업', '플랫폼', '운영업', '전문', '의', '에', '을', '를',
        '이', '가', '은', '는', '와', '과', '도', '로', '으로', '에서', '에게', '까지']);

    // Term frequency for each doc
    const tokenize = (text) =>
        text.toLowerCase().split(/[\s,()/·\-]+/).filter(w => w.length > 1 && !STOPWORDS.has(w));

    const allTokens = data.map(item => tokenize(item.business + ' ' + (item.category || '') + ' ' + (item.llm_tag || '')));

    // Count DF (document frequency)
    const df = {};
    allTokens.forEach(tokens => {
        new Set(tokens).forEach(t => { df[t] = (df[t] || 0) + 1; });
    });

    const N = data.length;
    // Build TF-IDF vectors
    return allTokens.map(tokens => {
        const tf = {};
        tokens.forEach(t => { tf[t] = (tf[t] || 0) + 1; });
        const vec = {};
        const len = tokens.length || 1;
        Object.entries(tf).forEach(([t, cnt]) => {
            const idf = Math.log((N + 1) / ((df[t] || 0) + 1)) + 1;
            vec[t] = (cnt / len) * idf;
        });
        return vec;
    });
}

function cosineSimilarity(vecA, vecB) {
    const keysA = Object.keys(vecA);
    let dot = 0, normA = 0, normB = 0;
    keysA.forEach(k => {
        dot += (vecA[k] || 0) * (vecB[k] || 0);
        normA += vecA[k] ** 2;
    });
    Object.values(vecB).forEach(v => { normB += v ** 2; });
    return normA && normB ? dot / (Math.sqrt(normA) * Math.sqrt(normB)) : 0;
}

let _tfidfVectors = null;

function showLookalikes(targetName) {
    // Lazy-build vectors once
    if (!_tfidfVectors) {
        _tfidfVectors = buildTfIdfVectors(allData);
    }

    const targetIdx = allData.findIndex(d => d.name === targetName);
    if (targetIdx === -1) return alert('해당 스타트업을 찾을 수 없습니다.');
    const targetVec = _tfidfVectors[targetIdx];
    const target = allData[targetIdx];

    // Score all others
    const scores = allData
        .map((item, idx) => ({ item, sim: idx === targetIdx ? -1 : cosineSimilarity(targetVec, _tfidfVectors[idx]) }))
        .filter(x => x.sim > 0)
        .sort((a, b) => b.sim - a.sim)
        .slice(0, 3);

    // Build modal content
    const gradeCls = g => `grade-${(g || 'b').toLowerCase()}`;
    const simBar = (sim) => {
        const pct = Math.round(sim * 100);
        const color = pct > 70 ? '#00d2ff' : pct > 45 ? '#8338ec' : '#adb5bd';
        return `<div style="height:4px; background:rgba(255,255,255,0.1); border-radius:2px; margin-top:6px;">
            <div style="height:4px; width:${pct}%; background:${color}; border-radius:2px; transition:width 0.6s ease;"></div>
        </div>
        <span style="font-size:10px; color:${color}; font-weight:700;">${pct}% match</span>`;
    };

    let html = `
        <div style="background:rgba(0,210,255,0.05); border:1px solid rgba(0,210,255,0.2); border-radius:10px; padding:12px; margin-bottom:16px;">
            <div style="font-size:11px; color:#adb5bd; text-transform:uppercase; letter-spacing:1px; margin-bottom:4px;">분석 대상</div>
            <div style="font-weight:700; font-size:15px;">${target.name}</div>
            <div style="font-size:12px; color:#adb5bd; margin-top:4px;">${target.business.substring(0, 80)}${target.business.length > 80 ? '...' : ''}</div>
        </div>
        <div style="font-size:11px; color:#adb5bd; text-transform:uppercase; letter-spacing:1px; margin-bottom:10px;">🔍 가장 유사한 스타트업 (TF-IDF 코사인 유사도)</div>
    `;

    if (scores.length === 0) {
        html += `<div style="color:#adb5bd; text-align:center; padding:20px;">유사한 기업을 찾지 못했습니다.</div>`;
    } else {
        scores.forEach((s, i) => {
            const medalEmojis = ['🥇', '🥈', '🥉'];
            const grade = s.item.investment_grade || 'B';
            html += `
                <div style="background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.08); border-radius:8px; padding:12px; margin-bottom:10px; transition:border-color 0.2s;" 
                     onmouseover="this.style.borderColor='rgba(0,210,255,0.3)'" 
                     onmouseout="this.style.borderColor='rgba(255,255,255,0.08)'">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div style="display:flex; align-items:center; gap:8px;">
                            <span style="font-size:18px;">${medalEmojis[i]}</span>
                            <div>
                                <span class="grade-badge ${gradeCls(grade)}" style="font-size:11px; padding:2px 6px;">${grade}</span>
                                <span style="font-weight:600; margin-left:6px;">${s.item.name}</span>
                            </div>
                        </div>
                        <span style="font-size:11px; color:#adb5bd;">${s.item.period}</span>
                    </div>
                    <div style="font-size:12px; color:#adb5bd; margin-top:6px; margin-left:26px;">${s.item.business.substring(0, 90)}${s.item.business.length > 90 ? '...' : ''}</div>
                    ${s.item.llm_tag ? `<div style="margin-top:4px; margin-left:26px;"><span class="tag-ai">✨ ${s.item.llm_tag}</span></div>` : ''}
                    <div style="margin-left:26px; margin-top:8px;">${simBar(s.sim)}</div>
                </div>
            `;
        });
    }

    // Use the shared emailModal but repurpose it for twins
    const modalTitle = document.querySelector('#emailModal h3');
    modalTitle.textContent = `👯 Startup Twins — "${target.name}"과 유사한 기업`;

    const textarea = document.getElementById('emailDraftContent');
    textarea.style.display = 'none';

    const copyBtn = document.getElementById('copyEmailBtn');
    if (copyBtn) copyBtn.style.display = 'none';

    let visualDiv = document.getElementById('twinVisualDiv');
    if (!visualDiv) {
        visualDiv = document.createElement('div');
        visualDiv.id = 'twinVisualDiv';
        textarea.parentNode.insertBefore(visualDiv, textarea);
    }
    visualDiv.style.display = 'block';
    visualDiv.innerHTML = html;

    document.getElementById('emailModal').style.display = 'flex';
}

// Init
document.addEventListener('DOMContentLoaded', loadData);
