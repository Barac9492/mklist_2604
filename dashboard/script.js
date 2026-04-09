let allData = [];
let filteredData = [];
let currentSort = { column: 'score', asc: false };
let watchlist = JSON.parse(localStorage.getItem('mk_watchlist') || '[]');
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
    setupEventListeners();
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
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:30px; color:#adb5bd;">검색 결과가 없습니다.</td></tr>';
        return;
    }
    
    filteredData.forEach(item => {
        const tr = document.createElement('tr');
        
        const isHigh = item.score >= 35;
        const scoreClass = isHigh ? 'score-high' : 'score-med';
        const capDisplay = item.capital_million_krw ? `${item.capital_million_krw.toLocaleString()}M` : '-';
        
        const gradeBadge = `<span class="grade-badge grade-${item.investment_grade ? item.investment_grade.toLowerCase() : 'c'}">${item.investment_grade || 'C'}</span>`;
        
        let eliteBadge = '';
        if (item.talent_signals && item.talent_signals.length > 0) {
            const tooltipText = "사유: " + item.talent_signals.join(", ");
            eliteBadge = `<span class="badge-elite" title="${tooltipText}">🎖️ Elite</span>`;
        }

        const isStarred = watchlist.includes(item.name);
        const starBtn = `<button class="star-btn ${isStarred ? 'active' : ''}" onclick="toggleWatchlistItem('${item.name.replace(/'/g, "\\'")}')">${isStarred ? '★' : '☆'}</button>`;
        
        const intelLinks = `<span class="intel-links">
            <a href="https://www.linkedin.com/search/results/all/?keywords=${encodeURIComponent(item.name + ' ' + (item.ceo || ''))}" target="_blank" class="intel-btn li" title="LinkedIn 검색">in</a>
            <a href="https://search.naver.com/search.naver?query=${encodeURIComponent(item.name + ' 스타트업')}" target="_blank" class="intel-btn nv" title="Naver 검색">N</a>
        </span>`;
        
        let aiTag = '';
        if (item.llm_tag) {
            aiTag = `<br><span class="tag-ai" title="OpenAI 자동 분류">✨ ${item.llm_tag}</span>`;
        }
        
        tr.innerHTML = `
            <td>${starBtn}</td>
            <td>${gradeBadge}</td>
            <td><span class="badge-score ${scoreClass}">${item.score}</span></td>
            <td>
                <div class="name-wrapper">
                    <a href="https://www.google.com/search?q=${encodeURIComponent(item.name + ' 스타트업 홈페이지')}" target="_blank" class="rank-biz-name" title="Google 웹사이트 검색">${item.name}</a>
                    ${eliteBadge}
                    ${intelLinks}
                </div>
            </td>
            <td><span class="tag-category">${item.category || '-'}</span>${aiTag}</td>
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
    
    filteredData = allData.filter(item => {
        const matchSearch = item.name.toLowerCase().includes(searchStr) || item.business.toLowerCase().includes(searchStr);
        const matchCat = (catVal === 'all') || (item.category === catVal);
        const matchGrade = (gradeVal === 'all') || (item.investment_grade === gradeVal);
        const matchWatchlist = isWatchlistView ? watchlist.includes(item.name) : true;
        
        return matchSearch && matchCat && matchGrade && matchWatchlist;
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
    if (watchlist.includes(name)) {
        watchlist = watchlist.filter(n => n !== name);
    } else {
        watchlist.push(name);
    }
    localStorage.setItem('mk_watchlist', JSON.stringify(watchlist));
    
    if (isWatchlistView) {
        applyFilters(); 
    } else {
        renderTable(); 
    }
}

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

function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', applyFilters);
    document.getElementById('categoryFilter').addEventListener('change', applyFilters);
    if(document.getElementById('gradeFilter')) document.getElementById('gradeFilter').addEventListener('change', applyFilters);
    
    document.getElementById('toggleWatchlistBtn')?.addEventListener('click', (e) => {
        isWatchlistView = !isWatchlistView;
        e.currentTarget.classList.toggle('active', isWatchlistView);
        applyFilters();
    });
    
    document.getElementById('exportCsvBtn')?.addEventListener('click', exportToCSV);
    
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

// Init
document.addEventListener('DOMContentLoaded', loadData);
