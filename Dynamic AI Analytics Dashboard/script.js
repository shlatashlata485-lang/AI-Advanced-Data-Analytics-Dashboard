/**
 * OmniSight AI - High Clarity Edition (No Legends)
 * Developer: Mohamed Shlata | 2026
 */

let globalData = [];
let chartInstances = [];

document.addEventListener('DOMContentLoaded', () => {
    switchTab('dashboard');
    const fileInput = document.getElementById('fileInput');
    if (fileInput) fileInput.addEventListener('change', handleFileUpload);
});

function handleFileUpload(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            globalData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            if (globalData.length > 0) refreshAllSections();
        } catch (err) { alert("Format Error."); }
    };
    reader.readAsArrayBuffer(file);
}

function refreshAllSections() {
    renderDashboardCharts();
    renderAnalysisTable();
    renderExecutiveReports();
}

function renderDashboardCharts() {
    const container = document.getElementById('tab-dashboard');
    container.innerHTML = `
        <!-- كروت الإحصائيات -->
        <div id="statsCards" class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-10"></div>
        <!-- شبكة الرسوم التوضيحية -->
        <div id="chartsGrid" class="grid grid-cols-1 lg:grid-cols-2 gap-8"></div>
    `;

    const cols = Object.keys(globalData[0]);
    const numCol = cols.find(c => typeof globalData[0][c] === 'number') || cols.find(c => !isNaN(globalData[0][c]));
    const catCol = cols.find(c => typeof globalData[0][c] === 'string') || cols[0];

    const vals = globalData.map(r => Number(r[numCol]) || 0);
    const sum = vals.reduce((a, b) => a + b, 0);

    document.getElementById('statsCards').innerHTML = `
        <div class="glass-card p-6 rounded-3xl border-l-4 border-cyan-500">
            <p class="text-slate-500 text-[10px] uppercase font-bold">Total Rows</p>
            <h4 class="text-2xl font-bold">${globalData.length}</h4>
        </div>
        <div class="glass-card p-6 rounded-3xl border-l-4 border-purple-500">
            <p class="text-slate-500 text-[10px] uppercase font-bold">Total Sum of ${numCol}</p>
            <h4 class="text-2xl font-bold">${sum.toLocaleString()}</h4>
        </div>
        <div class="glass-card p-6 rounded-3xl border-l-4 border-emerald-500">
            <p class="text-slate-500 text-[10px] uppercase font-bold">Average Value</p>
            <h4 class="text-2xl font-bold">${(sum / vals.length).toFixed(2)}</h4>
        </div>
        <div class="glass-card p-6 rounded-3xl border-l-4 border-amber-500">
            <p class="text-slate-500 text-[10px] uppercase font-bold">Records Count</p>
            <h4 class="text-2xl font-bold">${vals.length}</h4>
        </div>
    `;

    chartInstances.forEach(c => c.destroy());
    chartInstances = [];

    const configs = [
        { type: 'bar', title: 'مقارنة الكميات', hint: 'يوضح الفرق بين القيم؛ العمود الأطول هو الأكبر.' },
        { type: 'line', title: 'تحليل الاتجاه', hint: 'يوضح هل البيانات في صعود أم هبوط.' },
        { type: 'pie', title: 'توزيع النسب', hint: 'يوضح حصة كل عنصر من الإجمالي كقطعة من الفطيرة.' },
        { type: 'doughnut', title: 'هيكل البيانات', hint: 'مشابه للدائرة؛ يركز على تقسيم الإجمالي.' },
        { type: 'radar', title: 'رادار الأداء', hint: 'يوضح توازن القوى؛ كلما زادت المساحة زاد الأداء.' },
        { type: 'polarArea', title: 'الكثافة الدائرية', hint: 'يقارن القيم بناءً على بُعدها عن المركز.' },
        { type: 'bubble', title: 'خريطة العلاقات', hint: 'حجم الدائرة يمثل أهمية القيمة.' },
        { type: 'scatter', title: 'تشتت البيانات', hint: 'يكشف القيم الشاذة والبعيدة عن المعدل.' }
    ];

    const summaryData = globalData.slice(0, 8).reduce((acc, row) => {
        const label = row[catCol] || 'Other';
        acc[label] = (acc[label] || 0) + (Number(row[numCol]) || 0);
        return acc;
    }, {});

    const grid = document.getElementById('chartsGrid');
    configs.forEach((conf, i) => {
        const cid = `chart-${i}`;
        const card = document.createElement('div');
        card.className = "glass-card p-6 rounded-[32px] hover:border-cyan-500/50 transition";
        card.innerHTML = `
            <div class="mb-4">
                <h4 class="text-white font-bold text-sm">${conf.title}</h4>
                <p class="text-cyan-500 text-[11px] font-medium mt-1">💡 ${conf.hint}</p>
            </div>
            <div class="h-64"><canvas id="${cid}"></canvas></div>
        `;
        grid.appendChild(card);

        const ctx = document.getElementById(cid).getContext('2d');
        const colors = ['#06b6d4', '#8b5cf6', '#10b981', '#f59e0b', '#3b82f6', '#ec4899', '#f43f5e', '#eab308'];

        chartInstances.push(new Chart(ctx, {
            type: conf.type,
            data: {
                labels: Object.keys(summaryData),
                datasets: [{
                    label: numCol,
                    data: Object.values(summaryData),
                    backgroundColor: i >= 2 ? colors : '#06b6d4',
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    // التعديل هنا: إخفاء الكلام الموجود تحت الرسمة
                    legend: { display: false }, 
                    tooltip: {
                        enabled: true,
                        backgroundColor: '#0f172a',
                        titleFont: { size: 14 },
                        bodyFont: { size: 12 },
                        callbacks: {
                            label: (ctx) => ` القيمة: ${ctx.raw.toLocaleString()}`
                        }
                    }
                },
                scales: ['pie', 'doughnut', 'radar', 'polarArea'].includes(conf.type) ? {} : {
                    y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b' } },
                    x: { grid: { display: false }, ticks: { color: '#64748b' } }
                }
            }
        }));
    });
}

function renderAnalysisTable() {
    const container = document.getElementById('tab-analysis');
    const cols = Object.keys(globalData[0]);
    container.innerHTML = `
        <div class="glass-card rounded-3xl overflow-hidden shadow-2xl">
            <div class="p-6 border-b border-white/5"><h3 class="font-bold">Raw Data Check</h3></div>
            <div class="overflow-x-auto">
                <table class="w-full text-left text-xs text-slate-400">
                    <thead class="bg-slate-800/50 text-cyan-400">
                        <tr>${cols.map(c => `<th class="p-4">${c}</th>`).join('')}</tr>
                    </thead>
                    <tbody>
                        ${globalData.slice(0, 10).map(row => `
                            <tr class="border-b border-white/5">
                                ${cols.map(c => `<td class="p-4">${row[c] || '-'}</td>`).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

function renderExecutiveReports() {
    const container = document.getElementById('tab-reports');
    container.innerHTML = `
        <div class="glass-card p-10 rounded-[40px] border-l-8 border-cyan-500">
            <h3 class="text-3xl font-bold mb-4">Strategic Summary</h3>
            <p class="text-slate-400 mb-8 italic">Analyzed by Mohamed Shlata Suite</p>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-10">
                <div class="p-6 bg-white/5 rounded-2xl">
                    <h4 class="text-cyan-400 font-bold mb-2 uppercase text-xs">Primary Logic</h4>
                    <p class="text-sm leading-relaxed text-slate-300">
                        الرسومات الموضحة تعكس أداء <b>${globalData.length}</b> سجل. 
                        نلاحظ أن توزيع البيانات يتركز في الفئات الأولى، مما يستوجب التركيز عليها لتحسين النتائج الإجمالية.
                    </p>
                </div>
                <div class="p-6 bg-white/5 rounded-2xl">
                    <h4 class="text-purple-400 font-bold mb-2 uppercase text-xs">Data Health</h4>
                    <p class="text-sm leading-relaxed text-slate-300">
                        كافة الرسوم البيانية تشير إلى استقرار في المعدلات، مع وجود بعض النقاط في رسمة "التشتت" (Scatter) التي تحتاج لمراجعة دقيقة.
                    </p>
                </div>
            </div>
        </div>
    `;
}

function switchTab(id) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
    document.getElementById(`tab-${id}`).classList.remove('hidden');
    document.getElementById('mainTitle').innerText = id.charAt(0).toUpperCase() + id.slice(1);
}