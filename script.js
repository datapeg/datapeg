        const sheetId = '1yJKwg1NY9WddAYWSj56JVFATDSSbKWET';
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`;
        let fullData = [];
        let activeCharts = [];

        Chart.register(ChartDataLabels);

        function toRoman(text) {
            if (!text) return '-';
            return text.toString()
                .replace(/WILAYAH 1/gi, 'WILAYAH I')
                .replace(/WILAYAH 2/gi, 'WILAYAH II')
                .replace(/WILAYAH 3/gi, 'WILAYAH III');
        }

        function toggleSidebar() {
            document.getElementById('sidebar').classList.toggle('sidebar-hidden');
        }

        function showPage(pageId, element) {
            document.querySelectorAll('.nav-link').forEach(link => {
                link.classList.replace('bg-emerald-700', 'hover:bg-emerald-800');
                link.classList.replace('text-white', 'text-gray-300');
            });
            element.classList.replace('hover:bg-emerald-800', 'bg-emerald-700');
            element.classList.replace('text-gray-300', 'text-white');

            document.getElementById('page-dashboard').classList.remove('active');
            document.getElementById('page-data-container').classList.remove('active');

            if (pageId === 'dashboard') {
                document.getElementById('page-dashboard').classList.add('active');
                document.getElementById('page-title').innerText = "Dashboard";
            } else {
                document.getElementById('page-data-container').classList.add('active');
                let filtered = [];
                const search = (k) => fullData.filter(p => (p['ESELON 4'] || '').toString().toUpperCase().includes(k.toUpperCase()));

                if (pageId === 'all-data') {
                    document.getElementById('page-title').innerText = "Data Seluruh ASN";
                    filtered = fullData;
                    renderDataPage(filtered, 'all');
                } else {
                    if (pageId === 'subag-tu') {
                        document.getElementById('page-title').innerText = "Sub Bagian Tata Usaha";
                        filtered = search('TATA USAHA');
                    } else if (pageId === 'sekwil-1') {
                        document.getElementById('page-title').innerText = "Seksi Wilayah I";
                        filtered = search('WILAYAH 1');
                    } else if (pageId === 'sekwil-2') {
                        document.getElementById('page-title').innerText = "Seksi Wilayah II";
                        filtered = search('WILAYAH 2');
                    } else if (pageId === 'sekwil-3') {
                        document.getElementById('page-title').innerText = "Seksi Wilayah III";
                        filtered = search('WILAYAH 3');
                    }
                    renderDataPage(filtered, 'filtered');
                }
            }
        }

        function renderDataPage(data, viewType) {
            activeCharts.forEach(c => c.destroy());
            activeCharts = [];

            const clean = (v) => (v || "").toString().trim().toUpperCase();

            document.getElementById('card-total').innerText = data.length;
            document.getElementById('card-pns').innerText = data.filter(p => clean(p['STATUS ASN']) === 'PNS').length;
            document.getElementById('card-cpns').innerText = data.filter(p => clean(p['STATUS ASN']) === 'CPNS').length;
            document.getElementById('card-pppk').innerText = data.filter(p => clean(p['STATUS ASN']) === 'PPPK').length;

            const chartsArea = document.getElementById('charts-area');
            if (viewType === 'all') {
                chartsArea.className = "grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-4 gap-6 mb-8 uppercase font-bold";
                chartsArea.innerHTML = `
                    <div class="bg-white p-5 rounded-3xl shadow-sm border border-gray-100 flex flex-col items-center uppercase font-bold">
                        <h4 class="text-[10px] text-gray-600 mb-4 text-center font-bold uppercase uppercase font-bold">Data Berdasarkan Jenis Kelamin</h4>
                        <div class="h-64 w-full"><canvas id="chartGender"></canvas></div>
                    </div>
                    <div class="bg-white p-5 rounded-3xl shadow-sm border border-gray-100 flex flex-col items-center uppercase font-bold">
                        <h4 class="text-[10px] text-gray-600 mb-4 text-center font-bold uppercase uppercase font-bold">Data Berdasarkan Lokasi Kerja</h4>
                        <div class="h-64 w-full"><canvas id="chartLokasi"></canvas></div>
                    </div>
                    <div class="bg-white p-5 rounded-3xl shadow-sm border border-gray-100 flex flex-col items-center uppercase font-bold">
                        <h4 class="text-[10px] text-gray-600 mb-4 text-center font-bold uppercase uppercase font-bold">Data Berdasarkan Jenis Pegawai</h4>
                        <div class="h-64 w-full"><canvas id="chartASN"></canvas></div>
                    </div>
                    <div class="bg-white p-5 rounded-3xl shadow-sm border border-gray-100 flex flex-col items-center uppercase font-bold">
                        <h4 class="text-[10px] text-gray-600 mb-4 text-center font-bold uppercase uppercase font-bold">Data Berdasarkan Jenis Jabatan</h4>
                        <div class="h-64 w-full"><canvas id="chartJabatan"></canvas></div>
                    </div>
                `;
                initAllCharts(data);
            } else {
                chartsArea.className = "grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8 uppercase font-bold";
                chartsArea.innerHTML = `
                    <div class="bg-white p-6 rounded-[2rem] shadow-sm border border-gray-100 uppercase font-bold">
                        <h4 class="text-emerald-800 font-black mb-6 border-l-8 border-emerald-500 pl-4 text-xs tracking-[0.2em] uppercase uppercase font-bold">Gender</h4>
                        <div class="h-[250px] flex justify-center uppercase font-bold"><canvas id="chartGender"></canvas></div>
                    </div>
                    <div class="bg-white p-6 rounded-[2rem] shadow-sm border border-gray-100 uppercase font-bold">
                        <h4 class="text-emerald-800 font-black mb-6 border-l-8 border-emerald-500 pl-4 text-xs tracking-[0.2em] uppercase uppercase font-bold">Pendidikan</h4>
                        <div class="h-[250px] flex justify-center uppercase font-bold"><canvas id="chartPendidikan"></canvas></div>
                    </div>
                `;
                initFilteredCharts(data);
            }

            const tb = document.getElementById('mainTableBody');
            tb.innerHTML = '';
            data.forEach((p, i) => {
                const s = clean(p['STATUS ASN']);
                const color = s === 'PNS' ? 'bg-emerald-100 text-emerald-700' : s === 'PPPK' ? 'bg-blue-100 text-blue-700' : 'bg-amber-100 text-amber-700';
                tb.innerHTML += `<tr class="hover:bg-emerald-50 transition duration-150 uppercase font-bold">
                    <td class="px-6 py-5 text-center border-r font-bold">${i+1}</td>
                    <td class="px-6 py-4 uppercase font-bold">
                        <div class="text-gray-900 font-bold uppercase">${p['NAMA PEGAWAI']}</div>
                        <div class="text-[9px] text-gray-400 font-mono font-normal">NIP. ${p['NIP'] || '-'}</div>
                    </td>
                    <td class="px-6 py-4 uppercase font-bold leading-relaxed">
                        <div class="text-gray-700 text-[10px] uppercase font-bold">${p['JABATAN']}</div>
                        <div class="text-emerald-600 font-black text-[9px] uppercase">${toRoman(p['ESELON 4'])}</div>
                    </td>
                    <td class="px-6 py-4 text-center">
                        <span class="px-4 py-1.5 rounded-xl font-black text-[9px] shadow-sm uppercase ${color}">${p['STATUS ASN']}</span>
                    </td>
                </tr>`;
            });
        }

        function initAllCharts(data) {
            const clean = (v) => (v || "").toString().trim().toUpperCase();
            
            // 1. Gender Pie
            const lCount = data.filter(p => clean(p['GENDER']) === 'L').length;
            const pCount = data.filter(p => clean(p['GENDER']) === 'P').length;
            activeCharts.push(new Chart(document.getElementById('chartGender'), {
                type: 'doughnut',
                data: { labels: ['Laki-laki', 'Perempuan'], datasets: [{ data: [lCount, pCount], backgroundColor: ['#10b981', '#f97316'], borderWidth: 0 }] },
                options: getPieOptions()
            }));

            // 2. Lokasi Pie (4 Bagian: TU, Sekwil I, II, III)
            const countLok = (k) => data.filter(p => clean(p['ESELON 4']).includes(k)).length;
            activeCharts.push(new Chart(document.getElementById('chartLokasi'), {
                type: 'doughnut',
                data: { 
                    labels: ['TU', 'Sekwil I', 'Sekwil II', 'Sekwil III'], 
                    datasets: [{ data: [countLok('TATA USAHA'), countLok('WILAYAH 1'), countLok('WILAYAH 2'), countLok('WILAYAH 3')], backgroundColor: ['#10b981', '#f97316', '#3b82f6', '#fbbf24'], borderWidth: 0 }] 
                },
                options: getPieOptions()
            }));

            // 3. Jenis ASN Combo Bar
            const asnTypes = ['CPNS', 'PNS', 'PPPK'];
            const asnL = asnTypes.map(t => data.filter(p => clean(p['STATUS ASN']) === t && clean(p['GENDER']) === 'L').length);
            const asnP = asnTypes.map(t => data.filter(p => clean(p['STATUS ASN']) === t && clean(p['GENDER']) === 'P').length);
            activeCharts.push(new Chart(document.getElementById('chartASN'), {
                type: 'bar',
                data: {
                    labels: asnTypes,
                    datasets: [
                        { label: 'Laki-laki', data: asnL, backgroundColor: '#3b82f6' },
                        { label: 'Perempuan', data: asnP, backgroundColor: '#fbbf24' },
                        { label: 'Total', data: asnL.map((v, i) => v + asnP[i]), type: 'line', borderColor: '#3b82f6', tension: 0.4, pointBackgroundColor: '#fff' }
                    ]
                },
                options: getBarOptions()
            }));

            // 4. Jenis Jabatan Combo Bar
            const getJabType = (jab) => {
                const j = clean(jab);
                if (j.includes('KEPALA') || j.includes('KASUBAG')) return 'Struktural';
                if (j.includes('AHLI') || j.includes('MAHIR') || j.includes('TERAMPIL') || j.includes('PENYELIA')) return 'Fungsional';
                return 'Pelaksana';
            };
            const jTypes = ['Struktural', 'Fungsional', 'Pelaksana'];
            const jL = jTypes.map(t => data.filter(p => getJabType(p['JABATAN']) === t && clean(p['GENDER']) === 'L').length);
            const jP = jTypes.map(t => data.filter(p => getJabType(p['JABATAN']) === t && clean(p['GENDER']) === 'P').length);
            activeCharts.push(new Chart(document.getElementById('chartJabatan'), {
                type: 'bar',
                data: {
                    labels: jTypes,
                    datasets: [
                        { label: 'Laki-laki', data: jL, backgroundColor: '#3b82f6' },
                        { label: 'Perempuan', data: jP, backgroundColor: '#fbbf24' },
                        { label: 'Total', data: jL.map((v, i) => v + jP[i]), type: 'line', borderColor: '#3b82f6', tension: 0.4, pointBackgroundColor: '#fff' }
                    ]
                },
                options: getBarOptions()
            }));
        }

        function initFilteredCharts(data) {
            const clean = (v) => (v || "").toString().trim().toUpperCase();
            activeCharts.push(new Chart(document.getElementById('chartGender'), {
                type: 'doughnut',
                data: { labels: ['Laki-laki', 'Perempuan'], datasets: [{ data: [data.filter(p => clean(p['GENDER']) === 'L').length, data.filter(p => clean(p['GENDER']) === 'P').length], backgroundColor: ['#0284c7', '#db2777'], borderWidth: 0 }] },
                options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' }, datalabels: { color: '#fff', font: { weight: 'bold' } } } }
            }));
            const levels = ['S3', 'S2', 'S1', 'DIII', 'SLTA'];
            activeCharts.push(new Chart(document.getElementById('chartPendidikan'), {
                type: 'bar',
                data: { labels: levels, datasets: [{ label: 'Orang', data: levels.map(lvl => data.filter(p => clean(p['PENDIDIKAN']) === lvl).length), backgroundColor: '#10b981', borderRadius: 10 }] },
                options: { 
                    maintainAspectRatio: false, 
                    scales: { y: { beginAtZero: true, grid: { display: false } }, x: { grid: { display: false } } },
                    plugins: { legend: { display: false }, datalabels: { anchor: 'end', align: 'top', font: { weight: 'bold' } } } 
                }
            }));
        }

        function getPieOptions() {
            return {
                maintainAspectRatio: false, cutout: '40%',
                plugins: {
                    legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 9 } } },
                    datalabels: {
                        color: '#fff', font: { weight: 'bold', size: 9 },
                        formatter: (val, ctx) => val > 0 ? ctx.chart.data.labels[ctx.dataIndex] + '\n' + val : ''
                    }
                }
            };
        }

        function getBarOptions() {
            return {
                maintainAspectRatio: false,
                scales: { y: { beginAtZero: true, ticks: { font: { size: 8 } } }, x: { ticks: { font: { size: 8 } } } },
                plugins: {
                    legend: { position: 'bottom', labels: { boxWidth: 8, font: { size: 8 } } },
                    datalabels: { anchor: 'end', align: 'top', font: { weight: 'bold', size: 8 } }
                }
            };
        }

        function filterTableInPage() {
            const input = document.getElementById("innerSearch").value.toUpperCase();
            const rows = document.getElementById("mainTableBody").getElementsByTagName("tr");
            for (let i = 0; i < rows.length; i++) {
                const text = rows[i].innerText.toUpperCase();
                rows[i].style.display = text.indexOf(input) > -1 ? "" : "none";
            }
        }

        window.onload = () => {
            Papa.parse(sheetUrl, {
                download: true, header: true, skipEmptyLines: true,
                complete: function(results) {
                    fullData = results.data.map(row => {
                        const newRow = {};
                        for (let key in row) { newRow[key.trim().toUpperCase()] = row[key]; }
                        return newRow;
                    }).filter(row => row['NAMA PEGAWAI']);
                    document.getElementById('loading-screen').classList.add('hidden');
                }
            });
        };