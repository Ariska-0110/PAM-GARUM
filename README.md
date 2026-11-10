# PAM-GARUM
schedule PAM GARUM
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Jadwal Tugas Officer Security PT KCI Area 8 Surabaya</title>

    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- html2pdf.js CDN untuk Ekspor PDF -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <!-- SheetJS CDN untuk Ekspor Excel -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #f1f5f9;
        }

        /* Border Kop Surat Double Line */
        .kop-border {
            border-bottom: 4px double #000;
        }

        /* Styling Tabel KCI Presisi */
        .table-kci {
            border-collapse: collapse;
            width: 100%;
            font-size: 11px;
            color: #000;
        }
        .table-kci th, .table-kci td {
            border: 1px solid #000;
            text-align: center;
            vertical-align: middle;
            padding: 4px 1px;
        }
        
        /* Warna Header khusus KCI */
        .bg-kci-green { background-color: #92d050 !important; color: #000; font-weight: bold; }
        .bg-kci-blue { background-color: #0070c0 !important; color: #fff; font-weight: bold; }
        .bg-kci-red { background-color: #ff0000 !important; color: #fff; font-weight: bold; }
        .bg-kci-yellow { background-color: #ffff00 !important; color: #000; font-weight: bold; }
        
        /* Highlight khusus Shift Rest & Libur (R / L) */
        .shift-rl {
            background-color: #ff0000 !important;
            color: #ffffff !important;
            font-weight: bold;
        }

        .cell-shift {
            font-weight: bold;
        }

        /* Area Dokumen Cetak */
        #document-preview {
            background: white;
            width: 1180px;
            margin: 0 auto;
            padding: 25px 35px;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        }

        @media print {
            .no-print { display: none !important; }
            body { background: white; padding: 0; margin: 0; }
            #document-preview {
                box-shadow: none;
                width: 100%;
                padding: 0;
            }
        }
    </style>
</head>
<body class="p-4 md:p-8">

    <!-- PANEL PENGATURAN PERIODE & PERSONEL -->
    <div class="no-print max-w-6xl mx-auto bg-white p-6 rounded-xl shadow-md mb-8 border border-gray-200">
        <h2 class="text-xl font-bold text-gray-800 mb-4 flex items-center gap-2">
            📅 Pengaturan Periode & Personel Security
        </h2>

        <!-- Input Periode -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <div>
                <label class="block text-sm font-semibold text-gray-700 mb-1">Pilih Bulan</label>
                <select id="selectMonth" onchange="generateSchedule()" class="w-full border border-gray-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                    <option value="0">Januari</option>
                    <option value="1">Februari</option>
                    <option value="2">Maret</option>
                    <option value="3">April</option>
                    <option value="4">Mei</option>
                    <option value="5">Juni</option>
                    <option value="6">Juli</option>
                    <option value="7">Agustus</option>
                    <option value="8">September</option>
                    <option value="9">Oktober</option>
                    <option value="10" selected>November</option>
                    <option value="11">Desember</option>
                </select>
            </div>

            <div>
                <label class="block text-sm font-semibold text-gray-700 mb-1">Pilih Tahun</label>
                <input type="number" id="inputYear" value="2026" onchange="generateSchedule()" class="w-full border border-gray-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
            </div>

            <div>
                <label class="block text-sm font-semibold text-gray-700 mb-1">Nama Stasiun / Lokasi</label>
                <input type="text" id="inputStation" value="ST.GARUM" oninput="generateSchedule()" class="w-full border border-gray-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
            </div>
        </div>

        <!-- Form Tambah Personel -->
        <div class="bg-gray-50 p-4 rounded-lg border border-gray-200 mb-6">
            <h3 class="text-sm font-bold text-gray-700 mb-2">➕ Tambah / Kelola Personel</h3>
            <div class="grid grid-cols-1 md:grid-cols-4 gap-3">
                <input type="text" id="newName" placeholder="Nama Lengkap Personel" class="border border-gray-300 rounded-lg p-2 text-sm">
                <input type="text" id="newNip" placeholder="NIPKWT" class="border border-gray-300 rounded-lg p-2 text-sm">
                <select id="newOffset" class="border border-gray-300 rounded-lg p-2 text-sm">
                    <option value="0">Shift Awal Nov 1: S (Siang)</option>
                    <option value="6">Shift Awal Nov 1: P (Pagi)</option>
                    <option value="4">Shift Awal Nov 1: M (Malam)</option>
                    <option value="2">Shift Awal Nov 1: R (Rest)</option>
                </select>
                <button onclick="addOfficer()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg text-sm transition">
                    + Tambah Nama
                </button>
            </div>
            <!-- Daftar Personel Terdaftar -->
            <div id="officerList" class="flex flex-wrap gap-2 mt-3"></div>
        </div>

        <!-- Tombol Aksi Ekspor & Cetak -->
        <div class="flex flex-wrap gap-3 border-t pt-4">
            <button onclick="downloadPDF()" class="bg-red-600 hover:bg-red-700 text-white font-bold py-2.5 px-6 rounded-lg shadow transition flex items-center gap-2">
                📄 Download PDF
            </button>
            <button onclick="downloadExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2.5 px-6 rounded-lg shadow transition flex items-center gap-2">
                📊 Download Excel (.xlsx)
            </button>
            <button onclick="window.print()" class="bg-gray-800 hover:bg-gray-900 text-white font-bold py-2.5 px-6 rounded-lg shadow transition flex items-center gap-2">
                🖨 Cetak / Print
            </button>
        </div>
    </div>

    <!-- PREVIEW DOKUMEN JADWAL -->
    <div id="document-preview">
        <!-- HEADER KOP SURAT -->
        <div class="text-center font-bold text-black uppercase mb-1">
            <h1 class="text-base md:text-lg leading-tight tracking-wide">JADWAL TUGAS OFFICER SECURITY PT KERETA COMMUTER INDONESIA AREA 8 SURABAYA</h1>
            <h2 class="text-sm md:text-base leading-tight mt-0.5">BULAN <span id="docMonthYear">NOVEMBER 2026</span></h2>
            <h3 class="text-sm md:text-base leading-tight mt-0.5">(<span id="docStation">ST.GARUM</span>)</h3>
        </div>

        <!-- GARIS DOUBLE KOP SURAT -->
        <div class="kop-border mb-3"></div>

        <!-- TABEL JADWAL ROTASI OTOMATIS -->
        <div class="overflow-x-auto">
            <table class="table-kci" id="scheduleTable">
                <thead>
                    <tr id="rowHeaderStation">
                        <th colspan="3" class="bg-kci-green"></th>
                        <th id="thStationHeader" class="bg-kci-blue uppercase">STASIUN GARUM</th>
                    </tr>
                    <tr id="rowDayNames">
                        <th class="bg-kci-green" style="width: 28px;">NO</th>
                        <th class="bg-kci-green" style="width: 180px;">NAMA</th>
                        <th class="bg-kci-green" style="width: 80px;">NIPKWT</th>
                    </tr>
                    <tr id="rowDayNumbers"></tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>

        <!-- BARIS KETERANGAN SHIFT -->
        <table class="table-kci mt-2" style="font-size: 10px; font-weight: bold;">
            <tr class="bg-gray-100">
                <td style="width: 15%; text-align: center;" class="bg-gray-200">KETERANGAN :</td>
                <td style="width: 21.25%;">P= Masuk Piket Pagi 05:00-13:00</td>
                <td style="width: 21.25%;">S= Masuk Piket Siang 13:00-21:00</td>
                <td style="width: 21.25%;">M= Masuk Piket Malam 21:00-05:00</td>
                <td style="width: 21.25%;" class="bg-gray-200">R/L= Rest/Off</td>
            </tr>
        </table>

        <!-- KOTAK CATATAN 8 POIN -->
        <div class="mt-3 text-black text-xs leading-relaxed">
            <p class="font-bold mb-1">Catatan:</p>
            <ol class="list-decimal list-inside space-y-0.5 pl-1 text-[11px]">
                <li>Wajib Hadir 30 menit sebelum serah terima tugas sudah harus di tempat dan dalam keadaan pakaian lengkap rapih.</li>
                <li>Melaksanakan serah terima kepada petugas jaga berikutnya serta inventaris secara cermat dan teliti (cek fisik).</li>
                <li>Melakukan tindakan preventif terhadap tindakan kejahatan serta melaporkan kepada atasan.</li>
                <li>Melaporakan dan mencatat setiap kejadian sekecil apapun kepada atasan dan membuat laporan tertulis.</li>
                <li>Melaksanakan kontrol ke seluruh area secara simultan setiap 1 jam sekali.</li>
                <li>Menegur dan melarang setiap orang yang tidak berkepentingan berada di area stasiun (asongan,pengemis dll)</li>
                <li>Apabila petugas pengganti berhalangan hadir maka wajib menggantikannya lembur dan lapaoran ke Chief Security</li>
                <li>Menjaga kebersihan dan kerapihan pos jaga serta lingkungan sekitarnya.</li>
            </ol>
        </div>
    </div>

    <!-- LOGIKA KODE ROTASI DARI BULAN NOVEMBER 2026 -->
    <script>
        // Data Personel Default (Diperbarui dengan NAMA LENGKAP)
        let officers = [
            { id: 1, name: "M. RIDHO RAHMAN", nip: "50805752", offset: 0 },         // Nov 1: S
            { id: 2, name: "DENDY PRADANA", nip: "50805367", offset: 6 },           // Nov 1: P
            { id: 3, name: "RIFKY ARDIAN SUKMAJI", nip: "50804764", offset: 4 },    // Nov 1: M
            { id: 4, name: "NANDA AGUS HARDIANSYAH", nip: "50805176", offset: 2 }  // Nov 1: R
        ];

        const dayAbbr = ["MG", "SN", "SL", "RB", "KM", "JM", "SB"];
        const monthNames = [
            "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
            "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
        ];

        // SIKLUS ROTASI 8 HARI (S -> P -> R -> L -> M -> S -> P -> M)
        const cycle8 = ['S', 'P', 'R', 'L', 'M', 'S', 'P', 'M'];

        window.onload = function() {
            renderOfficerChips();
            generateSchedule();
        };

        function renderOfficerChips() {
            const container = document.getElementById('officerList');
            container.innerHTML = '';
            officers.forEach((o, index) => {
                const chip = document.createElement('span');
                chip.className = "inline-flex items-center gap-1 bg-blue-100 text-blue-800 text-xs font-semibold px-2.5 py-1 rounded-full border border-blue-200";
                chip.innerHTML = `
                    ${o.name} (${o.nip})
                    ${officers.length > 1 ? `<button onclick="removeOfficer(${index})" class="text-red-500 hover:text-red-700 font-bold ml-1">×</button>` : ''}
                `;
                container.appendChild(chip);
            });
        }

        function addOfficer() {
            const nameInput = document.getElementById('newName');
            const nipInput = document.getElementById('newNip');
            const offsetInput = document.getElementById('newOffset');

            if (!nameInput.value.trim()) {
                alert("Harap masukkan nama personel!");
                return;
            }

            officers.push({
                id: Date.now(),
                name: nameInput.value.trim().toUpperCase(),
                nip: nipInput.value.trim() || "-",
                offset: parseInt(offsetInput.value)
            });

            nameInput.value = '';
            nipInput.value = '';
            renderOfficerChips();
            generateSchedule();
        }

        function removeOfficer(index) {
            officers.splice(index, 1);
            renderOfficerChips();
            generateSchedule();
        }

        function generateSchedule() {
            const month = parseInt(document.getElementById('selectMonth').value);
            const year = parseInt(document.getElementById('inputYear').value);
            const station = document.getElementById('inputStation').value.toUpperCase();

            // Render Judul Header Dokumen
            document.getElementById('docMonthYear').innerText = `${monthNames[month]} ${year}`;
            document.getElementById('docStation').innerText = station;
            document.getElementById('thStationHeader').innerText = `STASIUN ${station.replace('ST.', '').trim()}`;

            // Hitung jumlah hari dalam bulan terpilih
            const daysInMonth = new Date(year, month + 1, 0).getDate();
            document.getElementById('thStationHeader').setAttribute('colspan', daysInMonth);

            // Render Header Nama Hari & Angka Tanggal
            const rowDayNames = document.getElementById('rowDayNames');
            const rowDayNumbers = document.getElementById('rowDayNumbers');

            rowDayNames.innerHTML = `
                <th class="bg-kci-green" style="width: 28px;">NO</th>
                <th class="bg-kci-green" style="width: 180px;">NAMA</th>
                <th class="bg-kci-green" style="width: 80px;">NIPKWT</th>
            `;
            rowDayNumbers.innerHTML = ``;

            for (let day = 1; day <= daysInMonth; day++) {
                const dateObj = new Date(year, month, day);
                const dayOfWeek = dateObj.getDay(); 
                const isSunday = dayOfWeek === 0;

                // Cell Nama Hari
                const thDay = document.createElement('th');
                thDay.innerText = dayAbbr[dayOfWeek];
                thDay.className = isSunday ? 'bg-kci-red' : 'bg-kci-yellow';
                thDay.style.width = '24px';
                rowDayNames.appendChild(thDay);

                // Cell Angka Tanggal
                const thNum = document.createElement('th');
                thNum.innerText = day;
                thNum.className = isSunday ? 'bg-kci-red' : 'bg-kci-yellow';
                rowDayNumbers.appendChild(thNum);
            }

            // Render Baris Shift Personel
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            officers.forEach((officer, index) => {
                const tr = document.createElement('tr');

                tr.innerHTML = `
                    <td class="font-bold">${index + 1}</td>
                    <td class="text-left font-bold px-1 whitespace-nowrap">${officer.name}</td>
                    <td>${officer.nip}</td>
                `;

                for (let day = 1; day <= daysInMonth; day++) {
                    const td = document.createElement('td');

                    // Hitung shift permanen berpatokan penuh pada 1 November 2026
                    const shiftValue = calculateShiftFromNov2026(officer.offset, year, month, day);

                    td.innerText = shiftValue;
                    // Beri warna merah khusus jika Rest (R) atau Libur (L)
                    const isRestOrOff = (shiftValue === 'R' || shiftValue === 'L' || shiftValue === 'R/L');
                    td.className = `cell-shift ${isRestOrOff ? 'shift-rl' : ''}`;

                    tr.appendChild(td);
                }

                tbody.appendChild(tr);
            });
        }

        // FUNGSI UTAMA: MENGHITUNG ROTASI BERPATOKAN DARI 1 NOVEMBER 2026
        function calculateShiftFromNov2026(baseOffset, year, month, day) {
            // Tanggal Acuan Utama: 1 November 2026 (UTC)
            const nov1_2026_UTC = Date.UTC(2026, 10, 1);
            const targetUTC = Date.UTC(year, month, day);

            // Hitung selisih hari presisi
            const diffDays = Math.floor((targetUTC - nov1_2026_UTC) / (1000 * 60 * 60 * 24));

            // Perhitungan Modulo Siklus 8 Hari
            let shiftIndex = (baseOffset + diffDays) % 8;
            if (shiftIndex < 0) {
                shiftIndex = (shiftIndex % 8) + 8;
            }

            return cycle8[shiftIndex];
        }

        // Ekspor ke PDF
        function downloadPDF() {
            const element = document.getElementById('document-preview');
            const monthStr = document.getElementById('docMonthYear').innerText;
            
            const opt = {
                margin:       [0.2, 0.2, 0.2, 0.2],
                filename:     `Jadwal_PAM_KCI_Garum_${monthStr.replace(' ', '_')}.pdf`,
                image:        { type: 'jpeg', quality: 0.98 },
                html2canvas:  { scale: 2, useCORS: true },
                jsPDF:        { unit: 'in', format: 'a4', orientation: 'landscape' }
            };

            html2pdf().set(opt).from(element).save();
        }

        // Ekspor ke Excel (.xlsx)
        function downloadExcel() {
            const table = document.getElementById('scheduleTable');
            const wb = XLSX.utils.table_to_book(table, { sheet: "Jadwal Security" });
            const monthStr = document.getElementById('docMonthYear').innerText;
            
            XLSX.writeFile(wb, `Jadwal_PAM_KCI_Garum_${monthStr.replace(' ', '_')}.xlsx`);
        }
    </script>
</body>
</html>
