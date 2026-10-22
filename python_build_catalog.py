import os
import pandas as pd
import json
import time
import re
import gspread

# Force Python to look in the exact folder where this script is saved
script_folder = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------
# 🔐 THE SECURE UMBC BYPASS: Google Service Account
# ---------------------------------------------------------
# Your Bot IDs (pointing to your copied sheets where you have Owner access)
LITERATURE_SHEET_ID = "1FlezPOj0-48AWBQe3su_Gp1Ugvt5lscDV97PhXatWcc" 
SYLLABI_SHEET_ID = "1YaRYO42MJmGT9CzTeU8bKdXjpJj4OktGRmWPIC98fnk"

# The secret key file downloaded from Google Cloud
CREDENTIALS_FILE = os.path.join(script_folder, "umbc_bot_credentials.json")

# Local Excel fallback files
lit_file = os.path.join(script_folder, "Literature .xlsx")
syl_file = os.path.join(script_folder, "Linguistics & AI Syllabi.xlsx")

# Standardizing tab names based on the design review feedback
TAB_NAME_MAPPING = {
    "Sheet1": "Linguistics & AI Syllabi",
    "lit_0": "Papers with Data",
    "lit_1": "Detection Models",
    "lit_2": "Generation Models"
}

html_template = r"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>iHARP Catalog</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">

    <!-- Pre-load Dark Mode to prevent flashing -->
    <script>
        if (localStorage.theme === 'dark' || (!('theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
            document.documentElement.classList.add('dark')
        } else {
            document.documentElement.classList.remove('dark')
        }
    </script>

    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'sans-serif'] },
                    colors: {
                        actred: '#e52024',
                        actgray: '#f5f5f5',
                        actborder: '#e0e0e0',
                        darkbg: '#0f172a',
                        darkcard: '#1e293b',
                        darkborder: '#334155'
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Inter', sans-serif; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .dark ::-webkit-scrollbar-thumb { background: #475569; }
        
        /* ACT-CMS Red Accent Border for Search */
        .search-container {
            position: relative;
            overflow: hidden;
        }
        .search-container::before {
            content: '';
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
            background-color: #e52024;
            border-top-left-radius: 0.5rem;
            border-bottom-left-radius: 0.5rem;
        }
        
        /* ACT-CMS Dropdown arrow styling */
        select {
            background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%236b7280' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e");
            background-position: right 0.5rem center;
            background-repeat: no-repeat;
            background-size: 1.5em 1.5em;
            padding-right: 2.5rem;
        }
    </style>
</head>
<body class="bg-slate-50 dark:bg-darkbg text-slate-800 dark:text-slate-100 min-h-screen transition-colors duration-200">

    <!-- Top Navigation (Catalog Tabs & Dark Mode Toggle) -->
    <header class="border-b border-slate-200 dark:border-darkborder bg-white dark:bg-darkcard sticky top-0 z-30 shadow-sm transition-colors duration-200">
        <div class="max-w-7xl mx-auto px-6 py-4 flex flex-col md:flex-row justify-between items-center gap-4">
            <div class="flex items-center gap-3">
                <h1 class="text-xl font-bold tracking-tight text-slate-800 dark:text-white">iHARP Catalog</h1>
                <span id="sync-status" class="text-[11px] font-semibold px-2.5 py-0.5 rounded-full bg-emerald-50 dark:bg-emerald-950/60 text-emerald-600 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800">
                    ● Auto-Synced Data
                </span>
            </div>
            <div class="flex items-center gap-6 overflow-x-auto w-full md:w-auto pb-2 md:pb-0">
                <nav id="tab-nav" class="flex gap-2">
                    <!-- Tabs injected dynamically -->
                </nav>
                <button id="theme-toggle" class="text-slate-500 hover:text-actred dark:text-slate-400 dark:hover:text-actred transition-colors text-xl">
                    <i id="theme-toggle-light-icon" class="hidden fas fa-sun"></i>
                    <i id="theme-toggle-dark-icon" class="hidden fas fa-moon"></i>
                </button>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto p-6 mt-4">
        
        <!-- ACT-CMS Style Search & Filter Box -->
        <div class="search-container bg-white dark:bg-darkcard border border-slate-200 dark:border-darkborder rounded-lg p-6 shadow-sm mb-10 transition-colors duration-200">
            <input type="text" id="search-input" 
                class="w-full bg-actgray dark:bg-slate-800 border border-actborder dark:border-slate-700 rounded-md px-4 py-3 text-slate-700 dark:text-slate-200 text-lg focus:outline-none focus:ring-1 focus:ring-slate-400 mb-6 transition-colors placeholder-slate-400 dark:placeholder-slate-500" 
                placeholder="Search lessons, datasets, authors...">
            
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div>
                    <label class="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-2">Year / Date</label>
                    <select id="filter-year" class="w-full bg-actgray dark:bg-slate-800 border border-actborder dark:border-slate-700 rounded-md px-3 py-2 text-slate-700 dark:text-slate-200 focus:outline-none appearance-none transition-colors">
                        <option value="">All Years</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-2">Author / Instructor</label>
                    <select id="filter-author" class="w-full bg-actgray dark:bg-slate-800 border border-actborder dark:border-slate-700 rounded-md px-3 py-2 text-slate-700 dark:text-slate-200 focus:outline-none appearance-none transition-colors">
                        <option value="">All Authors / Instructors</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm font-semibold text-slate-700 dark:text-slate-300 mb-2">Venue / Institution</label>
                    <select id="filter-venue" class="w-full bg-actgray dark:bg-slate-800 border border-actborder dark:border-slate-700 rounded-md px-3 py-2 text-slate-700 dark:text-slate-200 focus:outline-none appearance-none transition-colors">
                        <option value="">All Venues / Institutions</option>
                    </select>
                </div>
            </div>
        </div>

        <!-- Catalog Results Information -->
        <div id="page-info" class="mb-6 flex justify-between items-end">
            <h2 id="page-title" class="text-2xl font-bold text-slate-800 dark:text-white transition-colors">Catalog</h2>
            <p id="page-subtitle" class="text-sm text-slate-500 dark:text-slate-400 font-medium">0 results</p>
        </div>

        <!-- Cards Grid Area -->
        <div id="content-container" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 pb-12">
            <!-- Data cards injected dynamically -->
        </div>

    </main>

    <!-- Modal for viewing full details -->
    <div id="detail-modal" class="fixed inset-0 z-50 hidden" aria-labelledby="modal-title" role="dialog" aria-modal="true">
        <div class="fixed inset-0 bg-slate-900/70 backdrop-blur-sm transition-opacity" onclick="closeModal()"></div>
        <div class="fixed inset-0 z-10 overflow-y-auto">
            <div class="flex min-h-full items-end justify-center p-4 text-center sm:items-center sm:p-0">
                <div class="relative transform overflow-hidden rounded-xl bg-white dark:bg-darkcard text-left shadow-2xl transition-all sm:my-8 sm:w-full sm:max-w-4xl border border-slate-200 dark:border-darkborder">
                    <div class="bg-slate-50 dark:bg-slate-800 border-b border-slate-200 dark:border-darkborder px-6 py-4 flex justify-between items-center sticky top-0 z-10">
                        <h3 class="text-lg font-bold text-slate-800 dark:text-white" id="modal-title">Record Details</h3>
                        <button onclick="closeModal()" class="text-slate-400 hover:text-actred dark:text-slate-500 dark:hover:text-actred transition-colors text-xl">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                    <div class="px-6 py-6 max-h-[70vh] overflow-y-auto" id="modal-content">
                        <!-- Content goes here -->
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- JavaScript Data and Logic -->
    <script>
        let EMBEDDED_DATA = {json_data};
        let activeTabKey = null;

        function shouldSkipTab(title) {
            const t = (title || "").toLowerCase();
            return t.includes("without data") || t.includes("more computer");
        }

        function linkify(text) {
            if (!text) return "";
            let str = text.toString();
            const urlRegex = /(https?:\/\/[^\s]+)/g;
            str = str.replace(urlRegex, function(url) {
                let cleanUrl = url;
                let trailing = "";
                if (/[.,;)]$/.test(url)) {
                    trailing = url.slice(-1);
                    cleanUrl = url.slice(0, -1);
                }
                return `<a href="${cleanUrl}" target="_blank" rel="noopener noreferrer" class="inline-flex items-center gap-1.5 px-2.5 py-1 bg-actred text-white hover:bg-red-700 rounded-md transition-colors font-semibold text-xs tracking-wide whitespace-nowrap shadow-sm mx-1"><i class="fas fa-external-link-alt"></i> View Link</a>${trailing}`;
            });

            if ((str.match(/<a href/g) || []).length > 1) {
                str = str.replace(/,\s*(?=[A-Z0-9])/g, ',<br><span class="block mt-3 mb-3 border-t border-slate-200 dark:border-slate-700/50"></span>');
            }
            return str;
        }

        function extractYear(dateString) {
            if (!dateString) return null;
            const str = dateString.toString();
            const match = str.match(/\b(19|20)\d{2}\b/);
            return match ? match[0] : str.trim();
        }

        function isDateColumn(columnName) {
            return /^(year|when|date)/i.test(columnName.toString().trim());
        }

        function renderTabs() {
            const tabNav = document.getElementById('tab-nav');
            tabNav.innerHTML = '';
            let isFirst = !activeTabKey || !EMBEDDED_DATA[activeTabKey];

            for (const [key, source] of Object.entries(EMBEDDED_DATA)) {
                if (shouldSkipTab(source.title)) continue;

                const isSelected = isFirst || key === activeTabKey;
                const btn = document.createElement('button');
                btn.className = `tab-btn px-4 py-2 text-sm font-semibold rounded-md whitespace-nowrap transition-colors border ${isSelected ? 'bg-slate-800 text-white border-slate-800 dark:bg-slate-200 dark:text-slate-900 dark:border-slate-200' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-100 dark:bg-darkcard dark:text-slate-300 dark:border-darkborder dark:hover:bg-slate-800'}`;
                btn.id = `tab-btn-${key}`;
                btn.innerText = source.title;
                btn.onclick = () => {
                    resetFilters();
                    loadTab(key);
                };
                tabNav.appendChild(btn);

                if (isSelected) {
                    loadTab(key);
                    isFirst = false;
                }
            }
        }

        function init() {
            initDarkMode();
            populateDropdowns();
            renderTabs();

            document.getElementById('search-input').addEventListener('input', handleFilters);
            document.getElementById('filter-year').addEventListener('change', handleFilters);
            document.getElementById('filter-author').addEventListener('change', handleFilters);
            document.getElementById('filter-venue').addEventListener('change', handleFilters);
        }

        function initDarkMode() {
            const themeToggleBtn = document.getElementById('theme-toggle');
            const darkIcon = document.getElementById('theme-toggle-dark-icon');
            const lightIcon = document.getElementById('theme-toggle-light-icon');

            if (localStorage.getItem('theme') === 'dark' || (!('theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
                lightIcon.classList.remove('hidden');
            } else {
                darkIcon.classList.remove('hidden');
            }

            themeToggleBtn.addEventListener('click', function() {
                darkIcon.classList.toggle('hidden');
                lightIcon.classList.toggle('hidden');

                if (document.documentElement.classList.contains('dark')) {
                    document.documentElement.classList.remove('dark');
                    localStorage.setItem('theme', 'light');
                } else {
                    document.documentElement.classList.add('dark');
                    localStorage.setItem('theme', 'dark');
                }
            });
        }

        function populateDropdowns() {
            const years = new Set();
            const authors = new Set();
            const venues = new Set();

            Object.values(EMBEDDED_DATA).forEach(source => {
                if (!source.data || shouldSkipTab(source.title)) return;
                source.data.forEach(row => {
                    Object.keys(row).forEach(k => {
                        const kl = k.toLowerCase();
                        const val = row[k];
                        if (!val || val.length > 100) return;
                        
                        if (isDateColumn(k)) {
                            const y = extractYear(val);
                            if (y) years.add(y);
                        }
                        if (kl.includes('author') || kl.includes('instructor')) {
                            val.split(',').forEach(v => { if (v.trim()) authors.add(v.trim()); });
                        }
                        if (kl.includes('where published') || kl.includes('university') || kl.includes('venue') || kl.includes('journal')) {
                            venues.add(val.trim());
                        }
                    });
                });
            });

            const fillSelect = (id, set, defaultText) => {
                const select = document.getElementById(id);
                select.innerHTML = `<option value="">${defaultText}</option>`;
                Array.from(set).sort().forEach(item => {
                    select.innerHTML += `<option value="${item}">${item}</option>`;
                });
            };

            fillSelect('filter-year', years, 'All Years');
            fillSelect('filter-author', authors, 'All Authors / Instructors');
            fillSelect('filter-venue', venues, 'All Venues / Institutions');
        }

        function resetFilters() {
            document.getElementById('search-input').value = '';
            document.getElementById('filter-year').value = '';
            document.getElementById('filter-author').value = '';
            document.getElementById('filter-venue').value = '';
        }

        function loadTab(key) {
            activeTabKey = key;
            
            document.querySelectorAll('.tab-btn').forEach(t => {
                t.className = 'tab-btn px-4 py-2 text-sm font-semibold rounded-md whitespace-nowrap transition-colors border bg-white text-slate-600 border-slate-200 hover:bg-slate-100 dark:bg-darkcard dark:text-slate-300 dark:border-darkborder dark:hover:bg-slate-800';
            });
            const activeBtn = document.getElementById(`tab-btn-${key}`);
            if (activeBtn) {
                activeBtn.className = 'tab-btn px-4 py-2 text-sm font-semibold rounded-md whitespace-nowrap transition-colors border bg-slate-800 text-white border-slate-800 dark:bg-slate-200 dark:text-slate-900 dark:border-slate-200';
            }
            
            const source = EMBEDDED_DATA[key];
            document.getElementById('page-title').innerText = source.title;
            const data = source.data || [];
            document.getElementById('page-subtitle').innerText = `${data.length} entries`;
            
            renderCards(data, activeTabKey);
        }

        function handleFilters() {
            const searchText = document.getElementById('search-input').value.toLowerCase().trim();
            const yearFilter = document.getElementById('filter-year').value;
            const authorFilter = document.getElementById('filter-author').value;
            const venueFilter = document.getElementById('filter-venue').value;

            if (searchText === '' && yearFilter === '' && authorFilter === '' && venueFilter === '') {
                if (activeTabKey) loadTab(activeTabKey);
                return;
            }

            document.querySelectorAll('.tab-btn').forEach(t => {
                t.className = 'tab-btn px-4 py-2 text-sm font-semibold rounded-md whitespace-nowrap transition-colors border bg-white text-slate-600 border-slate-200 hover:bg-slate-100 dark:bg-darkcard dark:text-slate-300 dark:border-darkborder dark:hover:bg-slate-800';
            });
            document.getElementById('page-title').innerText = `Search Results`;

            let results = [];

            for (const [key, source] of Object.entries(EMBEDDED_DATA)) {
                if (!source.data || shouldSkipTab(source.title)) continue;
                
                source.data.forEach(row => {
                    const rowText = Object.values(row).join(' ').toLowerCase();
                    if (searchText !== '' && !rowText.includes(searchText)) return;
                    
                    if (yearFilter !== '') {
                        let rowYear = '';
                        Object.keys(row).forEach(k => {
                            if(isDateColumn(k)) rowYear = extractYear(row[k]);
                        });
                        if (rowYear !== yearFilter) return;
                    }

                    if (authorFilter !== '') {
                        let hasAuthor = false;
                        Object.keys(row).forEach(k => {
                            if(k.toLowerCase().includes('author') || k.toLowerCase().includes('instructor')) {
                                if (row[k].includes(authorFilter)) hasAuthor = true;
                            }
                        });
                        if (!hasAuthor) return;
                    }

                    if (venueFilter !== '') {
                        let hasVenue = false;
                        Object.keys(row).forEach(k => {
                            if(k.toLowerCase().includes('where published') || k.toLowerCase().includes('university') || k.toLowerCase().includes('venue') || k.toLowerCase().includes('journal')) {
                                if (row[k].includes(venueFilter)) hasVenue = true;
                            }
                        });
                        if (!hasVenue) return;
                    }

                    let cleanRow = {...row, "_source_catalog": source.title, "_tab_key": key};
                    results.push(cleanRow);
                });
            }

            document.getElementById('page-subtitle').innerText = `${results.length} matching entries`;
            renderCards(results, null);
        }

        function renderCards(data, defaultTabKey) {
            const container = document.getElementById('content-container');
            
            if (!data || data.length === 0) {
                container.innerHTML = `
                <div class="col-span-full py-16 text-center text-slate-500 dark:text-slate-400 bg-white dark:bg-darkcard rounded-xl border border-slate-200 dark:border-darkborder shadow-sm">
                    <i class="fas fa-search text-4xl mb-4 text-slate-300 dark:text-slate-600"></i>
                    <p class="text-lg">No records found matching your criteria.</p>
                </div>`;
                return;
            }

            let html = '';
            
            data.forEach((row) => {
                const headers = Object.keys(row).filter(k => k !== "_source_catalog" && k !== "_tab_key");
                const title = row[headers[0]] || 'Untitled';
                const description = headers.length > 1 ? row[headers[1]] : '';
                
                let tags = [];
                let metadataHeaders = [];
                
                headers.slice(2).forEach(h => {
                    const headerLower = h.toLowerCase();
                    if (headerLower.includes('tag') || headerLower.includes('keyword') || headerLower.includes('topic')) {
                        if (row[h]) {
                            tags = tags.concat(row[h].split(/[,;]/).map(t => t.trim()).filter(t => t !== ''));
                        }
                    } else if (row[h] && row[h].trim() !== '') {
                        metadataHeaders.push(h);
                    }
                });

                let metaHtml = '';
                metadataHeaders.slice(0, 4).forEach(h => {
                    metaHtml += `
                    <div class="mb-4">
                        <h4 class="text-[11px] font-semibold uppercase tracking-wider text-slate-500 dark:text-slate-400 mb-1">${h}:</h4>
                        <p class="text-[15px] font-medium text-slate-800 dark:text-slate-200 break-words">${linkify(row[h])}</p>
                    </div>`;
                });

                let tagsHtml = '';
                if (tags.length > 0) {
                    tagsHtml = `<div class="flex flex-wrap gap-2 pt-4 mt-auto border-t border-slate-100 dark:border-slate-700">`;
                    tags.forEach(tag => {
                        tagsHtml += `<span class="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-600 text-slate-600 dark:text-slate-300 text-xs px-3 py-1 rounded-full inline-block break-words max-w-full">${tag}</span>`;
                    });
                    tagsHtml += `</div>`;
                }

                const sourceBadge = row["_source_catalog"] 
                    ? `<span class="text-[10px] font-bold uppercase tracking-wider text-actred mb-2 block">${row["_source_catalog"]}</span>` 
                    : '';

                const tabKeyToUse = row["_tab_key"] || defaultTabKey;
                const originalIndex = EMBEDDED_DATA[tabKeyToUse].data.findIndex(origRow => origRow[headers[0]] === row[headers[0]]);

                html += `
                <div class="bg-white dark:bg-darkcard border border-slate-200 dark:border-darkborder rounded-xl p-6 shadow-sm hover:shadow-md transition-shadow flex flex-col h-full relative">
                    ${sourceBadge}
                    <h3 class="text-xl font-bold text-slate-800 dark:text-white mb-3 leading-tight pr-8 break-words">${title}</h3>
                    <button onclick="openModal('${tabKeyToUse}', ${originalIndex})" class="absolute top-6 right-6 text-slate-400 hover:text-actred dark:text-slate-500 dark:hover:text-actred transition-colors text-lg" title="View Details">
                        <i class="fas fa-expand-arrows-alt"></i>
                    </button>
                    ${description ? `<p class="text-slate-600 dark:text-slate-400 text-sm mb-6 line-clamp-3 break-words">${linkify(description)}</p>` : ''}
                    <div class="flex-1">${metaHtml}</div>
                    ${tagsHtml}
                </div>`;
            });
            
            container.innerHTML = html;
        }

        function openModal(tabKey, rowIndex) {
            const rowData = EMBEDDED_DATA[tabKey].data[rowIndex];
            const contentDiv = document.getElementById('modal-content');
            
            let html = '<dl class="grid grid-cols-1 gap-x-4 gap-y-6 sm:grid-cols-2">';
            
            for (const [key, value] of Object.entries(rowData)) {
                if (!value || value.trim() === '') continue;
                const isLongText = value.length > 80 || key.toLowerCase().includes('description') || key.toLowerCase().includes('abstract') || key.toLowerCase().includes('notes');
                const colSpanClass = isLongText ? 'sm:col-span-2' : 'sm:col-span-1';
                let finalValue = linkify(value);

                html += `
                <div class="${colSpanClass} bg-slate-50 dark:bg-slate-800/50 p-4 rounded-lg border border-slate-100 dark:border-slate-700/50">
                    <dt class="text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-2">${key}</dt>
                    <dd class="text-sm text-slate-900 dark:text-slate-200 leading-relaxed whitespace-pre-wrap break-words">${finalValue}</dd>
                </div>`;
            }
            
            html += '</dl>';
            contentDiv.innerHTML = html;
            document.getElementById('detail-modal').classList.remove('hidden');
            document.body.style.overflow = 'hidden';
        }

        function closeModal() {
            document.getElementById('detail-modal').classList.add('hidden');
            document.body.style.overflow = 'auto';
        }

        window.onload = init;
    </script>
</body>
</html>
"""

def should_skip_tab(title):
    t = str(title).lower()
    return "without data" in t or "more computer" in t

def authenticate_google_sheets():
    """Authenticates using your personal UMBC Google account (OAuth)."""
    print("⏳ Checking credentials... A web browser may open to ask for your login.")
    try:
        # This looks for your client_secret.json and generates a token.json
        client = gspread.oauth(
            credentials_filename=os.path.join(script_folder, "client_secret.json"),
            authorized_user_filename=os.path.join(script_folder, "token.json")
        )
        print("✅ Authentication successful! Logged in securely as YOU.")
        return client
    except Exception as e:
        print(f"⚠️ Could not authenticate: {e}")
        return None

def fetch_secure_sheet(client, sheet_id, file_prefix):
    """Securely fetches data directly from the restricted UMBC Google Sheets."""
    if not client: 
        return {}
    
    print(f"📡 Fetching live data from restricted UMBC Sheet ({file_prefix})...")
    data_dict = {}
    try:
        sheet = client.open_by_key(sheet_id)
        for idx, worksheet in enumerate(sheet.worksheets()):
            all_values = worksheet.get_all_values()
            if len(all_values) < 2:
                continue
                
            raw_headers = all_values[0]
            rows = all_values[1:]
            
            headers = [h.strip() if h.strip() != "" else f"Column_{i}" for i, h in enumerate(raw_headers)]
            df = pd.DataFrame(rows, columns=headers).fillna("")
            df = df[df.apply(lambda row: any(str(val).strip() != "" for val in row), axis=1)]
            
            if len(df) > 0:
                for col in df.columns:
                    if re.match(r'^(year|when|date)', str(col).strip(), re.IGNORECASE):
                        df[col] = df[col].astype(str).apply(
                            lambda x: re.search(r'\b(19|20)\d{2}\b', x).group(0) if re.search(r'\b(19|20)\d{2}\b', x) else x.replace('00:00:00', '').strip()
                        )
                tab_id = f"{file_prefix}_{idx}"
                raw_title = worksheet.title.strip()
                title = TAB_NAME_MAPPING.get(raw_title, TAB_NAME_MAPPING.get(tab_id, raw_title))
                
                if not should_skip_tab(title):
                    data_dict[tab_id] = {
                        "title": title,
                        "data": df.astype(str).to_dict(orient="records")
                    }
        print(f"   ✅ Successfully pulled {len(data_dict)} tabs from ({file_prefix})!")
    except Exception as e:
        print(f"❌ Error fetching from Google Sheets ({file_prefix}): {repr(e)}")
        print("   👉 Note: If this says PermissionError (403), make sure you opened your copied sheets and clicked 'Share' to add your bot's email address!")
    return data_dict

def build_site():
    print("⏳ Building catalog...")
    data_dict = {}

    client = authenticate_google_sheets()
    if client and LITERATURE_SHEET_ID != "YOUR_LITERATURE_SHEET_ID_HERE":
        lit_data = fetch_secure_sheet(client, LITERATURE_SHEET_ID, "lit")
        data_dict.update(lit_data)
        
        syl_data = fetch_secure_sheet(client, SYLLABI_SHEET_ID, "syl")
        data_dict.update(syl_data)

    if not data_dict:
        print("⚠️ Bot credentials not active. Falling back to Local Excel files...")
        
        def process_excel(file_path, file_prefix):
            if not os.path.exists(file_path): return
            try: xls = pd.ExcelFile(file_path)
            except Exception: return
            
            for idx, sheet_name in enumerate(xls.sheet_names):
                df = pd.read_excel(xls, sheet_name=sheet_name).fillna("")
                if len(df) > 0:
                    for col in df.columns:
                        if re.match(r'^(year|when|date)', str(col).strip(), re.IGNORECASE):
                            df[col] = df[col].astype(str).apply(
                                lambda x: re.search(r'\b(19|20)\d{2}\b', x).group(0) if re.search(r'\b(19|20)\d{2}\b', x) else x.replace('00:00:00', '').strip()
                            )
                    tab_id = f"{file_prefix}_{idx}"
                    raw_title = sheet_name.strip()
                    title = TAB_NAME_MAPPING.get(raw_title, TAB_NAME_MAPPING.get(tab_id, raw_title))
                    if not should_skip_tab(title):
                        data_dict[tab_id] = {"title": title, "data": df.astype(str).to_dict(orient="records")}
                        
        process_excel(lit_file, "lit")
        process_excel(syl_file, "syl")

    json_data = json.dumps(data_dict)
    output_html = os.path.join(script_folder, "index.html")
    
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_template.replace("{json_data}", json_data))
        
    print(f"✅ Successfully generated {output_html} at {time.strftime('%I:%M:%S %p')}")
    return json_data

if __name__ == "__main__":
    print("🚀 Starting the Secure UMBC Catalog Builder...")
    last_json = build_site()
    
    print("\n👀 Watching UMBC Google Sheets for live changes every 30 seconds... (Press Ctrl+C to stop)")
    try:
        if __name__ == "__main__":
            print("🚀 Starting the Cloud Catalog Builder...")
            build_site()
            print("✅ Build complete. Shutting down.")
    except KeyboardInterrupt:
        print("\n⏹️ Stopped watching. Have a great day!")
