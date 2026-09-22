    /* Functions referenced by inline onclick/onchange handlers in index.html. */
    /* exported refreshData, resetForm, saveData, cancelEdit, carryForward,
              openSheet, addAtm, saveSettings, setRange, generateReport,
              exportToExcel, handleDualPagePrint, generateDownloadablePDF,
              openCamera, capturePhoto, closeCamera, printEntry, editEntry, deleteEntry,
              installApp, multiSelectAll, multiClear, updateMultiSummary, sendWhatsApp, sendWhatsAppMulti */

    /* ================= CONFIGURATION ================= */
    // NOTE: Client-side secrets are visible to anyone. The token protects the sheet
    // from casual access only. For real security, deploy with Google sign-in and use
    // Session.getEffectiveUser() server-side (see CONTRACT.md). Token is sent ONLY
    // in POST bodies, never in URLs.
    const GAS_URL = "https://script.google.com/macros/s/AKfycbzKOsvF3DfhC9-WDuDHknreR6WTTFanTqVo1l2rvyx07LkEada1z7kBSPSg-QTmwd_Q/exec";
    const SECRET_TOKEN = "MY_SECRET_2025";
    const DEF_SHEET_ID = "1p3RJHBlNoX2agEIYkFScXZhOMh8NrZjQnmT9fK8KcHs";
    let D1 = 200, D2 = 100, D3 = 500;

    let globalData = [];
    let atmList = ['FZIS0005'];
    let isEditingId = null;

    const $ = id => document.getElementById(id);
    const showLoader = s => $('loader').style.display = s ? 'flex' : 'none';
    const notify = (msg, type='success') => {
      const n = $('notif'); n.textContent=msg; n.className=`notification ${type} show`;
      setTimeout(()=>n.classList.remove('show'),3000);
    };

    /* ================= PWA / OFFLINE SYNC ================= */
    function getPending(){ try{ return JSON.parse(localStorage.getItem('ctm_pending')||'[]'); }catch(e){ return []; } }
    function setPending(list){ localStorage.setItem('ctm_pending', JSON.stringify(list)); }

    async function sendPayload(payload, opts={}){
      const queueIfOffline = !!opts.queueIfOffline;
      try {
        const res = await fetch(GAS_URL, {method:'POST', body:JSON.stringify(payload)});
        const json = await res.json();
        if(json.status!=='ok') throw new Error(json.message||'Server error');
        return json;
      } catch(err) {
        if(queueIfOffline && (!navigator.onLine || err instanceof TypeError)){
          const q = getPending(); q.push({ts:Date.now(), payload}); setPending(q);
          updatePendingBadge();
          return {queued:true};
        }
        throw err;
      }
    }

    async function flushPending(){
      const q = getPending();
      if(!q.length) return 0;
      let synced = 0;
      for(const item of q){
        try { await sendPayload(item.payload); synced++; }
        catch(e){ break; }
      }
      if(synced > 0){
        setPending(getPending().slice(synced));
        updatePendingBadge();
        notify(synced + ' offline entr' + (synced>1?'ies':'y') + ' synced', 'success');
        await refreshData();
      }
      return synced;
    }

    function updatePendingBadge(){
      const n = getPending().length;
      if($('pendingBadge')) $('pendingBadge').style.display = n ? 'flex' : 'none';
      if($('pendingCount')) $('pendingCount').textContent = n;
    }

    function registerSW(){
      if('serviceWorker' in navigator && location.protocol.startsWith('http')){
        window.addEventListener('load', () => {
          navigator.serviceWorker.register('sw.js').catch(e => console.error('SW registration failed', e));
        });
      }
    }

    let deferredPrompt;
    function installApp(){
      if(!deferredPrompt) return notify('Use your browser menu to install', 'info');
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then(() => { deferredPrompt = null; });
    }
    function setupInstallPrompt(){
      window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        const btn = $('installBtn');
        if(btn) btn.style.display = 'inline-flex';
      });
      window.addEventListener('appinstalled', () => {
        deferredPrompt = null;
        const btn = $('installBtn');
        if(btn) btn.style.display = 'none';
        notify('App installed', 'success');
      });
      window.addEventListener('online', () => { flushPending(); });
      window.addEventListener('offline', () => { updatePendingBadge(); });
    }

    const getLocalDate = () => {
        const d = new Date();
        return new Date(d.getTime() - (d.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
    };
    const cleanNum = (v) => {
       if(typeof v === 'number') return v;
       if(!v) return 0;
       return Number(String(v).replace(/[^\d]/g, '')) || 0;
    };

    document.addEventListener('DOMContentLoaded', async () => {
      loadSettings();
      loadDraft();
      registerSW();
      setupInstallPrompt();
      updatePendingBadge();
      if(navigator.onLine) flushPending();

      document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(x=>x.classList.remove('active'));
        t.classList.add('active');
        $(t.dataset.tab+'Tab').classList.add('active');
        if(t.dataset.tab === 'analytics') renderAnalytics();
        if(t.dataset.tab === 'multiday') renderMultiList();
      }));

      document.querySelectorAll('.calc-trigger').forEach(i => i.addEventListener('input', () => { calculate(); saveDraft(); }));
      $('bank_indent_val').addEventListener('input', () => { calculate(); saveDraft(); });
      $('searchHistory').addEventListener('keyup', renderHistory);
      $('new_atm').addEventListener('keyup', (e) => { if(e.key==='Enter') addAtm(); });
      
      // Image Upload Listener
      $('img_input').addEventListener('change', function() {
          if(this.files && this.files[0]) {
              const reader = new FileReader();
              reader.onload = e => { 
                  $('preview_thumb').src = e.target.result; 
                  $('img_preview').style.display='block'; 
              };
              reader.readAsDataURL(this.files[0]);
          }
      });

      // PDF Upload -> auto-extract Indent No., Replenishment Amount, Date, ATM
      $('pdf_input').addEventListener('change', function(){ extractPdfData(this); });
      $('pdf_input_entry').addEventListener('change', function(){ extractPdfData(this); });

      const today = getLocalDate();
      if(!$('date').value) $('date').value = today;
      $('rep_from').value = today;
      $('rep_to').value = today;
      $('searchHistory').value = ''; // Show all entries; don't pre-filter to today

      await refreshData();
    });

    function saveDraft() {
        if(isEditingId) return;
        const draft = {
            date: $('date').value, bank_indent_no: $('bank_indent_no').value, bank_indent_val: $('bank_indent_val').value,
            atmIds: [$('atm_select').value].filter(Boolean),
            t1_open: $('t1_open').value, t1_added: $('t1_added').value, t2_open: $('t2_open').value, t2_added: $('t2_added').value, t3_open: $('t3_open').value, t3_added: $('t3_added').value
        };
        localStorage.setItem('ctm_draft_v1', JSON.stringify(draft));
    }
    function loadDraft() {
        const d = JSON.parse(localStorage.getItem('ctm_draft_v1'));
        if(d && !isEditingId) {
            if(d.date) $('date').value = d.date; if(d.bank_indent_no) $('bank_indent_no').value = d.bank_indent_no;
            if(d.bank_indent_val) $('bank_indent_val').value = d.bank_indent_val;
            const dAtms = d.atmIds && d.atmIds.length ? d.atmIds : (d.atm_select ? [d.atm_select] : []);
            if(dAtms.length) $('atm_select').value = dAtms[0];
            $('t1_open').value=d.t1_open||''; $('t1_added').value=d.t1_added||''; $('t2_open').value=d.t2_open||''; $('t2_added').value=d.t2_added||'';
            $('t3_open').value=d.t3_open||''; $('t3_added').value=d.t3_added||''; calculate();
        }
    }
    function clearDraft() { localStorage.removeItem('ctm_draft_v1'); }

    async function refreshData() {
      showLoader(true);
      try {
        const sid = $('set_sheet').value || DEF_SHEET_ID;
        // Read via GET (matches the deployed backend; Code.gs v9.8 supports this too).
        const res = await fetch(`${GAS_URL}?action=getAll&token=${SECRET_TOKEN}&sheetId=${sid}&_t=${Date.now()}`);
        const json = await res.json();
        if(json.status === 'ok'){
            // Fix: Clean Data Dates
            globalData = (json.data||[]).map(d => {
                if(d.date && d.date.includes('T')) {
                   const dt = new Date(d.date);
                   if(!isNaN(dt)) d.date = new Date(dt.getTime() - (dt.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
                }
                return d;
            }).filter(e => e.id && e.date); // Skip junk/empty rows
            $('statusDot').className = 'dot online'; $('statusText').textContent = 'Online';
            renderHistory();
            populateMergeDropdown();
            renderMultiList();
        }
        else throw new Error(json.message);
      } catch(e) { console.error(e); $('statusDot').className = 'dot'; $('statusText').textContent = 'Offline'; notify('Sync Error', 'error'); } finally { showLoader(false); }
    }

    function populateMergeDropdown() {
        const sel = $('merge_entry_select');
        sel.innerHTML = '<option value="">-- Select Daily Entry --</option>';
        // Sort by date desc
        const sorted = [...globalData].sort((a,b) => new Date(b.date) - new Date(a.date));
        sorted.forEach(e => {
            const opt = document.createElement('option');
            opt.value = e.id;
            opt.text = `${e.date} | ${e.bankIndentNo} | ₹${e.bankIndentVal}`;
            sel.add(opt);
        });
        if(sorted.length > 0) sel.selectedIndex = 1; // Auto select first valid
    }

    function parseIndentDate(str) {
        const months = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
        const m = str.match(/(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
        return (m && months[m[2]]) ? (m[3] + '-' + months[m[2]] + '-' + String(m[1]).padStart(2,'0')) : '';
    }

    async function extractPdfData(input) {
        const file = (input || $('pdf_input')).files[0];
        if(!file) return;
        showLoader(true);
        try {
            const pdf = await pdfjsLib.getDocument(await file.arrayBuffer()).promise;
            const page = await pdf.getPage(1);
            const text = (await page.getTextContent()).items.map(i => i.str).join(' ');
            const norm = text.replace(/\s+/g, ' ');

            const mNo   = norm.match(/Cash\s*Indent\s*No\.?\s*:?\s*([0-9][0-9\-/]*)/i) || norm.match(/Indent\s*No\.?\s*:?\s*([0-9][0-9\-/]*)/i);
            const mAmt  = norm.match(/Total\s*Replenishment\s*Amount\s*:?\s*([\d,]{3,})/i);
            const mDate = norm.match(/\bIndent\s*Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})/i);
            const mAtm  = norm.match(/Total\s+Amt\s+WDL\s+([A-Z0-9]+)/);

            let found = 0;
            if(mNo)   { $('bank_indent_no').value = mNo[1]; found++; }
            if(mAmt)  { $('bank_indent_val').value = Number(mAmt[1].replace(/[^\d]/g,'')); found++; }
            if(mDate) { const iso = parseIndentDate(mDate[1]); if(iso) { $('date').value = iso; found++; } }
            if(mAtm && atmList.includes(mAtm[1])) { $('atm_select').value = mAtm[1]; found++; }
            if(!found) { notify('Could not read Indent fields from this PDF', 'error'); return; }

            calculate();
            if(mNo) {
                const match = globalData.find(x => x.bankIndentNo === mNo[1]);
                if(match) $('merge_entry_select').value = match.id;
            }
            notify('Extracted: ' + (mNo ? mNo[1] : '?') + ' | ₹' + Number(mAmt ? mAmt[1].replace(/[^\d]/g,'') : 0).toLocaleString(), 'success');
        } catch(e) {
            console.error(e);
            notify('PDF extraction failed', 'error');
        } finally { showLoader(false); }
    }

    function calculate() {
      const getVal = id => Number($(id).value) || 0;
      const t1n = getVal('t1_open') + getVal('t1_added'); $('t1_total').value = t1n; $('t1_val').textContent = '₹'+(t1n*D1).toLocaleString();
      const t2n = getVal('t2_open') + getVal('t2_added'); $('t2_total').value = t2n; $('t2_val').textContent = '₹'+(t2n*D2).toLocaleString();
      const t3n = getVal('t3_open') + getVal('t3_added'); $('t3_total').value = t3n; $('t3_val').textContent = '₹'+(t3n*D3).toLocaleString();
      
      const totalNotes = t1n + t2n + t3n;
      const totalVal = (t1n*D1) + (t2n*D2) + (t3n*D3);
      const loadedVal = (getVal('t1_added')*D1) + (getVal('t2_added')*D2) + (getVal('t3_added')*D3);
      const indent = getVal('bank_indent_val');
      const diff = indent - loadedVal;

      $('sum_notes').textContent = totalNotes;
      $('sum_val').textContent = '₹'+totalVal.toLocaleString();
      $('sum_loaded').textContent = '₹'+loadedVal.toLocaleString();
      $('sum_indent').textContent = '₹'+indent.toLocaleString();
      $('sum_diff').textContent = '₹'+diff.toLocaleString();
      $('sum_diff').style.color = diff === 0 ? '#16a34a' : '#ef4444';
    }

    const isPosInt = v => Number.isInteger(v) && v >= 0;

    async function saveData() {
      if(! $('bank_indent_no').value.trim()) return notify('Indent No is required', 'error');
      if(! $('date').value) return notify('Date is required', 'error');
      const getVal = id => Number($(id).value) || 0;
      const atms = [$('atm_select').value].filter(Boolean);
      if(atms.length===0) return notify('Select at least one ATM', 'error');
      const trayFields = ['t1_open','t1_added','t2_open','t2_added','t3_open','t3_added'];
      for(const id of trayFields){ if(!isPosInt(getVal(id))) return notify('Tray values must be whole numbers (0 or more)', 'error'); }
      if(!isEditingId){
        const dup = globalData.find(x => x.date === $('date').value && x.bankIndentNo === $('bank_indent_no').value.trim());
        if(dup && !confirm(`Entry already exists for ${dup.date} / ${dup.bankIndentNo}. Save anyway?`)) return;
      }
      showLoader(true);
      try {
        const clean = t => Number(String(t).replace(/[^\d]/g,''));
        
        const val_notes = clean($('sum_notes').textContent);
        const val_value = clean($('sum_val').textContent);
        const val_loaded = clean($('sum_loaded').textContent);
        
        const payload = {
          _token: SECRET_TOKEN, action: isEditingId ? 'update' : 'save', id: isEditingId || 'ID_'+Date.now(),
          date: $('date').value, bankIndentNo: $('bank_indent_no').value.trim(), bankIndentVal: getVal('bank_indent_val'), atmIds: atms,
          totalNotesAll: val_notes, totalValueAll: val_value, totalLoadedVal: val_loaded,
          trays: {
            t1: {opening: getVal('t1_open'), added: getVal('t1_added'), totalNotes: getVal('t1_total'), totalValue: getVal('t1_total')*D1},
            t2: {opening: getVal('t2_open'), added: getVal('t2_added'), totalNotes: getVal('t2_total'), totalValue: getVal('t2_total')*D2},
            t3: {opening: getVal('t3_open'), added: getVal('t3_added'), totalNotes: getVal('t3_total'), totalValue: getVal('t3_total')*D3},
          },
          summary: { totalNotesAll: val_notes, totalValueAll: val_value, totalLoadedVal: val_loaded }
        };
        const result = await sendPayload(payload, {queueIfOffline:true});
        if(result && result.queued){
          notify('Offline — entry queued for sync', 'info');
          clearDraft(); resetForm();
        } else {
          notify('Entry Saved Successfully'); clearDraft(); resetForm(); await refreshData();
        }
      } catch(e){ notify(e.message,'error'); } finally{ showLoader(false); }
    }

    function cancelEdit(){
      isEditingId = null;
      $('cancelEditBtn').style.display = 'none';
      resetForm();
      notify('Edit cancelled', 'info');
    }

    function resetForm() {
      isEditingId = null;
      if($('cancelEditBtn')) $('cancelEditBtn').style.display='none';
      document.querySelectorAll('.input-box').forEach(i => { if(!i.readOnly && i.tagName==='INPUT' && i.id !== 'date') i.value=''; });
      if(atmList.length > 0) $('atm_select').value = atmList[0];
      calculate();
    }

    function carryForward(){
      const atms = [$('atm_select').value].filter(Boolean);
      const date = $('date').value;
      const candidates = globalData.filter(e => e.date < date && (atms.length===0 || (e.atmIds||[]).some(a=>atms.includes(a))));
      if(candidates.length===0) return notify('No earlier entry found to carry forward', 'info');
      candidates.sort((a,b)=> b.date.localeCompare(a.date));
      const prev = candidates[0];
      $('t1_open').value = Number(prev.trays.t1.totalNotes)||0;
      $('t2_open').value = Number(prev.trays.t2.totalNotes)||0;
      $('t3_open').value = Number(prev.trays.t3.totalNotes)||0;
      calculate(); saveDraft();
      notify(`Opening carried from ${prev.date}`, 'success');
    }

    /* ================= NUMBER TO WORDS & WHATSAPP SHARE ================= */
    function numToWordsEN(num) {
      num = Math.floor(Number(num) || 0);
      if (num === 0) return 'Zero';
      const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
                    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
      const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
      function convertBelowThousand(n) {
        let str = '';
        if (n >= 100) {
          str += ones[Math.floor(n / 100)] + ' Hundred ';
          n %= 100;
        }
        if (n >= 20) {
          str += tens[Math.floor(n / 10)] + (n % 10 !== 0 ? ' ' + ones[n % 10] : '');
        } else if (n > 0) {
          str += ones[n];
        }
        return str.trim();
      }
      let words = '';
      const crore = Math.floor(num / 10000000);
      num %= 10000000;
      const lakh = Math.floor(num / 100000);
      num %= 100000;
      const thousand = Math.floor(num / 1000);
      num %= 1000;
      const remainder = num;
      if (crore > 0) words += convertBelowThousand(crore) + ' Crore ';
      if (lakh > 0) words += convertBelowThousand(lakh) + ' Lakh ';
      if (thousand > 0) words += convertBelowThousand(thousand) + ' Thousand ';
      if (remainder > 0) words += convertBelowThousand(remainder);
      return words.trim();
    }

    function numToWordsTA(num) {
      num = Math.floor(Number(num) || 0);
      if (num === 0) return 'பூஜ்ஜியம்';
      const onesExact = ['', 'ஒன்று', 'இரண்டு', 'மூன்று', 'நான்கு', 'ஐந்து', 'ஆறு', 'ஏழு', 'எட்டு', 'ஒன்பது'];
      const onesPrefix = ['', 'ஒரு', 'இரண்டு', 'மூன்று', 'நான்கு', 'ஐந்து', 'ஆறு', 'ஏழு', 'எட்டு', 'ஒன்பது'];
      const teens = ['பத்து', 'பதினொன்று', 'பன்னிரண்டு', 'பதிமூன்று', 'பதினான்கு', 'பதினைந்து', 'பதினாறு', 'பதினேழு', 'பதினெட்டு', 'பத்தொன்பது'];
      const tensExact = ['', '', 'இருபது', 'முப்பது', 'நாற்பது', 'ஐம்பது', 'அறுபது', 'எழுபது', 'எண்பது', 'தொண்ணூறு'];
      const tensPrefix = ['', '', 'இருபத்து', 'முப்பத்து', 'நாற்பத்து', 'ஐம்பத்து', 'அறுபத்து', 'எழுபத்து', 'எண்பத்து', 'தொண்ணூற்று'];
      const hundredsExact = ['', 'நூறு', 'இருநூறு', 'முந்நூறு', 'நானூறு', 'ஐந்நூறு', 'ஆறு நூறு', 'ஏழு நூறு', 'எட்டு நூறு', 'தொள்ளாயிரம்'];
      const hundredsPrefix = ['', 'நூற்று', 'இருநூற்று', 'முந்நூற்று', 'நானூற்று', 'ஐந்நூற்று', 'ஆறு நூற்று', 'ஏழு நூற்று', 'எட்டு நூற்று', 'தொள்ளாயிரத்து'];

      function convertSub1000(n, isTerminal) {
        let parts = [];
        const h = Math.floor(n / 100);
        const rem = n % 100;
        if (h > 0) {
          parts.push(rem === 0 ? hundredsExact[h] : hundredsPrefix[h]);
        }
        if (rem > 0) {
          if (rem < 10) {
            parts.push(isTerminal ? onesExact[rem] : onesPrefix[rem]);
          } else if (rem < 20) {
            parts.push(teens[rem - 10]);
          } else {
            const t = Math.floor(rem / 10);
            const u = rem % 10;
            if (u === 0) {
              parts.push(tensExact[t]);
            } else {
              parts.push(tensPrefix[t] + ' ' + (isTerminal ? onesExact[u] : onesPrefix[u]));
            }
          }
        }
        return parts.join(' ');
      }

      let crore = Math.floor(num / 10000000);
      num %= 10000000;
      let lakh = Math.floor(num / 100000);
      num %= 100000;
      let thousand = Math.floor(num / 1000);
      num %= 1000;
      let rem = num;
      let parts = [];

      if (crore > 0) {
        const isTerm = (lakh === 0 && thousand === 0 && rem === 0);
        const cStr = (crore === 1 ? 'ஒரு' : convertSub1000(crore, false));
        parts.push(cStr + ' ' + (isTerm ? 'கோடி' : 'கோடியே'));
      }
      if (lakh > 0) {
        const isTerm = (thousand === 0 && rem === 0);
        const lStr = (lakh === 1 ? 'ஒரு' : convertSub1000(lakh, false));
        parts.push(lStr + ' ' + (isTerm ? 'இலட்சம்' : 'இலட்சத்து'));
      }
      if (thousand > 0) {
        const isTerm = (rem === 0);
        const tStr = (thousand === 1 ? 'ஒரு' : convertSub1000(thousand, false));
        parts.push(tStr + ' ' + (isTerm ? 'ஆயிரம்' : 'ஆயிரத்து'));
      }
      if (rem > 0) {
        parts.push(convertSub1000(rem, true));
      }
      return parts.join(' ').replace(/\s+/g, ' ').trim();
    }

    function formatWhatsAppDate(dateStr) {
      let d = dateStr ? new Date(dateStr + 'T00:00:00') : new Date();
      if (isNaN(d.getTime())) d = new Date();
      const day = d.getDate();
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
      const month = months[d.getMonth()];
      const year = d.getFullYear();
      return `${day} ${month} ${year}`;
    }

    function formatWhatsAppTime() {
      const d = new Date();
      let hours = d.getHours();
      const minutes = String(d.getMinutes()).padStart(2, '0');
      const ampm = hours >= 12 ? 'pm' : 'am';
      hours = hours % 12;
      hours = hours ? String(hours).padStart(2, '0') : '12';
      return `${hours}:${minutes} ${ampm}`;
    }

    function formEntryData(){
      const g = v => Number($(v).value)||0;
      const mk = (o,a) => { const total = g(o)+g(a); return {opening:g(o), added:g(a), totalNotes:total, totalValue:total*(o==='t1_open'?D1:o==='t2_open'?D2:D3)}; };
      return {
        date: $('date').value,
        atmIds: [$('atm_select').value].filter(Boolean),
        bankIndentNo: $('bank_indent_no').value,
        bankIndentVal: g('bank_indent_val'),
        trays: { t1:mk('t1_open','t1_added'), t2:mk('t2_open','t2_added'), t3:mk('t3_open','t3_added') }
      };
    }

    function buildWhatsAppText(e){
      const fmt = n => '₹' + Number(n||0).toLocaleString('en-IN');
      const fmtVal = n => '₹ ' + Number(n||0).toLocaleString('en-IN');
      const dDate = formatWhatsAppDate(e.date);
      const dTime = formatWhatsAppTime();

      const items = [
        { denom: D1, added: Number(e.trays?.t1?.added)||0 },
        { denom: D2, added: Number(e.trays?.t2?.added)||0 },
        { denom: D3, added: Number(e.trays?.t3?.added)||0 }
      ];
      items.sort((a, b) => b.denom - a.denom);

      let totalAmount = 0;
      const denomLines = [];
      items.forEach(item => {
        const val = item.added * item.denom;
        totalAmount += val;
        denomLines.push(`₹${item.denom}  x  ${item.added}  =  ${fmt(val)}`);
      });

      const wordsEn = numToWordsEN(totalAmount);
      const wordsTa = numToWordsTA(totalAmount);

      return [
        '💰 DENOMINATION REPORT',
        `📅 ${dDate} | 🕒 ${dTime}`,
        '──────────────────',
        ...denomLines,
        '──────────────────',
        `TOTAL: ${fmtVal(totalAmount)}`,
        '',
        'IN WORDS (EN):',
        `🔤 ${wordsEn} Rupees Only`,
        `🇮🇳 TAMIL: ${wordsTa} ரூபாய் மட்டும்`,
        '',
        `FINAL AMOUNT: ${fmtVal(totalAmount)}`
      ].join('\n');
    }

    function sendWhatsApp(idOrObj){
      const e = typeof idOrObj === 'string' ? globalData.find(x => x.id === idOrObj) : (idOrObj || formEntryData());
      if(!e || !e.date || !e.bankIndentNo) return notify('Fill the entry details first', 'error');
      let wa = '';
      try { wa = (JSON.parse(localStorage.getItem('ctm_v53')||'{}').wa)||''; } catch(err){ /* ignore malformed settings */ }
      wa = String(wa).replace(/[^\d]/g,'');
      window.open('https://wa.me/' + wa + '?text=' + encodeURIComponent(buildWhatsAppText(e)), '_blank');
    }

    /* ================= 2-PAGE PRINT LOGIC ================= */
    async function handleDualPagePrint(){
        const entryId = $('merge_entry_select').value;
        const pdfFile = $('pdf_input').files[0];
        
        if(!entryId) return notify('Select an entry first', 'error');
        if(!pdfFile) return notify('Upload Indent PDF', 'error');

        $('loader').style.display='flex';
        try {
            // Clear stale state from a previous run
            $('print_pdf_canvas').width = 0; $('print_pdf_canvas').height = 0;
            $('combined-slip-overlay').innerHTML = '';

            // PAGE 1: PDF Background
            const pdfBytes = await pdfFile.arrayBuffer();
            const pdf = await pdfjsLib.getDocument(pdfBytes).promise;
            const page = await pdf.getPage(1);
            const scale = 2.0; 
            const viewport = page.getViewport({scale});
            const canvas = $('print_pdf_canvas');
            canvas.height = viewport.height; canvas.width = viewport.width;
            await page.render({canvasContext: canvas.getContext('2d'), viewport}).promise;

            // PAGE 1: Slip Overlay
            $('combined-slip-overlay').innerHTML = generateSlipHtml(entryId);

            // PAGE 2: Image (File or Camera)
            const imgSrc = $('preview_thumb').src;
            const p2 = $('print-page-2');
            
            if(imgSrc && imgSrc.startsWith('data:')) {
                $('print-page2-img').src = imgSrc;
                p2.style.display = 'flex';
            } else {
                p2.style.display = 'none'; // Skip page 2 if no image
            }

            // Print
            document.body.classList.add('printing-mode', 'combined');
            window.print();
            setTimeout(() => document.body.classList.remove('printing-mode', 'combined'), 1000);

        } catch(e){ console.error(e); notify('Print Generation Failed', 'error'); } 
        finally { $('loader').style.display='none'; }
    }

    /* ================= DOWNLOAD PDF LOGIC (NEW) ================= */
    async function generateDownloadablePDF() {
        const entryId = $('merge_entry_select').value;
        const pdfFile = $('pdf_input').files[0];
        if(!entryId) return notify('Select an entry first', 'error');
        if(!pdfFile) return notify('Upload Indent PDF', 'error');

        $('loader').style.display='flex';
        
        try {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('p', 'mm', 'a4');
            const pageWidth = 210; 
            const pageHeight = 297;

            // 1. Convert Uploaded Indent PDF Page 1 to Image
            const pdfBytes = await pdfFile.arrayBuffer();
            const pdf = await pdfjsLib.getDocument(pdfBytes).promise;
            const page = await pdf.getPage(1);
            const viewport = page.getViewport({ scale: 2 });
            const canvas = document.createElement('canvas');
            canvas.height = viewport.height;
            canvas.width = viewport.width;
            await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
            
            // Add Indent PDF Image to Doc
            doc.addImage(canvas.toDataURL('image/jpeg', 0.8), 'JPEG', 0, 0, pageWidth, pageHeight);

            // 2. Generate Slip Image using html2canvas
            const tempDiv = $('temp-slip-container');
            tempDiv.style.display = 'block';
            tempDiv.innerHTML = generateSlipHtml(entryId);
            // Styling for capture to look good
            tempDiv.style.width = "750px"; 
            tempDiv.style.background = "white"; 
            tempDiv.style.padding = "20px";
            tempDiv.style.borderTop = "2px dashed black";
            
            const slipCanvas = await html2canvas(tempDiv, { scale: 2 });
            tempDiv.style.display = 'none'; // Hide again
            
            // Calculate aspect ratio to fit width
            const slipImgData = slipCanvas.toDataURL('image/png');
            let slipHeight = (slipCanvas.height * pageWidth) / slipCanvas.width;
            const maxSlip = pageHeight * 0.45; // Clamp so the slip never overflows the page
            if(slipHeight > maxSlip){ slipHeight = maxSlip; }
            
            // Add Slip at Bottom
            doc.addImage(slipImgData, 'PNG', 0, pageHeight - slipHeight, pageWidth, slipHeight);

            // 3. Page 2 (Scan Image) if exists
            const p2Img = $('preview_thumb').src;
            if(p2Img && p2Img.startsWith('data:')) {
                doc.addPage();
                const imgFmt = p2Img.startsWith('data:image/png') ? 'PNG' : 'JPEG';
                doc.addImage(p2Img, imgFmt, 10, 10, pageWidth - 20, pageHeight - 20);
            }

            // Save
            const e = globalData.find(x => x.id === entryId);
            const fname = `Daily_Indent_${e.atmIds[0]}_${e.date}.pdf`;
            doc.save(fname);
            notify('PDF Downloaded Successfully');

        } catch(e) {
            console.error(e);
            notify('Failed to generate PDF', 'error');
        } finally {
            $('loader').style.display='none';
        }
    }

    function generateSlipHtml(id) {
        const e = globalData.find(x => x.id === id); if(!e) return '';
        const t1=e.trays.t1, t2=e.trays.t2, t3=e.trays.t3;
        const openVal = (t1.opening*D1)+(t2.opening*D2)+(t3.opening*D3);
        const addVal = (t1.added*D1)+(t2.added*D2)+(t3.added*D3);
        const noteCnt = Number(t1.totalNotes)+Number(t2.totalNotes)+Number(t3.totalNotes);
        const totVal = cleanNum(t1.totalValue) + cleanNum(t2.totalValue) + cleanNum(t3.totalValue);

        // --- SIGNATURES REMOVED PER REQUEST ---
        return `
          <div style="font-family:'Segoe UI',sans-serif;padding:20px;max-width:100%;">
            <div style="text-align:center;border-bottom:2px solid #000;padding-bottom:5px;margin-bottom:10px;">
              <h1 style="margin:0;font-size:24px;line-height:1.2;">Cash Loading Slip</h1>
              <p style="margin:5px 0 0;color:#666;font-size:14px;">Date: ${e.date}</p>
            </div>
            <div style="display:flex;justify-content:space-between;margin-bottom:10px;font-size:14px;">
              <div><strong>ATM:</strong> ${(e.atmIds||[]).join(', ')}</div>
              <div><strong>Indent:</strong> ${e.bankIndentNo}</div>
              <div><strong>Cash:</strong> ₹${Number(e.bankIndentVal).toLocaleString()}</div>
            </div>
            <table style="width:100%;border-collapse:collapse;margin-bottom:auto;font-size:12px;">
              <thead>
                <tr style="border-bottom:2px solid #000;">
                  <th style="padding:5px;text-align:left;">Tray</th>
                  <th style="padding:5px;text-align:right;">Open</th>
                  <th style="padding:5px;text-align:right;">Add</th>
                  <th style="padding:5px;text-align:right;">Total</th>
                  <th style="padding:5px;text-align:right;">Val(₹)</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style="padding:5px;border-bottom:1px solid #eee;">Tray 1 (₹200)</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t1.opening}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t1.added}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t1.totalNotes}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">₹${cleanNum(t1.totalValue).toLocaleString()}</td>
                </tr>
                <tr>
                  <td style="padding:5px;border-bottom:1px solid #eee;">Tray 2 (₹100)</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t2.opening}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t2.added}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t2.totalNotes}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">₹${cleanNum(t2.totalValue).toLocaleString()}</td>
                </tr>
                <tr>
                  <td style="padding:5px;border-bottom:1px solid #eee;">Tray 3 (₹500)</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t3.opening}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t3.added}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">${t3.totalNotes}</td>
                  <td style="padding:5px;text-align:right;border-bottom:1px solid #eee;">₹${cleanNum(t3.totalValue).toLocaleString()}</td>
                </tr>
                <tr style="font-weight:bold;">
                  <td style="padding:8px 5px;border-top:2px solid #000;">TOTALS</td>
                  <td style="padding:8px 5px;text-align:right;border-top:2px solid #000;">₹${openVal.toLocaleString()}</td>
                  <td style="padding:8px 5px;text-align:right;border-top:2px solid #000;">₹${addVal.toLocaleString()}</td>
                  <td style="padding:8px 5px;text-align:right;border-top:2px solid #000;">${noteCnt}</td>
                  <td style="padding:8px 5px;text-align:right;border-top:2px solid #000;">₹${totVal.toLocaleString()}</td>
                </tr>
              </tbody>
            </table>
            <div style="text-align:center;font-size:10px;color:#999;margin-top:10px;">v9.8</div>
          </div>
        `;
    }

    /* ================= CAMERA LOGIC ================= */
    let videoStream;
    async function openCamera() {
        $('camera-modal').style.display = 'flex';
        try {
            videoStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
            $('camera-feed').srcObject = videoStream;
        } catch(e) { notify('Camera access denied', 'error'); closeCamera(); }
    }
    function capturePhoto() {
        const video = $('camera-feed');
        const cvs = document.createElement('canvas');
        cvs.width = video.videoWidth; cvs.height = video.videoHeight;
        cvs.getContext('2d').drawImage(video, 0, 0);
        const dataUrl = cvs.toDataURL('image/jpeg');
        $('preview_thumb').src = dataUrl;
        $('img_preview').style.display='block';
        closeCamera();
    }
    function closeCamera() {
        $('camera-modal').style.display = 'none';
        if(videoStream) videoStream.getTracks().forEach(t => t.stop());
    }

    function renderHistory() {
      const tbody = $('historyBody');
      const filter = $('searchHistory').value.toLowerCase();
      tbody.innerHTML = '';
      globalData.forEach(e => {
        const searchStr = (e.date + e.bankIndentNo + (e.atmIds||[]).join('')).toLowerCase();
        if(filter && !searchStr.includes(filter)) return;
        const t1=e.trays.t1, t2=e.trays.t2, t3=e.trays.t3;
        const openVal = (t1.opening*D1)+(t2.opening*D2)+(t3.opening*D3);
        const addVal = (t1.added*D1)+(t2.added*D2)+(t3.added*D3);
        const noteCnt = Number(t1.totalNotes)+Number(t2.totalNotes)+Number(t3.totalNotes);
        const totVal = cleanNum(t1.totalValue) + cleanNum(t2.totalValue) + cleanNum(t3.totalValue);

        tbody.innerHTML += `
          <tr>
            <td style="font-weight:500;">${e.date}</td><td>${(e.atmIds||[]).join(', ')}</td><td>${e.bankIndentNo}</td> 
            <td class="text-right">₹${Number(e.bankIndentVal).toLocaleString()}</td>
            <td style="color:var(--gray);">Tray 1 (200)</td><td class="text-right">${t1.opening}</td><td class="text-right">${t1.added}</td>
            <td class="text-right">${t1.totalNotes}</td><td class="text-right">₹${cleanNum(t1.totalValue).toLocaleString()}</td>
            <td rowspan="4" class="text-center" style="vertical-align:middle;border-left:1px solid #f1f5f9;">
               <button onclick="editEntry('${e.id}')" style="cursor:pointer;border:none;background:none;color:var(--warning);font-size:16px;margin-right:8px;"><i class="fas fa-edit"></i></button>
               <button onclick="deleteEntry('${e.id}')" style="cursor:pointer;border:none;background:none;color:var(--danger);font-size:16px;margin-right:8px;"><i class="fas fa-trash"></i></button>
               <button onclick="sendWhatsApp('${e.id}')" style="cursor:pointer;border:none;background:none;color:#25D366;font-size:16px;margin-right:8px;"><i class="fab fa-whatsapp"></i></button>
               <div style="margin-top:8px;"><button onclick="printEntry('${e.id}')" style="cursor:pointer;border:none;background:var(--primary);color:white;padding:6px 10px;border-radius:6px;font-size:12px;"><i class="fas fa-print"></i> Slip</button></div>
            </td>
          </tr>
          <tr><td colspan="4" style="border:none;"></td><td style="color:var(--gray);">Tray 2 (100)</td><td class="text-right">${t2.opening}</td><td class="text-right">${t2.added}</td><td class="text-right">${t2.totalNotes}</td><td class="text-right">₹${cleanNum(t2.totalValue).toLocaleString()}</td></tr>
          <tr><td colspan="4" style="border:none;"></td><td style="color:var(--gray);">Tray 3 (500)</td><td class="text-right">${t3.opening}</td><td class="text-right">${t3.added}</td><td class="text-right">${t3.totalNotes}</td><td class="text-right">₹${cleanNum(t3.totalValue).toLocaleString()}</td></tr>
          <tr class="row-total"><td colspan="4" class="text-right">TOTALS</td><td></td><td class="text-right">₹${openVal.toLocaleString()}</td><td class="text-right">₹${addVal.toLocaleString()}</td><td class="text-right">${noteCnt}</td><td class="text-right">₹${totVal.toLocaleString()}</td></tr>
        `;
      });
    }

    function printEntry(id) {
        const slipHtml = generateSlipHtml(id);
        $('single-print-container').innerHTML = slipHtml;
        document.body.classList.add('printing-mode', 'slip');
        window.print();
        setTimeout(() => { document.body.classList.remove('printing-mode', 'slip'); }, 1000);
    }

    function editEntry(id){
      const e = globalData.find(x => x.id === id); if(!e) return;
      isEditingId = id;
      $('cancelEditBtn').style.display='inline-flex';
      $('date').value = e.date; $('bank_indent_no').value = e.bankIndentNo; $('bank_indent_val').value = e.bankIndentVal;
      const atms = e.atmIds||[];
      if(atms.length > 0) $('atm_select').value = atms[0];
      const set = (id, val) => $(id).value = val;
      set('t1_open', e.trays.t1.opening); set('t1_added', e.trays.t1.added);
      set('t2_open', e.trays.t2.opening); set('t2_added', e.trays.t2.added);
      set('t3_open', e.trays.t3.opening); set('t3_added', e.trays.t3.added);
      calculate();
      document.querySelector('[data-tab="entry"]').click();
      notify('Edit Mode Enabled', 'info');
    }

    async function deleteEntry(id){
      if(confirm('Delete this entry permanently?')) {
        showLoader(true); try{
          // Delete via GET (matches the deployed backend; Code.gs v9.8 supports this too).
          const res = await fetch(`${GAS_URL}?action=delete&token=${SECRET_TOKEN}&id=${encodeURIComponent(id)}`);
          const json = await res.json();
          if(json.status!=='ok') throw new Error(json.message||'Delete failed');
          await refreshData(); notify('Entry Deleted', 'success');
        } catch(e){notify('Delete Failed','error');} finally{showLoader(false);}
      }
    }

    function setRange(type){
      const d = new Date(); const fmt = (date) => new Date(date.getTime() - (date.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
      $('rep_to').value = fmt(d);
      if(type==='today') $('rep_from').value = fmt(d);
      if(type==='week') { d.setDate(d.getDate() - 7); $('rep_from').value = fmt(d); }
      if(type==='month') { d.setDate(1); $('rep_from').value = fmt(d); }
    }

    function generateReport(reportType) {
      const from = $('rep_from').value, to = $('rep_to').value, atm = $('rep_atm').value;
      const filtered = globalData.filter(e => {
        if (e.date < from || e.date > to) return false;
        if (atm) { return (e.atmIds||[]).map(x=>x.toUpperCase()).includes(atm.toUpperCase()); }
        return true;
      });

      if(filtered.length === 0) {
        $('reportOutput').style.display='block'; $('reportBody').innerHTML = '<tr><td colspan="9" class="text-center" style="padding:30px;color:var(--danger);">No records found.</td></tr>'; return;
      }
      $('reportOutput').style.display = 'block'; $('rep_period').textContent = `Period: ${from} to ${to}`;

      let g_add=0, g_indent=0;
      let s200_c=0, s200_v=0, s100_c=0, s100_v=0, s500_c=0, s500_v=0; 
      
      const tbody = $('reportBody'); tbody.innerHTML = '';
      const thead = $('reportTableHead');

      if (reportType === 'detailed') {
        $('rep_main_title').textContent = "Detailed Cash Loading Report";
        thead.innerHTML = `<tr>
            <th>Date</th><th>ATM ID</th><th>Indent No.</th><th class="text-right">Indent (₹)</th>
            <th>Tray</th><th class="text-right">Opening</th><th class="text-right">Added</th><th class="text-right">Total Notes</th><th class="text-right">Value (₹)</th>
          </tr>`;
        
        filtered.forEach(e => {
          const t1=e.trays.t1, t2=e.trays.t2, t3=e.trays.t3;
          const openVal = (t1.opening*D1)+(t2.opening*D2)+(t3.opening*D3);
          const addVal = (t1.added*D1)+(t2.added*D2)+(t3.added*D3);
          const noteCnt = Number(t1.totalNotes)+Number(t2.totalNotes)+Number(t3.totalNotes);
          const totVal = cleanNum(t1.totalValue) + cleanNum(t2.totalValue) + cleanNum(t3.totalValue);
          
          // --- FIXED STRATEGY CALCULATION (ADDED NOTES ONLY) ---
          s200_c += Number(t1.added); s200_v += (Number(t1.added)*D1);
          s100_c += Number(t2.added); s100_v += (Number(t2.added)*D2);
          s500_c += Number(t3.added); s500_v += (Number(t3.added)*D3);

          g_add+=addVal; g_indent += Number(e.bankIndentVal);

          const loaded = addVal; 
          const variance = Number(e.bankIndentVal) - loaded;
          const varStyle = variance !== 0 ? 'color:var(--danger);font-weight:700;' : 'color:var(--success);font-weight:600;';
          const varText = variance === 0 ? 'OK' : (variance > 0 ? `Short: ₹${variance}` : `Excess: ₹${Math.abs(variance)}`);

          tbody.innerHTML += `
            <tr>
              <td>${e.date}</td><td>${(e.atmIds||[]).join(',')}</td><td>${e.bankIndentNo}</td>
              <td class="text-right">₹${Number(e.bankIndentVal).toLocaleString()}</td>
              <td>Tray 1 (200)</td><td class="text-right">${t1.opening}</td><td class="text-right">${t1.added}</td><td class="text-right">${t1.totalNotes}</td><td class="text-right">₹${cleanNum(t1.totalValue).toLocaleString()}</td>
            </tr>
            <tr><td colspan="4"></td><td>Tray 2 (100)</td><td class="text-right">${t2.opening}</td><td class="text-right">${t2.added}</td><td class="text-right">${t2.totalNotes}</td><td class="text-right">₹${cleanNum(t2.totalValue).toLocaleString()}</td></tr>
            <tr><td colspan="4"></td><td>Tray 3 (500)</td><td class="text-right">${t3.opening}</td><td class="text-right">${t3.added}</td><td class="text-right">${t3.totalNotes}</td><td class="text-right">₹${cleanNum(t3.totalValue).toLocaleString()}</td></tr>
            <tr class="row-total" style="border-bottom:2px solid #ccc">
               <td colspan="4" class="text-right" style="${varStyle}">Diff: ${varText}</td><td></td>
               <td class="text-right">₹${openVal.toLocaleString()}</td><td class="text-right">₹${addVal.toLocaleString()}</td>
               <td class="text-right">${noteCnt}</td><td class="text-right">₹${totVal.toLocaleString()}</td>
            </tr>
          `;
        });

      } else if (reportType === 'monthly') {
        $('rep_main_title').textContent = "Monthly Summary Report";
        thead.innerHTML = `<tr><th>Month</th><th>Total Transactions</th><th class="text-right">Total Indent (₹)</th><th class="text-right">Total Loaded (₹)</th><th class="text-right">Variance (₹)</th></tr>`;

        const monthlyData = {};
        filtered.forEach(e => {
            const month = e.date.substring(0, 7);
            if(!monthlyData[month]) monthlyData[month] = { indent:0, loaded:0, trans:0 };
            const loaded = (e.trays.t1.added*D1) + (e.trays.t2.added*D2) + (e.trays.t3.added*D3);
            monthlyData[month].indent += Number(e.bankIndentVal);
            monthlyData[month].loaded += loaded;
            monthlyData[month].trans++;
            g_indent += Number(e.bankIndentVal); g_add += loaded; 
            
            // --- FIXED STRATEGY CALCULATION (ADDED NOTES ONLY) ---
            s200_c += Number(e.trays.t1.added); s200_v += (Number(e.trays.t1.added)*D1);
            s100_c += Number(e.trays.t2.added); s100_v += (Number(e.trays.t2.added)*D2);
            s500_c += Number(e.trays.t3.added); s500_v += (Number(e.trays.t3.added)*D3);
        });

        for (const [month, data] of Object.entries(monthlyData)) {
            const variance = data.indent - data.loaded;
            const varStyle = variance !== 0 ? 'color:var(--danger);font-weight:bold;' : 'color:var(--success);';
            tbody.innerHTML += `<tr><td style="font-weight:700;">${month}</td><td class="text-center">${data.trans}</td><td class="text-right">₹${data.indent.toLocaleString()}</td><td class="text-right">₹${data.loaded.toLocaleString()}</td><td class="text-right" style="${varStyle}">₹${variance.toLocaleString()}</td></tr>`;
        }
      }
      
      const g_diff = g_indent - g_add;
      $('rep_sum_indent').textContent = '₹' + g_indent.toLocaleString();
      $('rep_sum_loaded').textContent = '₹' + g_add.toLocaleString();
      $('rep_sum_diff').textContent = '₹' + g_diff.toLocaleString();
      $('rep_sum_diff').style.color = g_diff !== 0 ? '#ef4444' : '#10b981';
      $('rep_sum_200').textContent = s200_c.toLocaleString(); $('rep_val_200').textContent = 'Value: ₹' + s200_v.toLocaleString();
      $('rep_sum_100').textContent = s100_c.toLocaleString(); $('rep_val_100').textContent = 'Value: ₹' + s100_v.toLocaleString();
      $('rep_sum_500').textContent = s500_c.toLocaleString(); $('rep_val_500').textContent = 'Value: ₹' + s500_v.toLocaleString();
      notify('Report Generated', 'success');
    }
    
    function exportToExcel() {
       let html = document.getElementById("reportTable").outerHTML;
       html = html.replace(/<th/g, '<th style="border:1px solid #000"').replace(/<td/g, '<td style="border:1px solid #ccc"');
       const blob = new Blob(['\ufeff', html], { type: 'application/vnd.ms-excel' });
       const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'Cash_Report.xls'; a.click();
    }

    /* ================= MULTI-DAY INDENT CALCULATOR ================= */
    function renderMultiList(){
      const tbody = $('multiBody');
      if(!tbody) return;
      const selected = new Set([...tbody.querySelectorAll('input:checked')].map(cb=>String(cb.value)));
      tbody.innerHTML = '';
      const sorted = [...globalData].sort((a,b)=> (b.date||'').localeCompare(a.date||''));
      sorted.forEach(e => {
        const t1 = Number(e.trays?.t1?.added)||0;
        const t2 = Number(e.trays?.t2?.added)||0;
        const t3 = Number(e.trays?.t3?.added)||0;
        const loaded = (t1*D1)+(t2*D2)+(t3*D3);
        const addTotal = t1+t2+t3;
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><input type="checkbox" value="${e.id}" ${selected.has(String(e.id))?'checked':''} onchange="updateMultiSummary()"></td>
          <td>${e.date}</td><td>${(e.atmIds||[]).join(', ')}</td><td>${e.bankIndentNo||'-'}</td>
          <td class="text-right">₹${Number(e.bankIndentVal||0).toLocaleString('en-IN')}</td>
          <td class="text-right">₹${loaded.toLocaleString('en-IN')}</td>
          <td class="text-right">${t1}</td>
          <td class="text-right">${t2}</td>
          <td class="text-right">${t3}</td>
          <td class="text-right" style="font-weight:600;">${addTotal}</td>`;
        tbody.appendChild(tr);
      });
      updateMultiSummary();
    }

    function multiSelectedEntries(){
      return [...document.querySelectorAll('#multiBody input:checked')]
        .map(cb => globalData.find(x => String(x.id) === String(cb.value)))
        .filter(Boolean);
    }

    function updateMultiSummary(){
      const wrap = $('multiSummary');
      const list = multiSelectedEntries();
      if(!list.length){ wrap.style.display='none'; return; }
      wrap.style.display='block';
      let indent=0, loaded=0, n200=0, n100=0, n500=0, notes=0, val=0;
      list.forEach(e => {
        const t1 = Number(e.trays?.t1?.added)||0;
        const t2 = Number(e.trays?.t2?.added)||0;
        const t3 = Number(e.trays?.t3?.added)||0;
        indent += Number(e.bankIndentVal)||0;
        loaded += (t1*D1)+(t2*D2)+(t3*D3);
        n200 += t1; n100 += t2; n500 += t3;
        notes += (t1 + t2 + t3);
        val += (t1*D1)+(t2*D2)+(t3*D3);
      });
      const variance = indent - loaded;
      $('ms_days').textContent = list.length;
      $('ms_indent').textContent = '₹'+indent.toLocaleString('en-IN');
      $('ms_loaded').textContent = '₹'+loaded.toLocaleString('en-IN');
      $('ms_variance').textContent = '₹'+variance.toLocaleString('en-IN');
      $('ms_variance').style.color = variance===0 ? '#16a34a' : '#ef4444';
      $('ms_n200').textContent = n200.toLocaleString('en-IN'); $('ms_v200').textContent = 'Value: ₹'+(n200*D1).toLocaleString('en-IN');
      $('ms_n100').textContent = n100.toLocaleString('en-IN'); $('ms_v100').textContent = 'Value: ₹'+(n100*D2).toLocaleString('en-IN');
      $('ms_n500').textContent = n500.toLocaleString('en-IN'); $('ms_v500').textContent = 'Value: ₹'+(n500*D3).toLocaleString('en-IN');
      $('ms_tnotes').textContent = notes.toLocaleString('en-IN'); $('ms_tval').textContent = 'Value: ₹'+val.toLocaleString('en-IN');
    }

    function multiSelectAll(){
      document.querySelectorAll('#multiBody input').forEach(cb => cb.checked = true);
      updateMultiSummary();
    }
    function multiClear(){
      document.querySelectorAll('#multiBody input').forEach(cb => cb.checked = false);
      updateMultiSummary();
    }

    function sendWhatsAppMulti(){
      const list = multiSelectedEntries();
      if(!list.length) return notify('Select at least one day in the Multi-Day tab', 'error');
      const fmt = n => '₹' + Number(n||0).toLocaleString('en-IN');
      const fmtVal = n => '₹ ' + Number(n||0).toLocaleString('en-IN');
      const dTime = formatWhatsAppTime();
      const dDate = formatWhatsAppDate();

      const sortedDays = [...list].sort((a,b) => (a.date||'').localeCompare(b.date||''));

      const denoms = [
        { key: 't1', denom: D1 },
        { key: 't2', denom: D2 },
        { key: 't3', denom: D3 }
      ];
      denoms.sort((a, b) => b.denom - a.denom);

      let totalAmount = 0;
      const denomLines = [];

      denoms.forEach(dObj => {
        const dayCounts = sortedDays.map(e => Number(e.trays?.[dObj.key]?.added) || 0);
        const sumNotes = dayCounts.reduce((a, b) => a + b, 0);
        const lineVal = sumNotes * dObj.denom;
        totalAmount += lineVal;

        let calcStr = '';
        if (sortedDays.length > 1) {
          calcStr = `${dayCounts.join(' + ')} (${sumNotes} Notes)`;
        } else {
          calcStr = `${sumNotes}`;
        }

        denomLines.push(`₹${dObj.denom}  x  ${calcStr}  =  ${fmt(lineVal)}`);
      });

      const wordsEn = numToWordsEN(totalAmount);
      const wordsTa = numToWordsTA(totalAmount);

      const lines = [
        '💰 MULTI-DAY DENOMINATION REPORT',
        `📅 ${dDate} | 🕒 ${dTime}`,
        `📊 Selected: ${sortedDays.length} Days`,
        '──────────────────',
        ...denomLines,
        '──────────────────',
        `TOTAL: ${fmtVal(totalAmount)}`,
        '',
        'IN WORDS (EN):',
        `🔤 ${wordsEn} Rupees Only`,
        `🇮🇳 TAMIL: ${wordsTa} ரூபாய் மட்டும்`,
        '',
        `FINAL AMOUNT: ${fmtVal(totalAmount)}`
      ];

      let wa = '';
      try { wa = (JSON.parse(localStorage.getItem('ctm_v53')||'{}').wa)||''; } catch(err){ /* ignore malformed settings */ }
      wa = String(wa).replace(/[^\d]/g,'');
      window.open('https://wa.me/' + wa + '?text=' + encodeURIComponent(lines.join('\n')), '_blank');
    }

    /* ================= FIXED ANALYTICS (ADDED NOTES ONLY) ================= */
    function renderAnalytics() {
      const dates={}, counts={t1:0,t2:0,t3:0}, atmData={};
      let totalCash=0, totalEntries=0;
      
      globalData.forEach(e => {
        // 1. Calculate Loaded Value
        const loaded = cleanNum(e.totalLoadedVal) || cleanNum(e.summary.totalLoadedVal);
        
        // 2. Map Dates for Trend Chart
        dates[e.date]=(dates[e.date]||0)+loaded;

        // 3. Count ONLY Added Notes (Fixed Logic)
        counts.t1 += Number(e.trays.t1.added); 
        counts.t2 += Number(e.trays.t2.added); 
        counts.t3 += Number(e.trays.t3.added);

        // 4. Map ATM Performance
        (e.atmIds||[]).forEach(id => { atmData[id]=(atmData[id]||0)+loaded; });
        
        totalCash+=loaded; totalEntries++;
      });

      // Update Top Stats
      $('an_total_cash').textContent = '₹'+totalCash.toLocaleString();
      $('an_total_entries').textContent = totalEntries;
      $('an_avg_cash').textContent = totalEntries>0 ? '₹'+Math.round(totalCash/totalEntries).toLocaleString() : '₹0';

      // 1. Trend Chart (Line)
      const sortedDates=Object.keys(dates).sort();
      if(window.chart1) window.chart1.destroy();
      window.chart1 = new Chart($('trendChart'),{
          type:'line',
          data:{
              labels:sortedDates,
              datasets:[{
                  label:'Daily Loading (₹)',
                  data:sortedDates.map(d=>dates[d]),
                  borderColor:'#4f46e5',
                  backgroundColor:'rgba(79, 70, 229, 0.1)',
                  tension:0.3,
                  fill:true
              }]
          }
      });
      
      // 2. ATM Chart (Bar)
      const sortedATMs=Object.keys(atmData).sort((a,b)=>atmData[b]-atmData[a]);
      if(window.chartAtm) window.chartAtm.destroy();
      window.chartAtm = new Chart($('atmChart'),{
          type:'bar',
          data:{
              labels:sortedATMs,
              datasets:[{
                  label:'Total Cash Loaded',
                  data:sortedATMs.map(k=>atmData[k]),
                  backgroundColor:'#0ea5e9',
                  borderRadius:6
              }]
          },
          options:{indexAxis:'y'}
      });

      // 3. Denomination Mix Chart (Pie) - Now reflects Added Notes only
      if(window.chart2) window.chart2.destroy();
      window.chart2 = new Chart($('denomChart'),{
          type:'pie',
          data:{
              labels:['₹200','₹100','₹500'],
              datasets:[{
                  data:[counts.t1,counts.t2,counts.t3],
                  backgroundColor:['#4f46e5','#10b981','#f59e0b']
              }]
          }
      });
    }

    function loadSettings(){
      const s=JSON.parse(localStorage.getItem('ctm_v53')||'{}');
      $('set_sheet').value=s.sheet||DEF_SHEET_ID;
      $('set_wa').value=s.wa||'';
      atmList=s.atms||['FZIS0005'];
      // Migrate the old placeholder default to the real ATM id
      if(atmList.length===1 && atmList[0]==='ATM001') atmList=['FZIS0005'];
      if(s.denoms){
        D1=Number(s.denoms.d1)||200; D2=Number(s.denoms.d2)||100; D3=Number(s.denoms.d3)||500;
        $('set_d1').value=D1; $('set_d2').value=D2; $('set_d3').value=D3;
      }
      updateDenomLabels();
      renderAtms();
    }
    function saveSettings(){
      const d1=Math.max(1,Math.round(Number($('set_d1').value)||200));
      const d2=Math.max(1,Math.round(Number($('set_d2').value)||100));
      const d3=Math.max(1,Math.round(Number($('set_d3').value)||500));
      D1=d1; D2=d2; D3=d3; updateDenomLabels();
      localStorage.setItem('ctm_v53',JSON.stringify({sheet:$('set_sheet').value,atms:atmList,denoms:{d1,d2,d3},wa:$('set_wa').value.trim()}));
      renderAtms(); calculate(); notify('Configuration Saved');
    }
    function updateDenomLabels(){
      $('d1_label').textContent='₹'+D1; $('d2_label').textContent='₹'+D2; $('d3_label').textContent='₹'+D3;
    }
    function renderAtms(){
      const div=$('atm_list'), sel=$('atm_select'), rsel=$('rep_atm');
      div.innerHTML='';
      const curr=sel.value;
      const rCurr=rsel.value;
      sel.innerHTML='<option value="">-- Select --</option>'; rsel.innerHTML='<option value="">All</option>';
      atmList.forEach(a=>{
        const chip=document.createElement('span'); chip.style.cssText="background:#f1f5f9;padding:6px 12px;border-radius:20px;font-size:12px;display:inline-flex;align-items:center;gap:8px;border:1px solid #e2e8f0;font-weight:500;";
        chip.innerHTML=`${a} <i class="fas fa-times" style="cursor:pointer;color:#ef4444;"></i>`;
        chip.querySelector('i').onclick=()=>window.remAtm(a); div.appendChild(chip);
        sel.add(new Option(a,a)); rsel.add(new Option(a,a));
      });
      if(atmList.includes(curr)) sel.value=curr;
      else if(atmList.length>0) sel.value=atmList[0];
      if(atmList.includes(rCurr)) rsel.value=rCurr;
    }
    function addAtm(){ const v=$('new_atm').value.trim().toUpperCase(); if(v&&!atmList.includes(v)){atmList.push(v);$('new_atm').value='';saveSettings();} }
    window.remAtm=v=>{if(confirm('Remove ATM?')){atmList=atmList.filter(x=>x!==v);saveSettings();}}

    // --- NEW FUNCTION TO OPEN SHEET ---
    function openSheet() {
      const id = $('set_sheet').value;
      if(!id) return notify('No Sheet ID found', 'error');
      // Opens the standard Google Sheets edit URL for the given ID
      window.open(`https://docs.google.com/spreadsheets/d/${id}/edit`, '_blank');
    }
