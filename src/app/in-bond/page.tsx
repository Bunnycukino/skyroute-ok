'use client';

import React, { Suspense, useState } from 'react';
import { useSearchParams, useRouter } from 'next/navigation';
import * as XLSX from 'xlsx';

type Direction = 'inbound' | 'outbound';

interface FormState {
  direction: Direction;
  c209_number: string;
  bar_number: string;
  pieces: string;
  flight_number: string;
  container_code: string;
  date_received: string;
  lock_seal_check: string;
  c209_present: string;
  recorded_on_despatch_sheet: string;
  section1_comments: string;
  section1_print_name: string;
  section1_sign_name: string;
  section2_comments: string;
  core_bar_locks_checked_prior: string;
  core_bar_locks_intact: string;
  core_bar_seal_match: string;
  core_bar_print_name: string;
  core_bar_sign_name: string;
  gift_cart_locks_checked_prior: string;
  gift_cart_locks_intact: string;
  gift_cart_seal_match: string;
  gift_cart_print_name: string;
  gift_cart_sign_name: string;
  section3_comments: string;
  manager_informed: string;
  manager_name: string;
  reseal_seal_numbers: string;
  reseal_from: string;
  reseal_to: string;
  core_bar_equipment_doors_locks: string;
  core_bar_equipment_wheels_brakes: string;
  core_bar_completion_comments: string;
  core_bar_completion_print_name: string;
  core_bar_completion_sign_name: string;
  gift_cart_equipment_doors_locks: string;
  gift_cart_equipment_wheels_brakes: string;
  gift_cart_completion_comments: string;
  gift_cart_completion_print_name: string;
  gift_cart_completion_sign_name: string;
  dispatch_date: string;
  dispatch_time: string;
  dispatch_recorded: string;
  dispatch_print_name: string;
  dispatch_sign_name: string;
  notes: string;
}

const todayISO = () => new Date().toISOString().slice(0, 10);

const emptyForm = (direction: Direction): FormState => ({
  direction,
  c209_number: '',
  bar_number: '',
  pieces: '',
  flight_number: '',
  container_code: '',
  date_received: todayISO(),
  lock_seal_check: '',
  c209_present: '',
  recorded_on_despatch_sheet: '',
  section1_comments: '',
  section1_print_name: '',
  section1_sign_name: '',
  section2_comments: '',
  core_bar_locks_checked_prior: '',
  core_bar_locks_intact: '',
  core_bar_seal_match: '',
  core_bar_print_name: '',
  core_bar_sign_name: '',
  gift_cart_locks_checked_prior: '',
  gift_cart_locks_intact: '',
  gift_cart_seal_match: '',
  gift_cart_print_name: '',
  gift_cart_sign_name: '',
  section3_comments: '',
  manager_informed: '',
  manager_name: '',
  reseal_seal_numbers: '',
  reseal_from: '',
  reseal_to: '',
  core_bar_equipment_doors_locks: '',
  core_bar_equipment_wheels_brakes: '',
  core_bar_completion_comments: '',
  core_bar_completion_print_name: '',
  core_bar_completion_sign_name: '',
  gift_cart_equipment_doors_locks: '',
  gift_cart_equipment_wheels_brakes: '',
  gift_cart_completion_comments: '',
  gift_cart_completion_print_name: '',
  gift_cart_completion_sign_name: '',
  dispatch_date: todayISO(),
  dispatch_time: '',
  dispatch_recorded: '',
  dispatch_print_name: '',
  dispatch_sign_name: '',
  notes: '',
});

function fmtDate(d: string) {
  if (!d) return '';
  const [y, m, day] = d.split('-');
  if (!y || !m || !day) return d;
  return `${day}/${m}/${y}`;
}

function YesNo({
  value,
  onChange,
  name,
  editable,
}: {
  value: string;
  onChange?: (v: string) => void;
  name: string;
  editable: boolean;
}) {
  if (!editable) {
    return (
      <span className="ib-yn-print">
        <span className={value === 'YES' ? 'ib-yn-selected' : ''}>YES</span>
        {' / '}
        <span className={value === 'NO' ? 'ib-yn-selected' : ''}>NO</span>
      </span>
    );

  }
  return (
    <select
      className="ib-select"
      name={name}
      value={value}
      onChange={(e) => onChange && onChange(e.target.value)}
    >
      <option value=""></option>
      <option value="YES">YES</option>
      <option value="NO">NO</option>
    </select>
  );
}

function InBondFormContent() {
  function findValueNear(rows: any[][], label: string): string {
  const target = label.toLowerCase();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const idx = row.findIndex(
      (c) => typeof c === 'string' && c.trim().toLowerCase().startsWith(target)
    );
    if (idx === -1) continue;
    for (let c = idx + 1; c < row.length; c++) {
      const v = row[c];
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        return String(v).trim();
      }
    }
    const below = rows[r + 1]?.[idx];
    if (below !== undefined && below !== null && String(below).trim() !== '') {
      return String(below).trim();
    }
  }
  return '';
}

function parseInBondExcel(rows: any[][]): Partial<FormState> {
  return {
    c209_number: findValueNear(rows, 'C209 Number'),
    bar_number: findValueNear(rows, 'Bar Number'),
    pieces: findValueNear(rows, 'Number of Pieces'),
    date_received: findValueNear(rows, 'Date Received').slice(0, 10),
    section1_comments: findValueNear(rows, 'Comments'),
    section1_print_name: findValueNear(rows, 'PRINT NAME'),
    section1_sign_name: findValueNear(rows, 'SIGN NAME'),
    manager_name: findValueNear(rows, 'Name of MANAGER'),
    reseal_seal_numbers: findValueNear(rows, 'SEAL NUMBERS'),
    reseal_from: findValueNear(rows, 'FROM'),
    reseal_to: findValueNear(rows, 'TO'),
  };
}
  const searchParams = useSearchParams();
  const router = useRouter();

  const initialDirection: Direction = searchParams.get('direction') === 'outbound' ? 'outbound' : 'inbound';
  const [mode, setMode] = useState<'form' | 'print'>('form');
  const [form, setForm] = useState<FormState>(emptyForm(initialDirection));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [savedId, setSavedId] = useState<number | null>(null);

  const set = (key: keyof FormState) => (v: string) =>
    setForm((prev) => ({ ...prev, [key]: v }));

  const setInput = (key: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement>) =>
    setForm((prev) => ({ ...prev, [key]: e.target.value }));

  async function handleSave() {
    setSaving(true);
    setError('');
    try {
      const payload = {
        ...form,
        pieces: form.pieces ? Number(form.pieces) : null,
        lock_seal_check: form.lock_seal_check === 'YES',
        c209_present: form.c209_present === 'YES',
        recorded_on_despatch_sheet: form.recorded_on_despatch_sheet === 'YES',
        core_bar_locks_checked_prior: form.core_bar_locks_checked_prior === 'YES',
        core_bar_locks_intact: form.core_bar_locks_intact === 'YES',
        core_bar_seal_match: form.core_bar_seal_match === 'YES',
        gift_cart_locks_checked_prior: form.gift_cart_locks_checked_prior === 'YES',
        gift_cart_locks_intact: form.gift_cart_locks_intact === 'YES',
        gift_cart_seal_match: form.gift_cart_seal_match === 'YES',
        manager_informed: form.manager_informed === 'YES',
        core_bar_equipment_doors_locks: form.core_bar_equipment_doors_locks === 'YES',
        core_bar_equipment_wheels_brakes: form.core_bar_equipment_wheels_brakes === 'YES',
        gift_cart_equipment_doors_locks: form.gift_cart_equipment_doors_locks === 'YES',
        gift_cart_equipment_wheels_brakes: form.gift_cart_equipment_wheels_brakes === 'YES',
        dispatch_recorded: form.dispatch_recorded === 'YES',
      };
      function handleExcelImport(e: React.ChangeEvent<HTMLInputElement>) {
  const file = e.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    const data = new Uint8Array(ev.target?.result as ArrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const parsed = parseInBondExcel(rows);
    setForm((prev) => ({ ...prev, ...parsed }));
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
}
      const res = await fetch('/api/in-bond', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || 'Save failed');
      setSavedId(json.data?.id ?? null);
      setMode('print');
    } catch (err: any) {
      setError(err.message || 'Save failed');
    } finally {
      setSaving(false);
    }
  }

  const editable = mode === 'form';

  return (
    <div className="ib-page">
      <div className="ib-toolbar">
        <div className="ib-toolbar-left">
          <button
            className={`ib-tab ${form.direction === 'inbound' ? 'active' : ''}`}
            disabled={!editable}
            onClick={() => setForm((p) => ({ ...p, direction: 'inbound' }))}
          >
            Inbound
          </button>
          <button
            className={`ib-tab ${form.direction === 'outbound' ? 'active' : ''}`}
            disabled={!editable}
            onClick={() => setForm((p) => ({ ...p, direction: 'outbound' }))}
          >
            Outbound
          </button>
        </div>
        <div className="ib-toolbar-right">
          {mode === 'form' ? (
            <button className="ib-btn-primary" disabled={saving} onClick={handleSave}>
              {saving ? 'Saving...' : 'Save & Preview'}
            </button>
          ) : (
            <>
              <button className="ib-btn-secondary" onClick={() => setMode('form')}>Back to Edit</button>
              <button className="ib-btn-primary" onClick={() => window.print()}>Print</button>
              
            </>
          )}
          <button className="ib-btn-secondary" onClick={() => router.push('/in-bond/list')}>View Register</button>
          {editable && (
  abel className="ib-btn-secondary" style={{ cursor: 'pointer' }}>
    Import z Excela
    <input
      type="file"
      accept=".xlsx,.xls"
      style={{ display: 'none' }}
      onChange={handleExcelImport}
    />
  </label>
)}
        </div>
      </div>

      {error && <div className="ib-error">{error}</div>}
      {savedId && <div className="ib-saved-note">Saved as record #{savedId}</div>}

      <div className="ib-sheet">
        <div className="ib-row ib-header-row">
          <div className="ib-col-5 ib-header-logo">
            <div className="ib-logo">dnata</div>
          </div>
          <div className="ib-col-4 ib-header-title">
            <div>IN BOND</div>
            <div>CONTROL SHEET</div>
            <div className="ib-direction-tag">{form.direction.toUpperCase()}</div>
          </div>
          <div className="ib-col-3 ib-header-c209">
            <div className="ib-label-center">C209 Number</div>
            <div className="ib-value-box">
              {editable ? (
                <input className="ib-input ib-input-center" value={form.c209_number} onChange={setInput('c209_number')} />
              ) : (
                form.c209_number
              )}
            </div>
          </div>
        </div>

        <div className="ib-section-title">SECTION 1: {form.direction === 'inbound' ? 'INBOUND BARS' : 'OUTBOUND BARS'}</div>
        <div className="ib-row ib-divide">
          <div className="ib-col-4 ib-cell">
            <label className="ib-label">Bar Number:</label>
            {editable ? (
              <input className="ib-input" value={form.bar_number} onChange={setInput('bar_number')} />
            ) : (
              <div className="ib-value">{form.bar_number}</div>
            )}
          </div>
          <div className="ib-col-3 ib-cell">
            <label className="ib-label">Number of Pieces:</label>
            {editable ? (
              <input className="ib-input" type="number" value={form.pieces} onChange={setInput('pieces')} />
            ) : (
              <div className="ib-value">{form.pieces}</div>
            )}
          </div>
          <div className="ib-col-5 ib-cell">
            <label className="ib-label">{form.direction === 'inbound' ? 'Date Received:' : 'Date Dispatched:'}</label>
            {editable ? (
              <input className="ib-input" type="date" value={form.date_received} onChange={setInput('date_received')} />
            ) : (
              <div className="ib-value">{fmtDate(form.date_received)}</div>
            )}
          </div>
        </div>

        <div className="ib-row ib-divide">
          <div className="ib-col-4 ib-cell">
            <label className="ib-label">Flight Number:</label>
            {editable ? (
              <input className="ib-input" value={form.flight_number} onChange={setInput('flight_number')} />
            ) : (
              <div className="ib-value">{form.flight_number}</div>
            )}
          </div>
          <div className="ib-col-8 ib-cell">
            <label className="ib-label">Container / ULD Code:</label>
            {editable ? (
              <input className="ib-input" value={form.container_code} onChange={setInput('container_code')} />
            ) : (
              <div className="ib-value">{form.container_code}</div>
            )}
          </div>
        </div>

        <div className="ib-row3 ib-divide ib-small">
          <div className="ib-cell ib-inline">
            <span>Lock &amp; Seal Check:</span>
            <YesNo editable={editable} name="lock_seal_check" value={form.lock_seal_check} onChange={set('lock_seal_check')} />
          </div>
          <div className="ib-cell ib-inline">
            <span>C209 Present:</span>
            <YesNo editable={editable} name="c209_present" value={form.c209_present} onChange={set('c209_present')} />
          </div>
          <div className="ib-cell ib-inline">
            <span>Bar Recorded on I/B Despatch:</span>
            <YesNo editable={editable} name="recorded_on_despatch_sheet" value={form.recorded_on_despatch_sheet} onChange={set('recorded_on_despatch_sheet')} />
          </div>
        </div>

        <div className="ib-cell ib-comments">
          <label className="ib-label">Comments:</label>
          {editable ? (
            <textarea className="ib-textarea" value={form.section1_comments} onChange={(e) => set('section1_comments')(e.target.value)} />
          ) : (
            <div className="ib-value">{form.section1_comments}</div>
          )}
        </div>

        <div className="ib-row2 ib-divide">
          <div className="ib-cell ib-inline">
            <label className="ib-label">PRINT NAME:</label>
            {editable ? <input className="ib-input" value={form.section1_print_name} onChange={setInput('section1_print_name')} /> : <span className="ib-value">{form.section1_print_name}</span>}
          </div>
          <div className="ib-cell ib-inline">
            <label className="ib-label">SIGN NAME:</label>
            {editable ? <input className="ib-input" value={form.section1_sign_name} onChange={setInput('section1_sign_name')} /> : <span className="ib-value">{form.section1_sign_name}</span>}
          </div>
        </div>

        <div className="ib-section-title2">
          SECTION 2: BAR STORAGE <span className="ib-italic">to be used for bars that are being stored and/or checked</span>
        </div>
        <div className="ib-cell ib-comments">
          <label className="ib-label">Comments:</label>
          {editable ? (
            <textarea className="ib-textarea" value={form.section2_comments} onChange={(e) => set('section2_comments')(e.target.value)} />
          ) : (
            <div className="ib-value">{form.section2_comments}</div>
          )}
        </div>


        <div className="ib-row2">
          <div className="ib-half">
            <div className="ib-section-title">SECTION 3: BAR PACKING - CORE BAR</div>
            <div className="ib-cell ib-inline-between"><span>Locks &amp; Seals Checked Prior to Opening Bar:</span><YesNo editable={editable} name="core_bar_locks_checked_prior" value={form.core_bar_locks_checked_prior} onChange={set('core_bar_locks_checked_prior')} /></div>
            <div className="ib-cell ib-inline-between"><span>Locks &amp; Seals Intact:</span><YesNo editable={editable} name="core_bar_locks_intact" value={form.core_bar_locks_intact} onChange={set('core_bar_locks_intact')} /></div>
            <div className="ib-cell ib-inline-between"><span>Seal numbers match paperwork?</span><YesNo editable={editable} name="core_bar_seal_match" value={form.core_bar_seal_match} onChange={set('core_bar_seal_match')} /></div>
            <div className="ib-italic-note">* If NO, complete details below &amp; inform Manager/Shift Leader</div>
            <div className="ib-row2">
              <div className="ib-cell ib-inline"><label className="ib-label">PRINT NAME:</label>{editable ? <input className="ib-input" value={form.core_bar_print_name} onChange={setInput('core_bar_print_name')} /> : <span className="ib-value">{form.core_bar_print_name}</span>}</div>
              <div className="ib-cell ib-inline"><label className="ib-label">SIGN NAME:</label>{editable ? <input className="ib-input" value={form.core_bar_sign_name} onChange={setInput('core_bar_sign_name')} /> : <span className="ib-value">{form.core_bar_sign_name}</span>}</div>
            </div>
          </div>
          <div className="ib-half">
            <div className="ib-section-title">BAR PACKING - GIFT CART</div>
            <div className="ib-cell ib-inline-between"><span>Locks &amp; Seals Checked Prior to Opening Bar:</span><YesNo editable={editable} name="gift_cart_locks_checked_prior" value={form.gift_cart_locks_checked_prior} onChange={set('gift_cart_locks_checked_prior')} /></div>
            <div className="ib-cell ib-inline-between"><span>Locks &amp; Seals Intact:</span><YesNo editable={editable} name="gift_cart_locks_intact" value={form.gift_cart_locks_intact} onChange={set('gift_cart_locks_intact')} /></div>
            <div className="ib-cell ib-inline-between"><span>Seal numbers match paperwork?</span><YesNo editable={editable} name="gift_cart_seal_match" value={form.gift_cart_seal_match} onChange={set('gift_cart_seal_match')} /></div>
            <div className="ib-italic-note">&nbsp;</div>
            <div className="ib-row2">
              <div className="ib-cell ib-inline"><label className="ib-label">PRINT NAME:</label>{editable ? <input className="ib-input" value={form.gift_cart_print_name} onChange={setInput('gift_cart_print_name')} /> : <span className="ib-value">{form.gift_cart_print_name}</span>}</div>
              <div className="ib-cell ib-inline"><label className="ib-label">SIGN NAME:</label>{editable ? <input className="ib-input" value={form.gift_cart_sign_name} onChange={setInput('gift_cart_sign_name')} /> : <span className="ib-value">{form.gift_cart_sign_name}</span>}</div>
            </div>
          </div>
        </div>

        <div className="ib-cell ib-comments">
          <label className="ib-label">Comments:</label>
          {editable ? <textarea className="ib-textarea" value={form.section3_comments} onChange={(e) => set('section3_comments')(e.target.value)} /> : <div className="ib-value">{form.section3_comments}</div>}
        </div>

        <div className="ib-row2 ib-divide ib-small">
          <div className="ib-cell ib-inline-between"><span>MANAGER or SHIFT LEADER Informed:</span><YesNo editable={editable} name="manager_informed" value={form.manager_informed} onChange={set('manager_informed')} /></div>
          <div className="ib-cell ib-inline">
            <label className="ib-label">Name of MANAGER/SHIFT LEADER informed:</label>
            {editable ? <input className="ib-input" value={form.manager_name} onChange={setInput('manager_name')} /> : <span className="ib-value">{form.manager_name}</span>}
          </div>
        </div>

        <div className="ib-section-title">
          SECTION 4: RE-SEALED or RE-ALLOCATED BAR <span className="ib-italic">To be completed for Incomplete Bar left by Previous Shift or Bar Re-opened for Bar Check or when bar Re-allocated</span>
        </div>
        <div className="ib-row ib-divide ib-small">
          <div className="ib-col-6 ib-cell">
            <label className="ib-label">SEAL NUMBERS</label>
            {editable ? <input className="ib-input" value={form.reseal_seal_numbers} onChange={setInput('reseal_seal_numbers')} /> : <div className="ib-value">{form.reseal_seal_numbers}</div>}
          </div>
          <div className="ib-col-3 ib-cell">
            <label className="ib-label">FROM</label>
            {editable ? <input className="ib-input" value={form.reseal_from} onChange={setInput('reseal_from')} /> : <div className="ib-value">{form.reseal_from}</div>}
          </div>
          <div className="ib-col-3 ib-cell">
            <label className="ib-label">TO</label>
            {editable ? <input className="ib-input" value={form.reseal_to} onChange={setInput('reseal_to')} /> : <div className="ib-value">{form.reseal_to}</div>}
          </div>
        </div>


        <div className="ib-row2">
          <div className="ib-half">
            <div className="ib-section-title">SECTION 5: BAR COMPLETION - CORE BAR</div>
            <div className="ib-cell ib-inline-between"><span>Equipment Serviceable (Doors &amp; Locks)</span><YesNo editable={editable} name="core_bar_equipment_doors_locks" value={form.core_bar_equipment_doors_locks} onChange={set('core_bar_equipment_doors_locks')} /></div>
            <div className="ib-cell ib-inline-between"><span>Equipment Serviceable (Wheels &amp; Brakes)</span><YesNo editable={editable} name="core_bar_equipment_wheels_brakes" value={form.core_bar_equipment_wheels_brakes} onChange={set('core_bar_equipment_wheels_brakes')} /></div>
            <div className="ib-cell ib-comments">
              <label className="ib-label">Comments:</label>
              {editable ? <textarea className="ib-textarea" value={form.core_bar_completion_comments} onChange={(e) => set('core_bar_completion_comments')(e.target.value)} /> : <div className="ib-value">{form.core_bar_completion_comments}</div>}
            </div>
            <div className="ib-row2">
              <div className="ib-cell ib-inline"><label className="ib-label">PRINT NAME:</label>{editable ? <input className="ib-input" value={form.core_bar_completion_print_name} onChange={setInput('core_bar_completion_print_name')} /> : <span className="ib-value">{form.core_bar_completion_print_name}</span>}</div>
              <div className="ib-cell ib-inline"><label className="ib-label">SIGN NAME:</label>{editable ? <input className="ib-input" value={form.core_bar_completion_sign_name} onChange={setInput('core_bar_completion_sign_name')} /> : <span className="ib-value">{form.core_bar_completion_sign_name}</span>}</div>
            </div>
          </div>
          <div className="ib-half">
            <div className="ib-section-title">SECTION 5: BAR COMPLETION - GIFT CART</div>
            <div className="ib-cell ib-inline-between"><span>Equipment Serviceable (Doors &amp; Locks)</span><YesNo editable={editable} name="gift_cart_equipment_doors_locks" value={form.gift_cart_equipment_doors_locks} onChange={set('gift_cart_equipment_doors_locks')} /></div>
            <div className="ib-cell ib-inline-between"><span>Equipment Serviceable (Wheels &amp; Brakes)</span><YesNo editable={editable} name="gift_cart_equipment_wheels_brakes" value={form.gift_cart_equipment_wheels_brakes} onChange={set('gift_cart_equipment_wheels_brakes')} /></div>
            <div className="ib-cell ib-comments">
              <label className="ib-label">Comments:</label>
              {editable ? <textarea className="ib-textarea" value={form.gift_cart_completion_comments} onChange={(e) => set('gift_cart_completion_comments')(e.target.value)} /> : <div className="ib-value">{form.gift_cart_completion_comments}</div>}
            </div>
            <div className="ib-row2">
              <div className="ib-cell ib-inline"><label className="ib-label">PRINT NAME:</label>{editable ? <input className="ib-input" value={form.gift_cart_completion_print_name} onChange={setInput('gift_cart_completion_print_name')} /> : <span className="ib-value">{form.gift_cart_completion_print_name}</span>}</div>
              <div className="ib-cell ib-inline"><label className="ib-label">SIGN NAME:</label>{editable ? <input className="ib-input" value={form.gift_cart_completion_sign_name} onChange={setInput('gift_cart_completion_sign_name')} /> : <span className="ib-value">{form.gift_cart_completion_sign_name}</span>}</div>
            </div>
          </div>
        </div>

        <div className="ib-row2">
          <div className="ib-half">
            <div className="ib-section-title">SECTION 6: RECORD BAR ON DISPATCH SHEET</div>
            <div className="ib-row2">
              <div className="ib-cell ib-inline"><label className="ib-label">PRINT NAME:</label>{editable ? <input className="ib-input" value={form.dispatch_print_name} onChange={setInput('dispatch_print_name')} /> : <span className="ib-value">{form.dispatch_print_name}</span>}</div>
              <div className="ib-cell ib-inline"><label className="ib-label">SIGN NAME:</label>{editable ? <input className="ib-input" value={form.dispatch_sign_name} onChange={setInput('dispatch_sign_name')} /> : <span className="ib-value">{form.dispatch_sign_name}</span>}</div>
            </div>
          </div>
          <div className="ib-half">
            <div className="ib-row2 ib-divide">
              <div className="ib-cell ib-inline">
                <label className="ib-label">Date:</label>
                {editable ? <input className="ib-input" type="date" value={form.dispatch_date} onChange={setInput('dispatch_date')} /> : <span className="ib-value">{fmtDate(form.dispatch_date)}</span>}
              </div>
              <div className="ib-cell ib-inline">
                <label className="ib-label">Time:</label>
                {editable ? <input className="ib-input" type="time" value={form.dispatch_time} onChange={setInput('dispatch_time')} /> : <span className="ib-value">{form.dispatch_time}</span>}
              </div>
            </div>
            <div className="ib-cell ib-inline-between">
              <span>Bar Details Entered on Despatch Sheet</span>
              <YesNo editable={editable} name="dispatch_recorded" value={form.dispatch_recorded} onChange={set('dispatch_recorded')} />
            </div>
          </div>
        </div>

        <div className="ib-cell ib-comments">
          <label className="ib-label">Notes:</label>
          {editable ? <textarea className="ib-textarea" value={form.notes} onChange={(e) => set('notes')(e.target.value)} /> : <div className="ib-value">{form.notes}</div>}
        </div>

        <div className="ib-footer">In Bond Control Sheet Template v1.2 250124 &mdash; SkyRoute.uk digital edition</div>
      </div>

      <style jsx global>{ibStyles}</style>
    </div>
  );
}

export default function InBondPage() {
  return (
    <Suspense fallback={<div className="ib-loading">Loading Official Template...</div>}>
      <InBondFormContent />
    </Suspense>
  );
}

const ibStyles = `
  * { box-sizing: border-box; }
  body { background: #f3f4f6; }
  .ib-page { max-width: 850px; margin: 0 auto; padding: 16px; font-family: Arial, Helvetica, sans-serif; color: #000; }
  .ib-toolbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; gap: 8px; flex-wrap: wrap; }
  .ib-toolbar-left, .ib-toolbar-right { display: flex; gap: 8px; align-items: center; }
  .ib-tab { padding: 6px 16px; border: 1px solid #1e3a8a; background: #fff; color: #1e3a8a; font-weight: 700; border-radius: 4px; cursor: pointer; }
  .ib-tab.active { background: #1e3a8a; color: #fff; }
  .ib-tab:disabled { opacity: 0.6; cursor: not-allowed; }
  .ib-btn-primary { background: #2563eb; color: #fff; border: none; padding: 8px 18px; border-radius: 4px; font-weight: 700; cursor: pointer; }
  .ib-btn-secondary { background: #fff; color: #1f2937; border: 1px solid #9ca3af; padding: 8px 14px; border-radius: 4px; font-weight: 600; cursor: pointer; }
  .ib-error { background: #fee2e2; color: #991b1b; padding: 8px 12px; border-radius: 4px; margin-bottom: 10px; font-size: 13px; }
  .ib-saved-note { background: #dcfce7; color: #166534; padding: 8px 12px; border-radius: 4px; margin-bottom: 10px; font-size: 13px; }
  .ib-loading { padding: 40px; text-align: center; font-weight: 700; }

  .ib-sheet { background: #fff; border: 2px solid #000; width: 100%; }
  .ib-row { display: grid; grid-template-columns: repeat(12, 1fr); border-bottom: 2px solid #000; }
  .ib-row.ib-divide > div { border-right: 1px solid #000; }
  .ib-row.ib-divide > div:last-child { border-right: none; }
  .ib-row3 { display: grid; grid-template-columns: 1fr 1fr 1fr; border-bottom: 2px solid #000; }
  .ib-row3.ib-divide > div { border-right: 1px solid #000; }
  .ib-row3.ib-divide > div:last-child { border-right: none; }
  .ib-row2 { display: grid; grid-template-columns: 1fr 1fr; border-bottom: 2px solid #000; }
  .ib-row2.ib-divide > div { border-right: 1px solid #000; }
  .ib-row2.ib-divide > div:last-child { border-right: none; }
  .ib-half { border-right: 2px solid #000; }
  .ib-half:last-child { border-right: none; }
  .ib-col-3 { grid-column: span 3; }
  .ib-col-4 { grid-column: span 4; }
  .ib-col-5 { grid-column: span 5; }
  .ib-col-6 { grid-column: span 6; }
  .ib-col-8 { grid-column: span 8; }

  .ib-header-row { border-bottom: 2px solid #000; }
  .ib-header-logo { padding: 14px; display: flex; align-items: center; border-right: 2px solid #000; }
  .ib-logo { font-size: 32px; font-weight: 900; font-style: italic; color: #0057b8; letter-spacing: -1px; }
  .ib-header-title { padding: 10px; border-right: 2px solid #000; font-size: 17px; font-weight: 900; line-height: 1.3; }
  .ib-direction-tag { font-size: 11px; margin-top: 4px; background: #000; color: #fff; display: inline-block; padding: 2px 8px; letter-spacing: 1px; }
  .ib-header-c209 { display: flex; flex-direction: column; }
  .ib-label-center { border-bottom: 2px solid #000; padding: 4px; font-size: 10px; font-weight: 700; text-align: center; }
  .ib-value-box { padding: 8px; text-align: center; font-size: 18px; font-weight: 900; flex: 1; display: flex; align-items: center; justify-content: center; }

  .ib-section-title { background: #000; color: #fff; padding: 4px 8px; font-size: 11px; font-weight: 900; letter-spacing: 1px; }
  .ib-section-title2 { background: #e2e8f0; padding: 4px 8px; font-size: 11px; font-weight: 900; border-bottom: 2px solid #000; }
  .ib-italic { font-weight: 400; font-style: italic; margin-left: 10px; font-size: 10px; }
  .ib-italic-note { padding: 2px 8px; font-style: italic; font-size: 9px; border-bottom: 1px solid #000; }

  .ib-cell { padding: 6px 8px; min-height: 28px; }
  .ib-comments { border-bottom: 2px solid #000; min-height: 48px; }
  .ib-label { font-size: 10px; font-weight: 700; text-transform: uppercase; display: block; margin-bottom: 2px; }
  .ib-value { font-size: 12px; font-weight: 700; min-height: 16px; }
  .ib-input { width: 100%; border: none; border-bottom: 1px solid #94a3b8; font-size: 12px; font-weight: 700; padding: 2px 0; background: transparent; }
  .ib-input:focus { outline: none; border-bottom: 1px solid #2563eb; }
  .ib-input-center { text-align: center; font-size: 16px; }
  .ib-textarea { width: 100%; border: 1px solid #cbd5e1; font-size: 12px; padding: 4px; min-height: 36px; resize: vertical; font-family: inherit; }
  .ib-select { font-size: 10px; font-weight: 700; border: 1px solid #94a3b8; background: #fff; }
  .ib-inline { display: flex; align-items: baseline; gap: 6px; border-bottom: 1px solid #000; }
  .ib-inline .ib-label { margin-bottom: 0; }
  .ib-inline-between { display: flex; align-items: center; justify-content: space-between; font-size: 10px; font-weight: 700; border-bottom: 1px solid #000; }
  .ib-small .ib-cell { padding: 4px 8px; }
  .ib-yn-print { font-size: 10px; font-weight: 700; }
  .ib-yn-selected { text-decoration: underline; background: #fde68a; padding: 0 3px; }
  .ib-footer { text-align: center; padding: 14px; font-size: 11px; color: #64748b; }

  @media print {
    body { background: #fff !important; margin: 0; }
    .ib-toolbar, .ib-error, .ib-saved-note { display: none !important; }
    .ib-page { max-width: 100%; padding: 0; }
    .ib-sheet { border: 2px solid #000; box-shadow: none; }
    .ib-input, .ib-textarea, .ib-select { display: none; }
    @page { size: A4; margin: 10mm; }
  }
`;
