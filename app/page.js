'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { supabase } from '../lib/supabaseClient';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, Label,
} from 'recharts';
import * as XLSX from 'xlsx';

// these mirror the actual limits enforced by handle_tap() in the database -
// keep them in sync if the SQL constants ever change
// used only when an employee's own "Durasi Maks"/"Frekuensi Maks" is null -
// the real limits now live per-employee in the employee table
const DEFAULT_MAX_DURATION_MINUTES = Number(process.env.NEXT_PUBLIC_MAX_DURATION_MINUTES || 30);
const DEFAULT_MAX_FREQUENCY = Number(process.env.NEXT_PUBLIC_MAX_FREQUENCY || 5);

// formats a JS Date as "YYYY-MM-DDTHH:MM:SS" using LOCAL time components
// (matches the plain `timestamp` columns storing Jakarta local time already)
function toLocalTimestamp(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// "YYYY-MM-DD" (native <input type="date"> value) -> local start-of-day Date
function dateInputToDate(str) {
  if (!str) return null;
  const [y, m, d] = str.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function todayInputValue() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function getPeriodRange(period, customFrom, customTo, pickedDate) {
  const now = new Date();
  let start;

  if (period === 'custom') {
    const from = dateInputToDate(customFrom);
    const to = dateInputToDate(customTo);
    if (!from || !to) return null; // incomplete, caller should skip fetching
    const end = new Date(to.getTime());
    end.setDate(end.getDate() + 1); // include the whole "to" day
    return { start: toLocalTimestamp(from), end: toLocalTimestamp(end) };
  }

  if (period === 'date') {
    const picked = dateInputToDate(pickedDate);
    if (!picked) return null; // no date chosen yet, caller should skip fetching
    const end = new Date(picked.getTime());
    end.setDate(end.getDate() + 1);
    return { start: toLocalTimestamp(picked), end: toLocalTimestamp(end) };
  }

  if (period === 'day') {
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  } else if (period === 'week') {
    const day = now.getDay(); // 0 = Sunday
    const diffToMonday = (day === 0 ? -6 : 1) - day;
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMonday);
  } else {
    start = new Date(now.getFullYear(), now.getMonth(), 1);
  }
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1); // through end of today
  return { start: toLocalTimestamp(start), end: toLocalTimestamp(end) };
}

function formatMinutes(mins) {
  if (mins == null) return '-';
  const totalSeconds = Math.round(mins * 60);
  const m = Math.floor(totalSeconds / 60);
  const s = totalSeconds % 60;
  return `${m}m ${s}s`;
}

// builds the export filename from the currently selected period - a single
// date (e.g. "2026-10-13") when the range covers one day, or "start-end"
// when it spans a range
function getExportFilename(period, customFrom, customTo, pickedDate) {
  const range = getPeriodRange(period, customFrom, customTo, pickedDate);
  if (!range) return 'export';

  const fmt = (d) => {
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  };

  const startDate = new Date(range.start);
  // range.end is exclusive (start of the day after the last included day)
  const endInclusive = new Date(new Date(range.end).getTime() - 24 * 60 * 60 * 1000);

  const startStr = fmt(startDate);
  const endStr = fmt(endInclusive);
  return startStr === endStr ? startStr : `${startStr}-${endStr}`;
}

const PERIODS = [
  { key: 'day', label: 'Hari Ini' },
  { key: 'week', label: 'Minggu Ini' },
  { key: 'month', label: 'Bulan Ini' },
];

const THEMES = {
  dark: {
    bodyBg: '#0f1115',
    panel: '#15171c',
    input: '#1a1c22',
    text: '#fff',
    textMuted: '#999',
    textMuted2: '#ccc',
    textDim: '#666',
    textDim2: '#777',
    textFaint: '#888',
    border: '#333',
    borderFaint: '#222',
    borderStrong: '#444',
    borderSubtle: '#292929',
    accent: '#3b82f6',
    danger: '#f87171',
    dangerBg: '#2a1618',
    dangerBorder: '#7f1d1d',
    dangerLight: '#fca5a5',
  },
  light: {
    bodyBg: '#f3f4f6',
    panel: '#ffffff',
    input: '#ffffff',
    text: '#111827',
    textMuted: '#6b7280',
    textMuted2: '#4b5563',
    textDim: '#9ca3af',
    textDim2: '#6b7280',
    textFaint: '#9ca3af',
    border: '#d1d5db',
    borderFaint: '#e5e7eb',
    borderStrong: '#9ca3af',
    borderSubtle: '#e5e7eb',
    accent: '#2563eb',
    danger: '#dc2626',
    dangerBg: '#fef2f2',
    dangerBorder: '#fecaca',
    dangerLight: '#b91c1c',
  },
};

export default function Dashboard() {
  const [theme, setTheme] = useState(() => {
    if (typeof window !== 'undefined') {
      const saved = window.localStorage.getItem('dashboard-theme');
      if (saved === 'light' || saved === 'dark') return saved;
    }
    return 'dark';
  });
  const t = THEMES[theme];

  useEffect(() => {
    window.localStorage.setItem('dashboard-theme', theme);
    document.body.style.background = t.bodyBg;
    document.body.style.color = t.text;
  }, [theme, t]);

  const [activePage, setActivePage] = useState('data'); // 'data' | 'requests' | 'backup'

  const [period, setPeriod] = useState('day');
  const [pickedDate, setPickedDate] = useState('');
  const [customFrom, setCustomFrom] = useState('');
  const [customTo, setCustomTo] = useState('');
  const [stats, setStats] = useState([]);
  const [hourly, setHourly] = useState([]);
  const [sortKey, setSortKey] = useState('total_minutes');
  const [sortDir, setSortDir] = useState('desc');
  const [initialLoading, setInitialLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [search, setSearch] = useState('');
  const [extraFiltersOpen, setExtraFiltersOpen] = useState(false);
  const [departmentFilter, setDepartmentFilter] = useState('');
  const [durasiMin, setDurasiMin] = useState('');
  const [durasiMax, setDurasiMax] = useState('');
  const [rataMin, setRataMin] = useState('');
  const [rataMax, setRataMax] = useState('');
  const [tapMin, setTapMin] = useState('');
  const [tapMax, setTapMax] = useState('');

  const [historyOpen, setHistoryOpen] = useState(false);
  const [overrideHistory, setOverrideHistory] = useState([]);

  const [pendingOverrides, setPendingOverrides] = useState([]);
  const [approvingId, setApprovingId] = useState(null);
  const [pendingError, setPendingError] = useState('');

  const [employeeOptions, setEmployeeOptions] = useState([]);
  const [backupRfids, setBackupRfids] = useState([]);
  const [newBackupEmployeeId, setNewBackupEmployeeId] = useState('');
  const [employeeSearchQuery, setEmployeeSearchQuery] = useState('');
  const [newBackupRfid, setNewBackupRfid] = useState('');
  const [backupError, setBackupError] = useState('');
  const [assigningBackup, setAssigningBackup] = useState(false);
  const [cancelingBackupRfid, setCancelingBackupRfid] = useState(null);

  const periodLabel = period === 'week' ? 'Rata-rata / Minggu' : period === 'month' ? 'Rata-rata / Bulan' : 'Rata-rata / Hari';
  // custom range and single-date both average per-day (group_unit sent to the DB function)
  const groupUnit = (period === 'custom' || period === 'date') ? 'day' : period;

  // guards against an older, slower request overwriting a newer one's result
  const fetchIdRef = useRef(0);
  // debounce handle so a burst of near-simultaneous taps triggers one fetch, not several
  const fetchTimeoutRef = useRef(null);

  function fetchData() {
    const range = getPeriodRange(period, customFrom, customTo, pickedDate);
    if (!range) return; // custom range / picked date not fully/validly entered yet

    const { start, end } = range;
    const requestId = ++fetchIdRef.current;
    setRefreshing(true);

    supabase.rpc('get_dashboard_stats', { period_start: start, period_end: end, group_unit: groupUnit }).then((statsRes) => {
      if (requestId !== fetchIdRef.current) return; // a newer fetch already superseded this one
      if (statsRes.error) console.error('get_dashboard_stats error:', statsRes.error);
      // swap the data in place - the table stays mounted, so this just
      // updates numbers instead of flashing the whole section out and back in
      setStats(statsRes.data || []);
      setInitialLoading(false);
      setRefreshing(false);
    });
  }

  // guards the busiest-hours chart fetch separately from the stats fetch above,
  // since it now runs on its own schedule (whenever the filtered employee set changes)
  const hourlyFetchIdRef = useRef(0);

  function fetchHourly(employeeIds, filtersActive) {
    const range = getPeriodRange(period, customFrom, customTo, pickedDate);
    if (!range) return;
    const { start, end } = range;
    const requestId = ++hourlyFetchIdRef.current;

    supabase.rpc('get_busiest_hours', {
      period_start: start,
      period_end: end,
      // null = no filter (count every session); an array (even empty) scopes
      // the chart to exactly the employees currently matching search/filters
      employee_ids: filtersActive ? employeeIds : null,
    }).then(({ data, error }) => {
      if (requestId !== hourlyFetchIdRef.current) return;
      if (error) console.error('get_busiest_hours error:', error);
      setHourly(fillAllHours(data || []));
    });
  }

  // schedule a fetch after `delay` ms, replacing any not-yet-run fetch already queued
  function scheduleFetch(delay = 0) {
    if (fetchTimeoutRef.current) clearTimeout(fetchTimeoutRef.current);
    fetchTimeoutRef.current = setTimeout(fetchData, delay);
  }

  useEffect(() => {
    scheduleFetch(0); // user-driven change (period/date) - fetch right away
  }, [period, customFrom, customTo, pickedDate]);

  // re-fetch automatically whenever a tap changes the sessions table -
  // no manual refresh needed, the page updates itself in near real time.
  // debounced slightly so several taps arriving close together only
  // trigger one fetch instead of a fetch per event.
  useEffect(() => {
    const channel = supabase
      .channel('sessions-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'sessions' }, () => {
        scheduleFetch(400);
      })
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
      if (fetchTimeoutRef.current) clearTimeout(fetchTimeoutRef.current);
    };
  }, [period, customFrom, customTo, pickedDate]);

  function fetchOverrideHistory() {
    const range = getPeriodRange(period, customFrom, customTo, pickedDate);
    if (!range) return;
    supabase.rpc('get_tap_override_history', { period_start: range.start, period_end: range.end }).then(({ data, error }) => {
      if (error) console.error('get_tap_override_history error:', error);
      setOverrideHistory(data || []);
    });
  }

  // refetch whenever the period changes, and keep it live via realtime
  useEffect(() => {
    fetchOverrideHistory();
    const channel = supabase
      .channel('tap-overrides-history-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'tap_overrides' }, () => {
        fetchOverrideHistory();
      })
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [period, customFrom, customTo, pickedDate]);

  // reads a field off a row regardless of whether Postgres folded the column
  // name to lowercase (unquoted create) or kept it as typed (quoted create)
  function field(row, name) {
    if (row[name] !== undefined) return row[name];
    if (row[name.toLowerCase()] !== undefined) return row[name.toLowerCase()];
    return undefined;
  }

  // pulls every tap_overrides row that hasn't been approved yet, for the
  // "Permintaan Tap Tambahan" dropdown below. Not tied to the selected period -
  // a pending request should show up regardless of which date range is being viewed.
  // select('*') is used instead of naming columns so this doesn't break if the
  // actual DB column casing differs from what's typed here.
  function fetchPendingOverrides() {
    supabase
      .from('tap_overrides')
      .select('*')
      .order('id', { ascending: false })
      .then(({ data, error }) => {
        if (error) {
          console.error('fetch tap_overrides error:', error);
          setPendingError(error.message || 'Gagal memuat permintaan.');
          return;
        }
        setPendingError('');
        // treat null/empty Approval as "not yet approved"; anything else (e.g. "Approved") is excluded
        setPendingOverrides((data || []).filter((r) => !field(r, 'Approval')));
      });
  }

  // independent of the period filters - loads once on mount and stays live via realtime
  useEffect(() => {
    fetchPendingOverrides();
    const channel = supabase
      .channel('tap-overrides-pending-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'tap_overrides' }, () => {
        fetchPendingOverrides();
      })
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, []);

  function handleApprove(id) {
    setApprovingId(id);
    supabase
      .from('tap_overrides')
      .update({ Approval: 'Approved' })
      .eq('id', id)
      .select()
      .then(({ data, error }) => {
        setApprovingId(null);
        if (error) {
          console.error('approve override error:', error);
          setPendingError(error.message || 'Gagal menyetujui permintaan.');
          return;
        }
        // Supabase/Postgres returns success with an empty array (not an error) when
        // a Row Level Security UPDATE policy silently blocks the write - catch that here
        if (!data || data.length === 0) {
          setPendingError('Update tidak tersimpan (0 baris berubah) - kemungkinan RLS policy pada tap_overrides memblokir UPDATE untuk role ini. Cek policy di Supabase.');
          return;
        }
        setPendingError('');
        // realtime subscriptions above will also catch this, but refetch immediately
        // so the UI doesn't wait on the realtime round-trip
        fetchPendingOverrides();
        fetchOverrideHistory();
      });
  }

  // full employee roster (not just those with sessions in the current period) -
  // goes through an RPC (like the rest of the dashboard's employee-data reads)
  // since the employee table's RLS blocks direct anon selects
  function fetchEmployeeOptions() {
    supabase
      .rpc('get_employee_directory')
      .then(({ data, error }) => {
        if (error) { console.error('get_employee_directory error:', error); return; }
        setEmployeeOptions((data || []).map((e) => ({ id: e.employee_id, nama: e.nama })));
      });
  }

  // goes through an RPC, same reasoning as get_employee_directory - backup_rfid
  // has RLS enabled with no SELECT policy yet, which silently returns zero
  // rows to anon rather than an error
  function fetchBackupRfids() {
    supabase
      .rpc('get_backup_rfid_list')
      .then(({ data, error }) => {
        if (error) {
          console.error('get_backup_rfid_list error:', error);
          setBackupError(error.message || 'Gagal memuat RFID backup.');
          return;
        }
        setBackupError('');
        setBackupRfids((data || []).map((b) => ({ RFID: b.rfid, 'Nama Kartu': b.nama_kartu, 'Employee ID': b.employee_id })));
      });
  }

  // independent of the period filters - loads once on mount and stays live via realtime
  useEffect(() => {
    fetchEmployeeOptions();
    fetchBackupRfids();
    const channel = supabase
      .channel('backup-rfid-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'backup_rfid' }, () => {
        fetchBackupRfids();
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'employee' }, () => {
        fetchEmployeeOptions();
      })
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, []);

  // backup_rfid rows are a pool of pre-registered physical cards - "assign"
  // links an existing unassigned card (Employee ID still null) to an
  // employee, it doesn't create a new card row
  function handleAssignBackup() {
    if (!newBackupEmployeeId || !newBackupRfid) return;
    setAssigningBackup(true);
    supabase
      .from('backup_rfid')
      .update({ 'Employee ID': newBackupEmployeeId })
      .eq('RFID', newBackupRfid)
      .is('Employee ID', null)
      .select()
      .then(({ data, error }) => {
        setAssigningBackup(false);
        if (error) {
          console.error('assign backup_rfid error:', error);
          setBackupError(error.message || 'Gagal menetapkan RFID backup.');
          return;
        }
        if (!data || data.length === 0) {
          setBackupError('Tidak tersimpan (0 baris berubah) - kartu mungkin sudah ditetapkan ke orang lain, atau RLS policy pada backup_rfid memblokir UPDATE untuk role ini.');
          return;
        }
        setBackupError('');
        setNewBackupEmployeeId('');
        setEmployeeSearchQuery('');
        setNewBackupRfid('');
        fetchBackupRfids();
      });
  }

  // "cancel" un-links the card from its employee (Employee ID -> null) rather
  // than deleting the card row, so the physical card stays registered and can
  // be reassigned to someone else later
  function handleCancelBackup(rfid) {
    setCancelingBackupRfid(rfid);
    supabase
      .from('backup_rfid')
      .update({ 'Employee ID': null })
      .eq('RFID', rfid)
      .select()
      .then(({ data, error }) => {
        setCancelingBackupRfid(null);
        if (error) {
          console.error('cancel backup_rfid error:', error);
          setBackupError(error.message || 'Gagal membatalkan RFID backup.');
          return;
        }
        // same silent-RLS-block pattern as the approve/update flows above
        if (!data || data.length === 0) {
          setBackupError('Tidak tersimpan (0 baris berubah) - kemungkinan RLS policy pada backup_rfid memblokir UPDATE untuk role ini.');
          return;
        }
        setBackupError('');
        fetchBackupRfids();
      });
  }

  // ensure all 24 hours show on the chart even if some had zero taps
  function fillAllHours(rows) {
    const map = new Map(rows.map((r) => [r.hour, r.tap_count]));
    return Array.from({ length: 24 }, (_, h) => ({ hour: h, tap_count: map.get(h) || 0 }));
  }

  const departmentOptions = useMemo(() => {
    const set = new Set(stats.map((s) => s.departemen).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [stats]);

  // shared by the table (sortedStats) and the busiest-hours chart, so both
  // reflect the same search / department / range filters consistently
  function applyStatFilters(list) {
    let filtered = search.trim()
      ? list.filter((s) => (s.name || '').toLowerCase().includes(search.trim().toLowerCase()))
      : list;

    if (departmentFilter) {
      filtered = filtered.filter((s) => s.departemen === departmentFilter);
    }
    const inRange = (value, min, max) => {
      const v = value || 0;
      if (min !== '' && v < Number(min)) return false;
      if (max !== '' && v > Number(max)) return false;
      return true;
    };
    filtered = filtered.filter((s) =>
      inRange(s.total_minutes, durasiMin, durasiMax) &&
      inRange(s.avg_period_minutes, rataMin, rataMax) &&
      inRange(s.total_taps, tapMin, tapMax)
    );
    return filtered;
  }

  const filtersActive = Boolean(
    search.trim() || departmentFilter || durasiMin !== '' || durasiMax !== ''
    || rataMin !== '' || rataMax !== '' || tapMin !== '' || tapMax !== ''
  );

  // which employees currently match Cari Nama + Filter Lainnya - drives the
  // busiest-hours chart so it scopes down the same way the table does
  const filteredEmployeeIds = useMemo(
    () => applyStatFilters(stats).map((s) => s.employee_id).filter(Boolean),
    [stats, search, departmentFilter, durasiMin, durasiMax, rataMin, rataMax, tapMin, tapMax]
  );

  // refetch the chart whenever the period changes OR the filtered employee
  // set changes (which itself updates whenever stats loads or a filter is edited)
  useEffect(() => {
    fetchHourly(filteredEmployeeIds, filtersActive);
  }, [period, customFrom, customTo, pickedDate, filteredEmployeeIds, filtersActive]);

  const sortedStats = useMemo(() => {
    const filtered = applyStatFilters(stats);

    const copy = [...filtered];
    copy.sort((a, b) => {
      let cmp;
      if (sortKey === 'name' || sortKey === 'departemen') {
        cmp = (a[sortKey] || '').localeCompare(b[sortKey] || '');
      } else {
        cmp = (a[sortKey] || 0) - (b[sortKey] || 0);
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });
    return copy;
  }, [stats, sortKey, sortDir, search, departmentFilter, durasiMin, durasiMax, rataMin, rataMax, tapMin, tapMax]);

  function resetFilters() {
    setDepartmentFilter('');
    setDurasiMin('');
    setDurasiMax('');
    setRataMin('');
    setRataMax('');
    setTapMin('');
    setTapMax('');
  }

  // exports the currently filtered/sorted table (sortedStats already
  // reflects Cari Nama + Filter Lainnya) to an .xlsx file named after the
  // selected period
  function handleExportExcel() {
    const rows = sortedStats.map((row) => ({
      Nama: row.name || '(tanpa nama)',
      Departemen: row.departemen || '-',
      'Total Durasi (menit)': row.total_minutes != null ? Number((row.total_minutes).toFixed(2)) : 0,
      [`${periodLabel} (menit)`]: row.avg_period_minutes != null ? Number((row.avg_period_minutes).toFixed(2)) : 0,
      'Jumlah Tap': row.total_taps ?? 0,
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');

    const filename = getExportFilename(period, customFrom, customTo, pickedDate);
    XLSX.writeFile(workbook, `${filename}.xlsx`);
  }

  function toggleSort(key) {
    if (sortKey === key) {
      setSortDir(sortDir === 'asc' ? 'desc' : 'asc');
    } else {
      setSortKey(key);
      setSortDir(key === 'name' ? 'asc' : 'desc');
    }
  }

  const busiestHour = hourly.reduce((max, h) => (h.tap_count > (max?.tap_count || 0) ? h : max), null);
  const overallAvg = stats.length
    ? stats.reduce((sum, s) => sum + (s.avg_period_minutes || 0), 0) / stats.length
    : 0;
  const overLimitPeople = stats.filter((s) => {
    const maxDuration = s.durasi_maks ?? DEFAULT_MAX_DURATION_MINUTES;
    const maxFrequency = s.frekuensi_maks ?? DEFAULT_MAX_FREQUENCY;
    return (s.max_daily_minutes || 0) >= maxDuration || (s.max_daily_taps || 0) >= maxFrequency;
  });

  return (
    <main style={{ maxWidth: 1000, margin: '0 auto', padding: '24px 16px', boxSizing: 'border-box' }}>
      <style jsx global>{`
        input[type="date"]::-webkit-calendar-picker-indicator {
          filter: ${theme === 'dark' ? 'invert(1)' : 'none'};
          cursor: pointer;
        }
      `}</style>
      <button
        onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
        style={{
          position: 'absolute', top: 16, right: 16,
          padding: '8px 14px', borderRadius: 6, border: `1px solid ${t.border}`,
          background: t.input, color: t.text, cursor: 'pointer', fontSize: 13,
        }}
      >
        {theme === 'dark' ? '☀️ Light Mode' : '🌙 Dark Mode'}
      </button>
      <div>
        <h1 style={{ marginBottom: 4, fontSize: 'clamp(20px, 5vw, 32px)' }}>
          Dashboard Monitoring Akses Toilet
          {refreshing && !initialLoading && (
            <span style={{
              fontSize: 12, fontWeight: 400, color: t.accent, marginLeft: 10,
              verticalAlign: 'middle', opacity: 0.85,
            }}>
              ● memperbarui...
            </span>
          )}
        </h1>
      </div>

      <div style={{
        display: 'inline-block', color: t.textMuted2, fontSize: 14,
        border: `1px solid ${t.border}`, borderRadius: 8, padding: '10px 14px',
        margin: '4px 0 20px', background: t.panel,
      }}>
        Rata-rata durasi keseluruhan: <strong style={{ color: t.text }}>{formatMinutes(overallAvg)}</strong>
        {busiestHour && (
          <> &nbsp;|&nbsp; Jam tersibuk: <strong style={{ color: t.text }}>{busiestHour.hour}:00 - {busiestHour.hour + 1}:00</strong> ({busiestHour.tap_count} tap)</>
        )}
      </div>
      <br />

      <div style={{
        border: `1px solid ${t.border}`, borderRadius: 8, padding: '16px', marginBottom: 20, background: t.panel,
      }}>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 14 }}>
          {PERIODS.map((p) => (
            <button
              key={p.key}
              onClick={() => setPeriod(p.key)}
              style={{
                padding: '8px 16px',
                borderRadius: 6,
                border: `1px solid ${t.border}`,
                background: period === p.key ? t.accent : t.input,
                color: t.text,
                cursor: 'pointer',
              }}
            >
              {p.label}
            </button>
          ))}
        </div>

        <div style={{
          display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 20,
          marginBottom: 14, paddingBottom: 14, borderBottom: `1px solid ${t.borderSubtle}`,
        }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 14 }}>
            <span style={{ color: t.textMuted }}>Pilih Tanggal:</span>
            <input
              type="date"
              value={pickedDate}
              onChange={(e) => { setPickedDate(e.target.value); setPeriod('date'); }}
              style={{ background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px' }}
            />
          </label>

          <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 14, color: t.textMuted }}>Rentang Kustom:</span>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 14 }}>
              Dari
              <input
                type="date"
                value={customFrom}
                onChange={(e) => { setCustomFrom(e.target.value); setPeriod('custom'); }}
                style={{ background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px' }}
              />
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 14 }}>
              Sampai
              <input
                type="date"
                value={customTo}
                onChange={(e) => { setCustomTo(e.target.value); setPeriod('custom'); }}
                style={{ background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px' }}
              />
            </label>
            {period === 'custom' && getPeriodRange('custom', customFrom, customTo) && (
              <span style={{ fontSize: 13, color: t.accent }}>● Menampilkan rentang kustom</span>
            )}
            {period === 'date' && pickedDate && (
              <span style={{ fontSize: 13, color: t.accent }}>● Menampilkan tanggal terpilih</span>
            )}
          </div>
        </div>

        <input
          type="text"
          placeholder="Cari nama..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            width: '100%', maxWidth: 320, boxSizing: 'border-box',
            background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6,
            padding: '8px 12px', fontSize: 14,
          }}
        />
      </div>

      <div style={{
        border: `1px solid ${t.border}`, borderRadius: 8, padding: '14px 16px',
        marginBottom: 20, background: t.panel,
      }}>
        <div
          onClick={() => setExtraFiltersOpen(!extraFiltersOpen)}
          style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}
        >
          <strong style={{ fontSize: 15 }}>Filter Lainnya</strong>
          <span style={{ color: t.textFaint }}>{extraFiltersOpen ? '▲' : '▼'}</span>
        </div>

        {extraFiltersOpen && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'flex-end', marginTop: 14 }}>
          <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: t.textMuted }}>
            Departemen
            <select
              value={departmentFilter}
              onChange={(e) => setDepartmentFilter(e.target.value)}
              style={{
                background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6,
                padding: '6px 8px', fontSize: 14, minWidth: 160,
              }}
            >
              <option value="">Semua Departemen</option>
              {departmentOptions.map((d) => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>
          </label>

          <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: t.textMuted }}>
            Total Durasi (menit)
            <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
              <input
                type="number" placeholder="Min" value={durasiMin}
                onChange={(e) => setDurasiMin(e.target.value)}
                style={{ width: 80, background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 14 }}
              />
              <span style={{ color: t.textDim }}>–</span>
              <input
                type="number" placeholder="Max" value={durasiMax}
                onChange={(e) => setDurasiMax(e.target.value)}
                style={{ width: 80, background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 14 }}
              />
            </div>
          </label>

          <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: t.textMuted }}>
            {periodLabel} (menit)
            <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
              <input
                type="number" placeholder="Min" value={rataMin}
                onChange={(e) => setRataMin(e.target.value)}
                style={{ width: 80, background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 14 }}
              />
              <span style={{ color: t.textDim }}>–</span>
              <input
                type="number" placeholder="Max" value={rataMax}
                onChange={(e) => setRataMax(e.target.value)}
                style={{ width: 80, background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 14 }}
              />
            </div>
          </label>

          <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: t.textMuted }}>
            Jumlah Tap
            <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
              <input
                type="number" placeholder="Min" value={tapMin}
                onChange={(e) => setTapMin(e.target.value)}
                style={{ width: 80, background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 14 }}
              />
              <span style={{ color: t.textDim }}>–</span>
              <input
                type="number" placeholder="Max" value={tapMax}
                onChange={(e) => setTapMax(e.target.value)}
                style={{ width: 80, background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 14 }}
              />
            </div>
          </label>

          <button
            onClick={resetFilters}
            style={{
              padding: '7px 14px', borderRadius: 6, border: `1px solid ${t.borderStrong}`,
              background: 'transparent', color: t.textMuted, cursor: 'pointer', fontSize: 13,
            }}
          >
            Reset Filter
          </button>
        </div>
        )}
      </div>

      <nav style={{ display: 'flex', flexWrap: 'wrap', gap: 8, margin: '4px 0 20px' }}>
        {[
          { key: 'data', label: 'Tabel & Tren' },
          { key: 'requests', label: 'Tap Tambahan' },
          { key: 'backup', label: 'RFID Backup' },
        ].map((p) => (
          <button
            key={p.key}
            onClick={() => setActivePage(p.key)}
            style={{
              padding: '8px 16px', borderRadius: 6,
              border: `1px solid ${activePage === p.key ? t.accent : t.border}`,
              background: activePage === p.key ? t.accent : t.input,
              color: activePage === p.key ? '#fff' : t.text,
              cursor: 'pointer', fontSize: 14, fontWeight: activePage === p.key ? 600 : 400,
            }}
          >
            {p.label}
          </button>
        ))}
      </nav>

      <hr style={{ border: 'none', borderTop: `1px solid ${t.border}`, margin: '0 0 20px' }} />

      {activePage === 'data' && (
      <>
      {initialLoading ? (
        <p>Memuat...</p>
      ) : (
        <>
          <h2 style={{ fontSize: 'clamp(18px, 4vw, 24px)' }}>Jumlah Tap Berdasarkan Jam</h2>
          <div style={{ height: 330, minWidth: 0, marginBottom: 24 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={hourly} margin={{ top: 5, right: 5, left: 25, bottom: 40 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={t.borderFaint} />
                <XAxis
                  dataKey="hour"
                  tickFormatter={(h) => `${h}:00`}
                  stroke={t.textFaint}
                  tick={{ fill: t.text, fontSize: 10 }}
                  interval={0}
                  angle={-60}
                  textAnchor="end"
                  height={50}
                >
                  <Label value="Jam" position="bottom" offset={0} fill={t.text} fontSize={13} />
                </XAxis>
                <YAxis stroke={t.textFaint} tick={{ fill: t.text }} allowDecimals={false} width={30}>
                  <Label value="Jumlah Tap" angle={-90} position="left" offset={15} fill={t.text} fontSize={13} style={{ textAnchor: 'middle' }} />
                </YAxis>
                <Tooltip
                  contentStyle={{ background: t.input, border: `1px solid ${t.border}` }}
                  labelStyle={{ color: t.text }}
                  itemStyle={{ color: t.text }}
                  labelFormatter={(h) => `${h}:00 - ${h + 1}:00`}
                />
                <Bar dataKey="tap_count" radius={[4, 4, 0, 0]}>
                  {hourly.map((entry, i) => (
                    <Cell key={i} fill={busiestHour && entry.hour === busiestHour.hour ? t.danger : t.accent} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 10 }}>
            <button
              onClick={handleExportExcel}
              disabled={sortedStats.length === 0}
              style={{
                padding: '8px 16px', borderRadius: 6, border: 'none', whiteSpace: 'nowrap',
                background: sortedStats.length === 0 ? t.textDim : '#16a34a',
                color: '#fff', fontWeight: 600, fontSize: 13,
                cursor: sortedStats.length === 0 ? 'not-allowed' : 'pointer',
              }}
            >
              Export ke Excel
            </button>
          </div>

          <div style={{ overflowX: 'auto', marginBottom: 24 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 480 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${t.border}`, textAlign: 'left' }}>
                  <SortableHeader label="Nama" sortKey="name" current={sortKey} dir={sortDir} onClick={toggleSort} t={t} />
                  <SortableHeader label="Departemen" sortKey="departemen" current={sortKey} dir={sortDir} onClick={toggleSort} t={t} />
                  <SortableHeader label="Total Durasi" sortKey="total_minutes" current={sortKey} dir={sortDir} onClick={toggleSort} t={t} />
                  <SortableHeader label={periodLabel} sortKey="avg_period_minutes" current={sortKey} dir={sortDir} onClick={toggleSort} t={t} />
                  <SortableHeader label="Jumlah Tap" sortKey="total_taps" current={sortKey} dir={sortDir} onClick={toggleSort} t={t} />
                </tr>
              </thead>
              <tbody>
                {sortedStats.map((row) => (
                  <tr key={row.employee_id} style={{ borderBottom: `1px solid ${t.borderFaint}` }}>
                    <td style={{ padding: '10px 8px' }}>{row.name || '(tanpa nama)'}</td>
                    <td style={{ padding: '10px 8px' }}>{row.departemen || '-'}</td>
                    <td style={{ padding: '10px 8px' }}>{formatMinutes(row.total_minutes)}</td>
                    <td style={{ padding: '10px 8px' }}>{formatMinutes(row.avg_period_minutes)}</td>
                    <td style={{ padding: '10px 8px' }}>{row.total_taps ?? 0}</td>
                  </tr>
                ))}
                {sortedStats.length === 0 && (
                  <tr><td colSpan={5} style={{ padding: 20, textAlign: 'center', color: t.textDim2 }}>
                    {search.trim() ? 'Tidak ada nama yang cocok' : 'Tidak ada sesi dalam periode ini'}
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>

          {(period === 'day' || period === 'date') && overLimitPeople.length > 0 && (
            <div style={{
              background: t.dangerBg, border: `1px solid ${t.dangerBorder}`, borderRadius: 8,
              padding: '14px 16px', marginBottom: 40,
            }}>
              <strong style={{ color: t.danger }}>⚠ Peringatan batas sesi</strong>
              <ul style={{ margin: '8px 0 0', paddingLeft: 20, color: t.dangerLight }}>
                {overLimitPeople.map((p) => {
                  const maxDuration = p.durasi_maks ?? DEFAULT_MAX_DURATION_MINUTES;
                  const maxFrequency = p.frekuensi_maks ?? DEFAULT_MAX_FREQUENCY;
                  const hitDuration = (p.max_daily_minutes || 0) >= maxDuration;
                  const hitFrequency = (p.max_daily_taps || 0) >= maxFrequency;
                  const reasons = [];
                  if (hitDuration) reasons.push(`durasi ${formatMinutes(p.max_daily_minutes)}`);
                  if (hitFrequency) reasons.push(`${p.max_daily_taps} tap`);
                  return (
                    <li key={p.employee_id}>
                      {p.name || '(tanpa nama)'} sudah mencapai batas sesi mereka ({reasons.join(' dan ')})
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
        </>
      )}
      </>
      )}

      {activePage === 'requests' && (
      <>
      <div style={{ margin: '4px 0 24px' }}>
        <h2 style={{ fontSize: 'clamp(18px, 4vw, 24px)', marginBottom: 12 }}>
          Permintaan Tap Tambahan{pendingOverrides.length > 0 && ` (${pendingOverrides.length})`}
        </h2>

        {pendingError && (
          <p style={{ color: t.danger, fontSize: 13, margin: '0 0 10px' }}>
            Error: {pendingError}
          </p>
        )}

        {!pendingError && pendingOverrides.length === 0 ? (
          <p style={{ color: t.textDim2, fontSize: 13, margin: 0 }}>
            Tidak ada permintaan yang menunggu persetujuan.
          </p>
        ) : null}

        {pendingOverrides.length > 0 && (
          <ul style={{ listStyle: 'none', margin: 0, padding: 0 }}>
            {pendingOverrides.map((o) => {
              const id = field(o, 'id');
              const nama = field(o, 'Nama') || field(o, 'EmployeeID') || '(tanpa nama)';
              const jenis = field(o, 'Jenis');
              const jumlah = field(o, 'Jumlah');
              const isApproving = approvingId === id;
              return (
                <li key={id} style={{
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                  padding: '8px 10px', borderBottom: `1px solid ${t.borderFaint}`, fontSize: 14, gap: 8,
                }}>
                  <span>
                    {nama}{' '}
                    <span style={{ color: t.textDim2 }}>
                      — {jenis === 'Durasi' ? `+${jumlah} menit` : `+${jumlah} tap`}
                    </span>
                  </span>
                  <button
                    onClick={() => handleApprove(id)}
                    disabled={isApproving}
                    style={{
                      padding: '6px 14px', borderRadius: 6, border: 'none', whiteSpace: 'nowrap',
                      background: isApproving ? t.textDim : '#16a34a',
                      color: '#fff', fontWeight: 600, fontSize: 13,
                      cursor: isApproving ? 'not-allowed' : 'pointer',
                    }}
                  >
                    {isApproving ? 'Menyetujui...' : 'Approve'}
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </div>
      </>
      )}

      {activePage === 'backup' && (
      <>
      <div style={{ margin: '4px 0 24px' }}>
        <h2 style={{ fontSize: 'clamp(18px, 4vw, 24px)', marginBottom: 12 }}>
            RFID Backup{backupRfids.filter((b) => field(b, 'Employee ID')).length > 0 && ` (${backupRfids.filter((b) => field(b, 'Employee ID')).length})`}
        </h2>

          {backupError && (
            <p style={{ color: t.danger, fontSize: 13, margin: '0 0 10px' }}>
              Error: {backupError}
            </p>
          )}

          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'flex-start', marginBottom: 14 }}>
            <div style={{ position: 'relative', flex: '1 1 220px', minWidth: 220 }}>
              <input
                type="text"
                value={employeeSearchQuery}
                onChange={(e) => {
                  setEmployeeSearchQuery(e.target.value);
                  setNewBackupEmployeeId('');
                }}
                placeholder="Cari nama karyawan..."
                style={{
                  background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6,
                  padding: '8px 10px', fontSize: 14, width: '100%', boxSizing: 'border-box',
                }}
              />
              {employeeSearchQuery.trim() && !newBackupEmployeeId && (() => {
                const q = employeeSearchQuery.trim().toLowerCase();
                const matches = employeeOptions
                  .filter((e) => (e.nama || '').toLowerCase().includes(q))
                  .slice(0, 8);
                return (
                  <ul style={{
                    position: 'absolute', zIndex: 10, top: '100%', left: 0, right: 0, marginTop: 4,
                    background: t.panel, border: `1px solid ${t.border}`, borderRadius: 6,
                    maxHeight: 200, overflowY: 'auto', listStyle: 'none', padding: 4,
                  }}>
                    {matches.length === 0 ? (
                      <li style={{ padding: '8px 10px', fontSize: 13, color: t.textDim2 }}>Tidak ditemukan.</li>
                    ) : (
                      matches.map((e) => (
                        <li
                          key={e.id}
                          onClick={() => {
                            setNewBackupEmployeeId(e.id);
                            setEmployeeSearchQuery(e.nama || e.id);
                          }}
                          style={{ padding: '8px 10px', fontSize: 14, borderRadius: 4, cursor: 'pointer', color: t.text }}
                          onMouseEnter={(ev) => { ev.currentTarget.style.background = t.rowHover || t.border; }}
                          onMouseLeave={(ev) => { ev.currentTarget.style.background = 'transparent'; }}
                        >
                          {e.nama || e.id}
                        </li>
                      ))
                    )}
                  </ul>
                );
              })()}
            </div>
            <div style={{ flex: '1 1 200px', minWidth: 200 }}>
              <select
                value={newBackupRfid}
                onChange={(e) => setNewBackupRfid(e.target.value)}
                style={{
                  background: t.input, color: t.text, border: `1px solid ${t.border}`, borderRadius: 6,
                  padding: '8px 10px', fontSize: 14, width: '100%', boxSizing: 'border-box',
                }}
              >
                <option value="">Pilih kartu RFID backup...</option>
                {backupRfids
                  .filter((b) => !field(b, 'Employee ID'))
                  .map((b) => {
                    const rfid = field(b, 'RFID');
                    const namaKartu = field(b, 'Nama Kartu');
                    return (
                      <option key={rfid} value={rfid}>{namaKartu ? `${namaKartu} (${rfid})` : rfid}</option>
                    );
                  })}
              </select>
              {backupRfids.filter((b) => !field(b, 'Employee ID')).length === 0 && (
                <p style={{ color: t.textDim2, fontSize: 12, margin: '4px 0 0' }}>
                  Tidak ada kartu yang belum ditetapkan.
                </p>
              )}
            </div>
            <button
              onClick={handleAssignBackup}
              disabled={!newBackupEmployeeId || !newBackupRfid || assigningBackup}
              style={{
                padding: '8px 16px', borderRadius: 6, border: 'none', whiteSpace: 'nowrap',
                background: (!newBackupEmployeeId || !newBackupRfid || assigningBackup) ? t.textDim : t.accent,
                color: '#fff', fontWeight: 600, fontSize: 14,
                cursor: (!newBackupEmployeeId || !newBackupRfid || assigningBackup) ? 'not-allowed' : 'pointer',
              }}
            >
              {assigningBackup ? 'Menetapkan...' : 'Assign'}
            </button>
          </div>

          {backupRfids.filter((b) => field(b, 'Employee ID')).length === 0 ? (
            <p style={{ color: t.textDim2, fontSize: 13, margin: 0 }}>
              Belum ada RFID backup yang ditetapkan.
            </p>
          ) : (
            <ul style={{ listStyle: 'none', margin: 0, padding: 0 }}>
              {backupRfids.filter((b) => field(b, 'Employee ID')).map((b) => {
                const rfid = field(b, 'RFID');
                const namaKartu = field(b, 'Nama Kartu');
                const empId = field(b, 'Employee ID');
                const emp = employeeOptions.find((e) => e.id === empId);
                const isCanceling = cancelingBackupRfid === rfid;
                return (
                  <li key={rfid} style={{
                    display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                    padding: '8px 10px', borderBottom: `1px solid ${t.borderFaint}`, fontSize: 14, gap: 8,
                  }}>
                    <span>
                      {(emp && emp.nama) || empId || '(tanpa nama)'}{' '}
                      <span style={{ color: t.textDim2 }}>
                        — {namaKartu ? `${namaKartu} (${rfid})` : `RFID ${rfid}`}
                      </span>
                    </span>
                    <button
                      onClick={() => handleCancelBackup(rfid)}
                      disabled={isCanceling}
                      style={{
                        padding: '6px 14px', borderRadius: 6, border: 'none', whiteSpace: 'nowrap',
                        background: isCanceling ? t.textDim : '#16a34a',
                        color: '#fff', fontWeight: 600, fontSize: 13,
                        cursor: isCanceling ? 'not-allowed' : 'pointer',
                      }}
                    >
                      {isCanceling ? 'Memproses...' : 'Finish'}
                    </button>
                  </li>
                );
              })}
            </ul>
          )}
      </div>
      </>
      )}

      {activePage === 'requests' && (
      <>
      <div style={{
        border: `1px solid ${t.border}`, borderRadius: 8, padding: '14px 16px',
        margin: '0 0 20px', background: t.panel,
      }}>
        <div
          onClick={() => setHistoryOpen(!historyOpen)}
          style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}
        >
          <strong style={{ fontSize: 15, color: t.danger }}>
            Riwayat Tap Tambahan{overrideHistory.length > 0 && ` (${overrideHistory.length})`}
          </strong>
          <span style={{ color: t.textFaint }}>{historyOpen ? '▲' : '▼'}</span>
        </div>

        {historyOpen && (
        <div style={{ marginTop: 12 }}>
          <p style={{ fontSize: 13, color: t.textMuted, margin: '0 0 10px' }}>
            Riwayat izin tap tambahan yang diberikan pada periode yang dipilih di atas ({PERIODS.find((p) => p.key === period)?.label || (period === 'date' ? 'Tanggal Terpilih' : 'Rentang Kustom')}).
          </p>

          {overrideHistory.length === 0 && (
            <p style={{ color: t.textDim2, fontSize: 13 }}>Belum ada izin tap tambahan pada periode ini.</p>
          )}

          {overrideHistory.length > 0 && (
            <ul style={{ listStyle: 'none', margin: 0, padding: 0 }}>
              {overrideHistory.map((o) => (
                <li key={`${o.employee_id}-${o.diberikan}`} style={{
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                  padding: '8px 10px', borderBottom: `1px solid ${t.borderFaint}`, fontSize: 14, gap: 8,
                }}>
                  <span>
                    {o.nama || '(tanpa nama)'}{' '}
                    <span style={{ color: t.textDim2 }}>
                      — diberikan {new Date(o.diberikan).toLocaleString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
                    </span>
                  </span>
                  <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexShrink: 0 }}>
                    <span style={{
                      fontSize: 13, whiteSpace: 'nowrap', padding: '3px 10px', borderRadius: 12,
                      background: t.dangerBg, border: `1px solid ${t.dangerBorder}`, color: t.danger,
                    }}>
                      {o.jenis === 'Durasi' ? `+${o.jumlah} menit` : `+${o.jumlah} tap`}
                    </span>
                    {(o.approval || o.Approval) === 'Approved' && (
                      <span style={{
                        fontSize: 13, whiteSpace: 'nowrap', padding: '3px 10px', borderRadius: 12,
                        background: 'rgba(22,163,74,0.15)', border: '1px solid #16a34a', color: '#16a34a',
                      }}>
                        Approved
                      </span>
                    )}
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
        )}
      </div>
      </>
      )}
    </main>
  );
}

function SortableHeader({ label, sortKey, current, dir, onClick, t }) {
  const active = current === sortKey;
  return (
    <th
      onClick={() => onClick(sortKey)}
      style={{ padding: '10px 8px', cursor: 'pointer', userSelect: 'none', color: active ? t.accent : t.text, whiteSpace: 'nowrap' }}
    >
      {label} {active ? (dir === 'asc' ? '▲' : '▼') : ''}
    </th>
  );
}
