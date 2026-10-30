import React, { useState, useEffect, useRef } from 'react';
import {
  Play,
  Pause,
  RotateCcw,
  Monitor,
  Trophy,
  Volume2,
  Swords,
  Award,
  Maximize2,
  Upload,
  Download,
  FileSpreadsheet,
  ChevronRight,
  CheckCircle2,
  ListOrdered,
  Trash2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { doc, setDoc, onSnapshot } from 'firebase/firestore';

const PENALTIES = ['C1', 'C2', 'C3', 'HC', 'H'];
const STORAGE_KEY = 'kiaido_live_scoreboard_state';
const SCHEDULE_STORAGE_KEY = 'kiaido_tournament_schedule';

const DEFAULT_MATCH_STATE = {
  matchId: '',
  mode: 'kumite', // 'kumite' | 'kata'
  categoryName: 'Kumite Cá Nhân Nam - Lứa tuổi 10-11 (-35kg)',
  akaName: 'NGUYỄN VĂN A',
  akaTeam: 'CLB KIAIDO Q1',
  akaScore: 0,
  akaSenshu: false,
  akaPenalties: { C1: false, C2: false, C3: false, HC: false, H: false },

  aoName: 'TRẦN VĂN B',
  aoTeam: 'CLB KIAIDO Q7',
  aoScore: 0,
  aoSenshu: false,
  aoPenalties: { C1: false, C2: false, C3: false, HC: false, H: false },

  duration: 90,
  timeLeft: 90,
  isRunning: false,
  winner: null, // 'AKA' | 'AO' | null

  kataSystem: 'flags', // 'flags' | 'points'
  akaKataName: 'Heian Shodan',
  aoKataName: 'Heian Nidan',
  akaFlags: 0,
  aoFlags: 0,
  akaJudges: [7.5, 7.5, 7.5, 7.5, 7.5],
  aoJudges: [7.5, 7.5, 7.5, 7.5, 7.5]
  
};

// Hệ thống âm thanh Chuông / Kẻng võ đài chuyên nghiệp (Web Audio API)
const playRealBellMp3 = (durationSec = 1.0) => {
  try {
    const audio = new Audio('/bell.mp3');
    audio.volume = 1.0;
    audio.currentTime = 0;
    audio.play().catch((err) => console.error(err));
    setTimeout(() => {
      audio.pause();
      audio.currentTime = 0;
    }, durationSec * 1000);
  } catch (e) {
    console.error(e);
  }
};

// 1. Bắt đầu trận: Reng 1.2 giây
const playStartGong = () => playRealBellMp3(0.5);

// 2. Còn 15s cuối: 2 tiếng Reng-Reng liên tục
const playWarning15sBell = () => {
  playRealBellMp3(0.2);
  setTimeout(() => playRealBellMp3(0.2), 650);
};

// 3. Hết trận: Reng kéo dài đúng 3 giây
const playEndMatchBell = () => playRealBellMp3(2.0);

const calcKataTotal = (scores) => {
  const nums = (scores || [7.5, 7.5, 7.5, 7.5, 7.5]).map((n) => Number(n) || 0);
  if (nums.length < 5) return nums.reduce((a, b) => a + b, 0).toFixed(1);
  const sorted = [...nums].sort((a, b) => a - b);
  return (sorted[1] + sorted[2] + sorted[3]).toFixed(1);
};

export default function KarateScoreboard({ db, students = [] }) {
  const [match, setMatch] = useState(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) return { ...DEFAULT_MATCH_STATE, ...JSON.parse(saved) };
    } catch (e) {
      console.error(e);
    }
    return DEFAULT_MATCH_STATE;
  });

  // Danh sách các trận đấu / hạng cân được Import từ Excel (Sima)
  const [scheduleList, setScheduleList] = useState(() => {
    try {
      const saved = localStorage.getItem(SCHEDULE_STORAGE_KEY);
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error(e);
    }
    return [];
  });

  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const fileInputRef = useRef(null);
  const channelRef = useRef(null);

  const isDisplayOnly = new URLSearchParams(window.location.search).get('scoreboard') === 'display';

  // Lưu danh sách lịch thi đấu vào localStorage & Firebase
  const saveScheduleList = (newList) => {
    setScheduleList(newList);
    try {
      localStorage.setItem(SCHEDULE_STORAGE_KEY, JSON.stringify(newList));
    } catch (e) {
      console.error(e);
    }
    if (db) {
      setDoc(doc(db, 'scoreboard', 'tournament_schedule'), { matches: newList }).catch(console.error);
    }
  };

  // Kết nối đồng bộ Real-time
  useEffect(() => {
    if ('BroadcastChannel' in window) {
      channelRef.current = new BroadcastChannel('kiaido_scoreboard_channel');
      channelRef.current.onmessage = (event) => {
        if (event.data) setMatch((prev) => ({ ...prev, ...event.data }));
      };
    }

    const handleStorageChange = (e) => {
      if (e.key === STORAGE_KEY && e.newValue) {
        try {
          setMatch((prev) => ({ ...prev, ...JSON.parse(e.newValue) }));
        } catch (err) {
          console.error(err);
        }
      }
    };
    window.addEventListener('storage', handleStorageChange);

    let unsubMatch = null;
    let unsubSchedule = null;
    if (db) {
      unsubMatch = onSnapshot(doc(db, 'scoreboard', 'live_match'), (snap) => {
        if (snap.exists()) setMatch((prev) => ({ ...prev, ...snap.data() }));
      });
      if (!isDisplayOnly) {
        unsubSchedule = onSnapshot(doc(db, 'scoreboard', 'tournament_schedule'), (snap) => {
          if (snap.exists() && Array.isArray(snap.data().matches)) {
            setScheduleList(snap.data().matches);
          }
        });
      }
    }

    return () => {
      if (channelRef.current) channelRef.current.close();
      window.removeEventListener('storage', handleStorageChange);
      if (unsubMatch) unsubMatch();
      if (unsubSchedule) unsubSchedule();
    };
  }, [db, isDisplayOnly]);

  const updateMatch = (updater) => {
    setMatch((prev) => {
      const next = typeof updater === 'function' ? updater(prev) : { ...prev, ...updater };

      if (next.mode === 'kumite') {
        const akaP = next.akaPenalties || {};
        const aoP = next.aoPenalties || {};
        if (akaP.H) next.winner = 'AO';
        else if (aoP.H) next.winner = 'AKA';
        else if (next.akaScore - next.aoScore >= 8) {
          next.winner = 'AKA';
          next.isRunning = false;
        } else if (next.aoScore - next.akaScore >= 8) {
          next.winner = 'AO';
          next.isRunning = false;
        }
      }

      // Nếu trận đấu nằm trong danh sách Sima và đã có người thắng -> Tự cập nhật kết quả vào danh sách Sima
      if (next.matchId && next.winner && scheduleList.length > 0) {
        const updatedSchedule = scheduleList.map((item) => {
          if (item.id === next.matchId) {
            const scoreStr = next.mode === 'kumite'
              ? `${next.akaScore} - ${next.aoScore}`
              : next.kataSystem === 'flags'
              ? `Cờ: ${next.akaFlags} - ${next.aoFlags}`
              : `Điểm: ${calcKataTotal(next.akaJudges)} - ${calcKataTotal(next.aoJudges)}`;
            return {
              ...item,
              status: 'completed',
              winner: next.winner,
              winnerName: next.winner === 'AKA' ? next.akaName : next.aoName,
              finalScore: scoreStr
            };
          }
          return item;
        });
        saveScheduleList(updatedSchedule);
      }

      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      } catch (e) {
        console.error(e);
      }
      if (channelRef.current) channelRef.current.postMessage(next);
      if (db) setDoc(doc(db, 'scoreboard', 'live_match'), next).catch(console.error);

      return next;
    });
  };

  useEffect(() => {
    if (isDisplayOnly || !match.isRunning) return;

    const timer = setInterval(() => {
      updateMatch((prev) => {
        if (!prev.isRunning || prev.timeLeft <= 0) return prev;
        const nextTime = prev.timeLeft - 1;

        // Còn 15 giây cuối: Đánh 2 tiếng chuông liên tục
        if (nextTime === 15) {
          playWarning15sBell();
        }
        // Hết giờ (00:00): Đánh tiếng chuông lớn kéo dài 3 giây
        if (nextTime === 0) {
          playEndMatchBell();
          return { ...prev, timeLeft: 0, isRunning: false };
        }
        return { ...prev, timeLeft: nextTime };
      });
    }, 1000);

    return () => clearInterval(timer);
  }, [match.isRunning, isDisplayOnly]);

  const formatClock = (seconds) => {
    const s = Math.max(0, Number(seconds) || 0);
    const mins = Math.floor(s / 60);
    const secs = s % 60;
    return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  };

  const handleOpenDisplayWindow = () => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(match));
    } catch (e) {
      console.error(e);
    }
    const url = `${window.location.origin}${window.location.pathname}?scoreboard=display`;
    window.open(url, 'KaratePublicDisplay', 'width=1280,height=720');
  };

  const handleResetMatch = () => {
    if (!window.confirm('Đặt lại điểm số, lỗi và thời gian về ban đầu cho trận hiện tại?')) return;
    updateMatch((prev) => ({
      ...DEFAULT_MATCH_STATE,
      matchId: prev.matchId,
      mode: prev.mode,
      categoryName: prev.categoryName,
      akaName: prev.akaName,
      akaTeam: prev.akaTeam,
      aoName: prev.aoName,
      aoTeam: prev.aoTeam,
      duration: prev.duration,
      timeLeft: prev.duration
    }));
  };

  // ============================================================================
  // CHỨC NĂNG SAAS: IMPORT / EXPORT DANH SÁCH HẠNG CÂN & TRẬN ĐẤU (SIMA)
  // ============================================================================
  const handleDownloadSimaTemplate = () => {
    const templateRows = [
      {
        'Hạng Cân / Nội Dung': 'Kumite Cá Nhân Nam 10-11 tuổi (-35kg)',
        'Thể Thức': 'Kumite',
        'Thời Gian (Giây)': 90,
        'VĐV Đỏ (AKA)': 'Nguyễn Văn A',
        'Đơn Vị Đỏ': 'KIAIDO Quận 1',
        'VĐV Xanh (AO)': 'Trần Văn B',
        'Đơn Vị Xanh': 'KIAIDO Quận 7',
        'Bài Kata Đỏ': '',
        'Bài Kata Xanh': ''
      },
      {
        'Hạng Cân / Nội Dung': 'Kumite Cá Nhân Nữ 12-13 tuổi (-42kg)',
        'Thể Thức': 'Kumite',
        'Thời Gian (Giây)': 90,
        'VĐV Đỏ (AKA)': 'Lê Thị C',
        'Đơn Vị Đỏ': 'KIAIDO Tân Bình',
        'VĐV Xanh (AO)': 'Phạm Thị D',
        'Đơn Vị Xanh': 'KIAIDO Quận 1',
        'Bài Kata Đỏ': '',
        'Bài Kata Xanh': ''
      },
      {
        'Hạng Cân / Nội Dung': 'Kata Cá Nhân Nam U10',
        'Thể Thức': 'Kata',
        'Thời Gian (Giây)': 120,
        'VĐV Đỏ (AKA)': 'Hoàng Minh E',
        'Đơn Vị Đỏ': 'KIAIDO Quận 7',
        'VĐV Xanh (AO)': 'Đặng Quốc F',
        'Đơn Vị Xanh': 'KIAIDO Phú Nhuận',
        'Bài Kata Đỏ': 'Heian Shodan',
        'Bài Kata Xanh': 'Heian Nidan'
      }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(templateRows), 'DanhSachThiDau_Sima');
    XLSX.writeFile(wb, 'FileMau_DanhSachHangCan_Sima.xlsx');
  };

  const handleImportSimaExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws);

        // Lọc các dòng hợp lệ trước, sau đó tự động sinh Mã Trận (T01, T02, T03...)
        const validRows = json.filter((row) =>
          Boolean(row['Hạng Cân / Nội Dung'] || row['Hạng Cân'] || row['Nội Dung'] || row['VĐV Đỏ (AKA)'] || row['AKA'])
        );

        const mapped = validRows.map((row, idx) => {
          // Tự động sinh mã trận đấu theo thứ tự T01, T02, T03...
          const generatedId = `T${String(idx + 1).padStart(2, '0')}`;
          const categoryName = String(row['Hạng Cân / Nội Dung'] || row['Hạng Cân'] || row['Nội Dung'] || 'Nội dung thi đấu').trim();
          const rawMode = String(row['Thể Thức'] || row['Môn'] || 'Kumite').trim().toLowerCase();
          const mode = rawMode.includes('kata') ? 'kata' : 'kumite';
          const duration = Number(row['Thời Gian (Giây)'] || row['Thời Gian'] || 90) || 90;

          return {
            id: generatedId,
            categoryName,
            mode,
            duration,
            akaName: String(row['VĐV Đỏ (AKA)'] || row['AKA'] || 'VĐV ĐỎ').trim().toUpperCase(),
            akaTeam: String(row['Đơn Vị Đỏ'] || row['CLB Đỏ'] || '').trim(),
            aoName: String(row['VĐV Xanh (AO)'] || row['AO'] || 'VĐV XANH').trim().toUpperCase(),
            aoTeam: String(row['Đơn Vị Xanh'] || row['CLB Xanh'] || '').trim(),
            akaKataName: String(row['Bài Kata Đỏ'] || 'Heian Shodan').trim(),
            aoKataName: String(row['Bài Kata Xanh'] || 'Heian Nidan').trim(),
            status: 'pending',
            winner: null,
            winnerName: '',
            finalScore: ''
          };
        });

        if (mapped.length === 0) {
          alert('Không tìm thấy dữ liệu hợp lệ trong file Excel!');
          return;
        }

        saveScheduleList(mapped);
        handleLoadMatchFromSchedule(mapped[0]);
      } catch (err) {
        alert('Lỗi đọc file Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = null;
  };

  // Nạp 1 trận đấu từ danh sách Sima lên Bảng Điều Khiển & Màn Hình TV
  const handleLoadMatchFromSchedule = (item) => {
    updateMatch({
      ...DEFAULT_MATCH_STATE,
      matchId: item.id,
      mode: item.mode || 'kumite',
      categoryName: `[Trận ${item.id}] ${item.categoryName}`,
      akaName: item.akaName || 'VĐV ĐỎ',
      akaTeam: item.akaTeam || '',
      aoName: item.aoName || 'VĐV XANH',
      aoTeam: item.aoTeam || '',
      duration: Number(item.duration) || 90,
      timeLeft: Number(item.duration) || 90,
      akaKataName: item.akaKataName || 'Heian Shodan',
      aoKataName: item.aoKataName || 'Heian Nidan',
      isRunning: false,
      winner: item.winner || null
    });
  };

  // Chuyển nhanh sang trận tiếp theo trong danh sách Sima
  const handleNextMatch = () => {
    if (scheduleList.length === 0) return;
    const currentIndex = scheduleList.findIndex((m) => m.id === match.matchId);
    const nextIndex = currentIndex + 1 < scheduleList.length ? currentIndex + 1 : 0;
    handleLoadMatchFromSchedule(scheduleList[nextIndex]);
  };

  // Xuất kết quả toàn bộ giải đấu ra Excel
  const handleExportTournamentResults = () => {
    if (scheduleList.length === 0) return;
    const rows = scheduleList.map((m, i) => ({
      'STT': i + 1,
      'Mã Trận': m.id,
      'Hạng Cân / Nội Dung': m.categoryName,
      'Thể Thức': m.mode === 'kata' ? 'Kata' : 'Kumite',
      'VĐV Đỏ (AKA)': m.akaName,
      'Đơn Vị Đỏ': m.akaTeam,
      'VĐV Xanh (AO)': m.aoName,
      'Đơn Vị Xanh': m.aoTeam,
      'Tỉ Số / Điểm': m.finalScore || 'Chưa đấu',
      'Người Chiến Thắng': m.winnerName ? `${m.winnerName} (${m.winner})` : 'Chưa xác định'
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'KetQuaThiDau');
    XLSX.writeFile(wb, 'KetQuaGiaiDau_Karate.xlsx');
  };

  const uniqueCategories = Array.from(new Set(scheduleList.map((m) => m.categoryName)));
  const filteredSchedule = scheduleList.filter(
    (m) => categoryFilter === 'ALL' || m.categoryName === categoryFilter
  );

  // ============================================================================
  // GIAO DIỆN MÀN HÌNH CHIẾU TV & XEM TRƯỚC
  // ============================================================================
  const renderPublicScoreboard = (isPreview = false) => {
    const isAtoshiBaraku = match.timeLeft <= 15 && match.timeLeft > 0;
    const akaP = match.akaPenalties || {};
    const aoP = match.aoPenalties || {};

    return (
      <div
        className={`${
          isPreview
            ? 'rounded-2xl overflow-hidden border-4 border-slate-800 shadow-lg min-h-[420px]'
            : 'fixed inset-0 z-50 h-screen w-screen'
        } bg-black text-white flex flex-col select-none`}
      >
        <div className="bg-slate-950 border-b border-slate-800 px-4 sm:px-6 py-3 flex items-center justify-between gap-2">
          <div className="flex items-center gap-3 min-w-0">
            <span className="px-2.5 py-1 rounded-lg bg-red-600 text-[11px] font-black uppercase tracking-wider shrink-0">
              KIAIDO
            </span>
            <h2 className={`${isPreview ? 'text-xs sm:text-base' : 'text-lg sm:text-2xl'} font-extrabold text-amber-400 uppercase tracking-wide truncate`}>
              {match.categoryName}
            </h2>
          </div>
          {!isPreview && (
            <button
              onClick={() => {
                if (!document.fullscreenElement) document.documentElement.requestFullscreen();
                else document.exitFullscreen();
              }}
              className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold flex items-center gap-1.5 shrink-0"
            >
              <Maximize2 className="w-4 h-4" /> Toàn màn hình
            </button>
          )}
        </div>

        {match.mode === 'kumite' ? (
          <div className="flex-1 grid grid-cols-2 relative">
            {/* BÊN ĐỎ (AKA) */}
            <div className="bg-gradient-to-br from-red-700 via-red-800 to-red-950 p-4 sm:p-6 flex flex-col justify-between border-r border-black/40">
              <div>
                <div className="flex items-center justify-between gap-2">
                  <span className="px-2.5 py-1 rounded bg-black/40 text-red-200 font-bold text-xs uppercase tracking-widest">
                    AKA (ĐỎ)
                  </span>
                  {match.akaSenshu && (
                    <span className="px-3 py-1 rounded-full bg-yellow-400 text-black font-black text-xs sm:text-sm shadow-lg">
                      ★ SENSHU
                    </span>
                  )}
                </div>
                <h3 className={`${isPreview ? 'text-lg sm:text-2xl' : 'text-2xl sm:text-5xl'} font-black uppercase mt-2 tracking-tight truncate`}>
                  {match.akaName}
                </h3>
                <p className="text-red-200 text-xs sm:text-base font-semibold opacity-90 truncate">{match.akaTeam}</p>
              </div>

              <div className="my-auto text-center py-4">
                <span className={`${isPreview ? 'text-6xl sm:text-8xl' : 'text-8xl sm:text-[170px]'} font-black leading-none drop-shadow-2xl`}>
                  {match.akaScore}
                </span>
              </div>

              <div className="space-y-1.5">
                <p className="text-[10px] sm:text-xs uppercase tracking-wider text-red-200 font-bold">Lỗi (Penalties)</p>
                <div className="grid grid-cols-5 gap-1.5 sm:gap-2">
                  {PENALTIES.map((p) => (
                    <div
                      key={p}
                      className={`py-1.5 sm:py-2.5 rounded-xl text-center font-black text-xs sm:text-base border-2 ${
                        akaP[p]
                          ? p === 'H'
                            ? 'bg-red-500 text-white border-white'
                            : 'bg-yellow-400 text-black border-yellow-200 shadow-lg'
                          : 'bg-black/30 text-red-300/40 border-red-900/50'
                      }`}
                    >
                      {p}
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* BÊN XANH (AO) */}
            <div className="bg-gradient-to-bl from-blue-700 via-blue-800 to-blue-950 p-4 sm:p-6 flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between gap-2">
                  <span className="px-2.5 py-1 rounded bg-black/40 text-blue-200 font-bold text-xs uppercase tracking-widest">
                    AO (XANH)
                  </span>
                  {match.aoSenshu && (
                    <span className="px-3 py-1 rounded-full bg-yellow-400 text-black font-black text-xs sm:text-sm shadow-lg">
                      ★ SENSHU
                    </span>
                  )}
                </div>
                <h3 className={`${isPreview ? 'text-lg sm:text-2xl' : 'text-2xl sm:text-5xl'} font-black uppercase mt-2 tracking-tight truncate`}>
                  {match.aoName}
                </h3>
                <p className="text-blue-200 text-xs sm:text-base font-semibold opacity-90 truncate">{match.aoTeam}</p>
              </div>

              <div className="my-auto text-center py-4">
                <span className={`${isPreview ? 'text-6xl sm:text-8xl' : 'text-8xl sm:text-[170px]'} font-black leading-none drop-shadow-2xl`}>
                  {match.aoScore}
                </span>
              </div>

              <div className="space-y-1.5">
                <p className="text-[10px] sm:text-xs uppercase tracking-wider text-blue-200 font-bold">Lỗi (Penalties)</p>
                <div className="grid grid-cols-5 gap-1.5 sm:gap-2">
                  {PENALTIES.map((p) => (
                    <div
                      key={p}
                      className={`py-1.5 sm:py-2.5 rounded-xl text-center font-black text-xs sm:text-base border-2 ${
                        aoP[p]
                          ? p === 'H'
                            ? 'bg-red-500 text-white border-white'
                            : 'bg-yellow-400 text-black border-yellow-200 shadow-lg'
                          : 'bg-black/30 text-blue-300/40 border-blue-900/50'
                      }`}
                    >
                      {p}
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* ĐỒNG HỒ TRUNG TÂM */}
            <div className="absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 z-10 pointer-events-none">
              <div
                className={`px-4 py-2 sm:px-7 sm:py-4 rounded-2xl sm:rounded-3xl border-4 shadow-2xl text-center backdrop-blur-md ${
                  isAtoshiBaraku
                    ? 'bg-red-950/95 border-yellow-400 text-yellow-400 animate-pulse'
                    : match.timeLeft === 0
                    ? 'bg-rose-950 border-rose-500 text-rose-400'
                    : 'bg-slate-950/90 border-slate-700 text-emerald-400'
                }`}
              >
                <p className="text-[9px] sm:text-[11px] uppercase tracking-widest text-slate-400 font-bold">
                  {isAtoshiBaraku ? 'ATOSHI BARAKU (15s)' : 'THỜI GIAN'}
                </p>
                <span className={`${isPreview ? 'text-2xl sm:text-4xl' : 'text-4xl sm:text-7xl'} font-mono font-black tracking-wider`}>
                  {formatClock(match.timeLeft)}
                </span>
              </div>
            </div>
          </div>
        ) : (
          <div className="flex-1 grid grid-cols-2">
            <div className="bg-gradient-to-br from-red-800 to-red-950 p-6 flex flex-col justify-between border-r border-black/40">
              <div>
                <span className="px-3 py-1 rounded bg-black/40 text-red-200 font-bold text-xs uppercase">AKA (ĐỎ)</span>
                <h3 className={`${isPreview ? 'text-xl' : 'text-4xl'} font-black uppercase mt-2`}>{match.akaName}</h3>
                <p className="text-amber-300 font-bold text-sm sm:text-lg mt-1">Kata: {match.akaKataName}</p>
              </div>
              <div className="my-auto text-center py-6">
                <p className="text-xs uppercase text-red-200 font-bold mb-1">
                  {match.kataSystem === 'flags' ? 'SỐ CỜ ĐỎ' : 'TỔNG ĐIỂM KATA'}
                </p>
                <span className={`${isPreview ? 'text-6xl' : 'text-9xl'} font-black`}>
                  {match.kataSystem === 'flags' ? match.akaFlags : calcKataTotal(match.akaJudges)}
                </span>
              </div>
            </div>

            <div className="bg-gradient-to-bl from-blue-800 to-blue-950 p-6 flex flex-col justify-between">
              <div>
                <span className="px-3 py-1 rounded bg-black/40 text-blue-200 font-bold text-xs uppercase">AO (XANH)</span>
                <h3 className={`${isPreview ? 'text-xl' : 'text-4xl'} font-black uppercase mt-2`}>{match.aoName}</h3>
                <p className="text-amber-300 font-bold text-sm sm:text-lg mt-1">Kata: {match.aoKataName}</p>
              </div>
              <div className="my-auto text-center py-6">
                <p className="text-xs uppercase text-blue-200 font-bold mb-1">
                  {match.kataSystem === 'flags' ? 'SỐ CỜ XANH' : 'TỔNG ĐIỂM KATA'}
                </p>
                <span className={`${isPreview ? 'text-6xl' : 'text-9xl'} font-black`}>
                  {match.kataSystem === 'flags' ? match.aoFlags : calcKataTotal(match.aoJudges)}
                </span>
              </div>
            </div>
          </div>
        )}

        {match.winner && (
          <div className="bg-amber-400 text-black py-3 px-6 flex items-center justify-center gap-2 font-black text-sm sm:text-2xl uppercase tracking-wider">
            <Trophy className="w-6 h-6 shrink-0" />
            CHIẾN THẮNG: {match.winner === 'AKA' ? `AKA (ĐỎ) — ${match.akaName}` : `AO (XANH) — ${match.aoName}`}
          </div>
        )}
      </div>
    );
  };

  if (isDisplayOnly) {
    return renderPublicScoreboard(false);
  }

  // ============================================================================
  // GIAO DIỆN BẢNG ĐIỀU KHIỂN ADMIN + QUẢN LÝ SIMA GIẢI ĐẤU
  // ============================================================================
  return (
    <div className="space-y-5">
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleImportSimaExcel}
        accept=".xlsx, .xls, .csv"
        className="hidden"
      />

      {/* 1. Thanh công cụ chế độ & Mở màn chiếu TV */}
      <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={() => updateMatch({ mode: 'kumite', winner: null })}
            className={`px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 transition ${
              match.mode === 'kumite' ? 'bg-red-600 text-white shadow-sm' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
            }`}
          >
            <Swords className="w-4 h-4" /> Thi Đấu KUMITE (Đối Kháng)
          </button>
          <button
            onClick={() => updateMatch({ mode: 'kata', winner: null })}
            className={`px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 transition ${
              match.mode === 'kata' ? 'bg-indigo-600 text-white shadow-sm' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
            }`}
          >
            <Award className="w-4 h-4" /> Thi Đấu KATA (Biểu Diễn Quyền)
          </button>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={playStartGong}
            className="px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold flex items-center gap-1.5"
          >
            <Volume2 className="w-4 h-4 text-amber-600" /> Còi Trọng Tài
          </button>
          <button
            onClick={handleResetMatch}
            className="px-3.5 py-2.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-xl text-xs font-bold flex items-center gap-1.5"
          >
            <RotateCcw className="w-4 h-4" /> Reset Trận
          </button>
          <button
            onClick={handleOpenDisplayWindow}
            className="px-4 py-2.5 bg-slate-900 hover:bg-black text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm"
          >
            <Monitor className="w-4 h-4 text-emerald-400" /> Mở Màn Hình Chiếu TV
          </button>
        </div>
      </div>

      {/* 2. KHỐI QUẢN LÝ DANH SÁCH HẠNG CÂN / LỊCH THI ĐẤU (SIMA EXCEL) */}
      <div className="bg-white rounded-2xl border border-indigo-200 overflow-hidden shadow-xs">
        <div className="p-4 bg-indigo-50/50 border-b border-indigo-100 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h3 className="font-bold text-sm text-indigo-950 flex items-center gap-2">
              <ListOrdered className="w-4 h-4 text-indigo-600" /> Danh Sách Hạng Cân & Bốc Thăm Thi Đấu (Sima)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Import danh sách hạng cân / cặp đấu từ Excel để tự động nạp tên VĐV, đơn vị và thời gian lên bảng điểm
            </p>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={handleDownloadSimaTemplate}
              className="px-3 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold flex items-center gap-1.5"
            >
              <Download className="w-3.5 h-3.5 text-indigo-600" /> Tải File Mẫu Sima
            </button>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs"
            >
              <Upload className="w-3.5 h-3.5" /> Nhập Sima / Hạng Cân (Excel)
            </button>
            {scheduleList.length > 0 && (
              <>
                <button
                  onClick={handleExportTournamentResults}
                  className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" /> Xuất Kết Quả Giải
                </button>
                <button
                  onClick={() => {
                    if (window.confirm('Xóa toàn bộ danh sách lịch thi đấu hiện tại?')) saveScheduleList([]);
                  }}
                  className="p-2 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-xl text-xs"
                  title="Xóa danh sách Sima"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </>
            )}
          </div>
        </div>

        {scheduleList.length > 0 && (
          <div className="p-4 space-y-3">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-xs font-bold text-slate-600">Lọc theo Hạng cân:</span>
                <select
                  value={categoryFilter}
                  onChange={(e) => setCategoryFilter(e.target.value)}
                  className="px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-indigo-700"
                >
                  <option value="ALL">Tất cả hạng cân ({scheduleList.length} trận)</option>
                  {uniqueCategories.map((cat) => (
                    <option key={cat} value={cat}>{cat}</option>
                  ))}
                </select>
              </div>

              <button
                onClick={handleNextMatch}
                className="px-4 py-1.5 bg-amber-500 hover:bg-amber-600 text-white rounded-xl text-xs font-black flex items-center gap-1 shadow-xs"
              >
                Gọi Trận Tiếp Theo <ChevronRight className="w-4 h-4" />
              </button>
            </div>

            <div className="max-h-48 overflow-y-auto border border-slate-200 rounded-xl divide-y divide-slate-100">
              {filteredSchedule.map((item) => {
                const isCurrent = match.matchId === item.id;
                return (
                  <div
                    key={item.id}
                    className={`p-2.5 px-3.5 flex items-center justify-between gap-3 text-xs transition ${
                      isCurrent ? 'bg-indigo-50/80 font-bold' : 'hover:bg-slate-50'
                    }`}
                  >
                    <div className="flex items-center gap-2.5 min-w-0 flex-wrap">
                      <span className="font-mono font-black px-2 py-0.5 rounded bg-slate-900 text-white text-[11px]">
                        {item.id}
                      </span>
                      <span className="px-2 py-0.5 rounded bg-indigo-100 text-indigo-800 font-bold text-[11px]">
                        {item.categoryName}
                      </span>
                      <span className="text-red-600 font-bold">{item.akaName}</span>
                      <span className="text-slate-400 font-bold">vs</span>
                      <span className="text-blue-600 font-bold">{item.aoName}</span>
                      {item.winnerName && (
                        <span className="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 font-bold text-[10px] inline-flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3" /> Thắng: {item.winnerName} ({item.finalScore})
                        </span>
                      )}
                    </div>

                    <button
                      onClick={() => handleLoadMatchFromSchedule(item)}
                      className={`px-3 py-1.5 rounded-lg font-bold shrink-0 ${
                        isCurrent
                          ? 'bg-indigo-600 text-white'
                          : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                      }`}
                    >
                      {isCurrent ? 'Đang thi đấu' : 'Gọi trận này'}
                    </button>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* 3. Chỉnh sửa trực tiếp Hạng cân & VĐV trận hiện tại */}
      <div className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-5 shadow-xs space-y-4">
        <div>
          <label className="block text-xs font-bold text-slate-500 uppercase mb-1.5">Tên nội dung / Hạng cân đang thi đấu</label>
          <input
            type="text"
            value={match.categoryName}
            onChange={(e) => updateMatch({ categoryName: e.target.value })}
            className="w-full px-3.5 py-2.5 border border-slate-200 rounded-xl text-sm font-bold text-slate-800 focus:outline-none focus:border-indigo-500"
          />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* VĐV ĐỎ */}
          <div className="p-4 rounded-2xl bg-red-50/50 border border-red-200 space-y-2.5">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <span className="text-xs font-black text-red-700 uppercase">VĐV BÊN ĐỎ (AKA)</span>
              <select
                onChange={(e) => {
                  const st = students.find((s) => s.id === e.target.value);
                  if (st) updateMatch({ akaName: st.name.toUpperCase(), akaTeam: st.className });
                }}
                className="text-xs bg-white border border-red-200 rounded-lg px-2.5 py-1.5 font-semibold text-slate-700"
              >
                <option value="">-- Chọn nhanh học sinh trung tâm --</option>
                {students.map((s) => (
                  <option key={s.id} value={s.id}>{s.name} ({s.className})</option>
                ))}
              </select>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <input
                type="text"
                value={match.akaName}
                onChange={(e) => updateMatch({ akaName: e.target.value })}
                placeholder="Họ tên VĐV Đỏ"
                className="px-3 py-2 bg-white border border-red-200 rounded-xl text-xs font-bold"
              />
              <input
                type="text"
                value={match.akaTeam}
                onChange={(e) => updateMatch({ akaTeam: e.target.value })}
                placeholder="Đơn vị / CLB"
                className="px-3 py-2 bg-white border border-red-200 rounded-xl text-xs font-semibold"
              />
            </div>
          </div>

          {/* VĐV XANH */}
          <div className="p-4 rounded-2xl bg-blue-50/50 border border-blue-200 space-y-2.5">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <span className="text-xs font-black text-blue-700 uppercase">VĐV BÊN XANH (AO)</span>
              <select
                onChange={(e) => {
                  const st = students.find((s) => s.id === e.target.value);
                  if (st) updateMatch({ aoName: st.name.toUpperCase(), aoTeam: st.className });
                }}
                className="text-xs bg-white border border-blue-200 rounded-lg px-2.5 py-1.5 font-semibold text-slate-700"
              >
                <option value="">-- Chọn nhanh học sinh trung tâm --</option>
                {students.map((s) => (
                  <option key={s.id} value={s.id}>{s.name} ({s.className})</option>
                ))}
              </select>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <input
                type="text"
                value={match.aoName}
                onChange={(e) => updateMatch({ aoName: e.target.value })}
                placeholder="Họ tên VĐV Xanh"
                className="px-3 py-2 bg-white border border-blue-200 rounded-xl text-xs font-bold"
              />
              <input
                type="text"
                value={match.aoTeam}
                onChange={(e) => updateMatch({ aoTeam: e.target.value })}
                placeholder="Đơn vị / CLB"
                className="px-3 py-2 bg-white border border-blue-200 rounded-xl text-xs font-semibold"
              />
            </div>
          </div>
        </div>
      </div>

      {/* 4. BẢNG ĐIỀU KHIỂN KUMITE */}
      {match.mode === 'kumite' && (
        <div className="space-y-4">
          <div className="bg-slate-900 text-white rounded-2xl p-4 sm:p-5 flex flex-col md:flex-row items-center justify-between gap-4 shadow-md">
            <div className="flex items-center gap-4">
              <div className="text-center md:text-left">
                <span className="text-[10px] uppercase tracking-widest text-slate-400 font-bold block">Đồng Hồ Thi Đấu</span>
                <span className="text-4xl sm:text-5xl font-mono font-black text-emerald-400">
                  {formatClock(match.timeLeft)}
                </span>
              </div>
              <div className="flex items-center gap-1.5 flex-wrap">
                {[60, 90, 120, 180].map((sec) => (
                  <button
                    key={sec}
                    onClick={() => updateMatch({ duration: sec, timeLeft: sec, isRunning: false })}
                    className={`px-3 py-1.5 rounded-xl text-xs font-bold transition ${
                      match.duration === sec ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    {sec / 60} phút
                  </button>
                ))}
              </div>
            </div>

            <div className="flex items-center gap-2 w-full md:w-auto justify-center">
              <button
                onClick={() => {
                  if (!match.isRunning && match.timeLeft > 0) {
                    playStartGong(); // Phát tiếng kẻng lớn bắt đầu trận đấu
                  }
                  updateMatch({ isRunning: !match.isRunning });
                }}
                className={`px-6 py-3 rounded-xl font-black text-xs flex items-center justify-center gap-2 shadow-sm ${
                  match.isRunning ? 'bg-amber-500 text-black' : 'bg-emerald-600 hover:bg-emerald-700 text-white'
                }`}
              >
                {match.isRunning ? <><Pause className="w-4 h-4" /> TẠM DỪNG (STOP)</> : <><Play className="w-4 h-4" /> BẮT ĐẦU (HAJIME)</>}
              </button>
              <button
                onClick={() => updateMatch({ timeLeft: match.duration, isRunning: false })}
                className="p-3 bg-slate-800 hover:bg-slate-700 rounded-xl text-xs font-bold"
                title="Đặt lại đồng hồ"
              >
                <RotateCcw className="w-4 h-4" />
              </button>
            </div>

            <div className="flex items-center gap-2 w-full md:w-auto justify-end">
              <button
                onClick={() => updateMatch({ winner: match.winner === 'AKA' ? null : 'AKA', isRunning: false })}
                className={`px-3.5 py-2.5 rounded-xl text-xs font-bold transition ${
                  match.winner === 'AKA' ? 'bg-yellow-400 text-black' : 'bg-red-600 hover:bg-red-700 text-white'
                }`}
              >
                🏆 AKA Thắng
              </button>
              <button
                onClick={() => updateMatch({ winner: match.winner === 'AO' ? null : 'AO', isRunning: false })}
                className={`px-3.5 py-2.5 rounded-xl text-xs font-bold transition ${
                  match.winner === 'AO' ? 'bg-yellow-400 text-black' : 'bg-blue-600 hover:bg-blue-700 text-white'
                }`}
              >
                🏆 AO Thắng
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* ĐIỀU KHIỂN AKA */}
            <div className="bg-white rounded-2xl border-2 border-red-200 p-5 space-y-4 shadow-xs">
              <div className="flex items-center justify-between gap-2 border-b border-red-100 pb-3">
                <div>
                  <span className="font-black text-base text-red-600 uppercase block">ĐIỀU KHIỂN AKA (ĐỎ)</span>
                  <span className="text-xs font-bold text-slate-500">{match.akaName} — Điểm: <strong className="text-red-600 text-base">{match.akaScore}</strong></span>
                </div>
                <button
                  onClick={() => updateMatch({ akaSenshu: !match.akaSenshu })}
                  className={`px-3.5 py-1.5 rounded-xl text-xs font-bold border transition shrink-0 ${
                    match.akaSenshu ? 'bg-yellow-400 text-black border-yellow-500 shadow-xs' : 'bg-slate-100 text-slate-600 border-slate-200'
                  }`}
                >
                  ★ Senshu (Điểm đầu)
                </button>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                <button onClick={() => updateMatch({ akaScore: match.akaScore + 1 })} className="py-3 px-2 bg-red-600 hover:bg-red-700 active:scale-95 text-white rounded-xl font-black text-xs transition shadow-xs">+1 Yuko</button>
                <button onClick={() => updateMatch({ akaScore: match.akaScore + 2 })} className="py-3 px-2 bg-red-600 hover:bg-red-700 active:scale-95 text-white rounded-xl font-black text-xs transition shadow-xs">+2 Waza-ari</button>
                <button onClick={() => updateMatch({ akaScore: match.akaScore + 3 })} className="py-3 px-2 bg-red-600 hover:bg-red-700 active:scale-95 text-white rounded-xl font-black text-xs transition shadow-xs">+3 Ippon</button>
                <button onClick={() => updateMatch({ akaScore: Math.max(0, match.akaScore - 1) })} className="py-3 px-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold text-xs transition">-1 Điểm</button>
              </div>

              <div>
                <p className="text-xs font-bold text-slate-500 uppercase mb-2">Phạt lỗi AKA (C1 → H):</p>
                <div className="grid grid-cols-5 gap-2">
                  {PENALTIES.map((p) => (
                    <button
                      key={p}
                      onClick={() => updateMatch({ akaPenalties: { ...match.akaPenalties, [p]: !match.akaPenalties[p] } })}
                      className={`py-2.5 rounded-xl font-black text-xs border transition ${
                        match.akaPenalties?.[p] ? 'bg-amber-400 text-black border-amber-500 shadow-xs' : 'bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-200'
                      }`}
                    >
                      {p}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* ĐIỀU KHIỂN AO */}
            <div className="bg-white rounded-2xl border-2 border-blue-200 p-5 space-y-4 shadow-xs">
              <div className="flex items-center justify-between gap-2 border-b border-blue-100 pb-3">
                <div>
                  <span className="font-black text-base text-blue-600 uppercase block">ĐIỀU KHIỂN AO (XANH)</span>
                  <span className="text-xs font-bold text-slate-500">{match.aoName} — Điểm: <strong className="text-blue-600 text-base">{match.aoScore}</strong></span>
                </div>
                <button
                  onClick={() => updateMatch({ aoSenshu: !match.aoSenshu })}
                  className={`px-3.5 py-1.5 rounded-xl text-xs font-bold border transition shrink-0 ${
                    match.aoSenshu ? 'bg-yellow-400 text-black border-yellow-500 shadow-xs' : 'bg-slate-100 text-slate-600 border-slate-200'
                  }`}
                >
                  ★ Senshu (Điểm đầu)
                </button>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                <button onClick={() => updateMatch({ aoScore: match.aoScore + 1 })} className="py-3 px-2 bg-blue-600 hover:bg-blue-700 active:scale-95 text-white rounded-xl font-black text-xs transition shadow-xs">+1 Yuko</button>
                <button onClick={() => updateMatch({ aoScore: match.aoScore + 2 })} className="py-3 px-2 bg-blue-600 hover:bg-blue-700 active:scale-95 text-white rounded-xl font-black text-xs transition shadow-xs">+2 Waza-ari</button>
                <button onClick={() => updateMatch({ aoScore: match.aoScore + 3 })} className="py-3 px-2 bg-blue-600 hover:bg-blue-700 active:scale-95 text-white rounded-xl font-black text-xs transition shadow-xs">+3 Ippon</button>
                <button onClick={() => updateMatch({ aoScore: Math.max(0, match.aoScore - 1) })} className="py-3 px-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold text-xs transition">-1 Điểm</button>
              </div>

              <div>
                <p className="text-xs font-bold text-slate-500 uppercase mb-2">Phạt lỗi AO (C1 → H):</p>
                <div className="grid grid-cols-5 gap-2">
                  {PENALTIES.map((p) => (
                    <button
                      key={p}
                      onClick={() => updateMatch({ aoPenalties: { ...match.aoPenalties, [p]: !match.aoPenalties[p] } })}
                      className={`py-2.5 rounded-xl font-black text-xs border transition ${
                        match.aoPenalties?.[p] ? 'bg-amber-400 text-black border-amber-500 shadow-xs' : 'bg-slate-100 text-slate-600 border-slate-200 hover:bg-slate-200'
                      }`}
                    >
                      {p}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 5. BẢNG ĐIỀU KHIỂN KATA */}
      {match.mode === 'kata' && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 space-y-5 shadow-xs">
          <div className="flex items-center justify-between flex-wrap gap-2 border-b pb-3">
            <h3 className="font-bold text-sm text-slate-800">Chế độ chấm điểm Kata</h3>
            <div className="flex items-center gap-2">
              <button
                onClick={() => updateMatch({ kataSystem: 'flags' })}
                className={`px-3.5 py-2 rounded-xl text-xs font-bold ${
                  match.kataSystem === 'flags' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-600'
                }`}
              >
                Chấm Cờ (5 Trọng Tài)
              </button>
              <button
                onClick={() => updateMatch({ kataSystem: 'points' })}
                className={`px-3.5 py-2 rounded-xl text-xs font-bold ${
                  match.kataSystem === 'points' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-600'
                }`}
              >
                Chấm Điểm Số (Bỏ Cao/Thấp Nhất)
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-3 p-4 rounded-xl bg-red-50/50 border border-red-200">
              <label className="block text-xs font-bold text-red-700">Tên bài Kata AKA (Đỏ)</label>
              <input
                type="text"
                value={match.akaKataName}
                onChange={(e) => updateMatch({ akaKataName: e.target.value })}
                className="w-full px-3 py-2 bg-white border border-red-200 rounded-xl text-sm font-bold"
              />
              {match.kataSystem === 'flags' ? (
                <div>
                  <span className="block text-xs font-bold text-slate-600 mb-1.5">Số Cờ Đỏ (0 - 5):</span>
                  <div className="flex gap-2">
                    {[0, 1, 2, 3, 4, 5].map((num) => (
                      <button
                        key={num}
                        onClick={() => updateMatch({ akaFlags: num, aoFlags: 5 - num, winner: num >= 3 ? 'AKA' : 'AO' })}
                        className={`flex-1 py-2.5 rounded-xl font-black text-sm ${
                          match.akaFlags === num ? 'bg-red-600 text-white' : 'bg-white border text-slate-700'
                        }`}
                      >
                        {num}
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <div>
                  <span className="block text-xs font-bold text-slate-600 mb-1.5">Điểm 5 Trọng Tài (Tự bỏ Max/Min):</span>
                  <div className="grid grid-cols-5 gap-2">
                    {match.akaJudges.map((sc, i) => (
                      <input
                        key={i}
                        type="number"
                        step="0.1"
                        value={sc}
                        onChange={(e) => {
                          const next = [...match.akaJudges];
                          next[i] = e.target.value;
                          updateMatch({ akaJudges: next });
                        }}
                        className="px-2 py-2 bg-white border border-red-200 rounded-lg text-center font-bold text-sm"
                      />
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div className="space-y-3 p-4 rounded-xl bg-blue-50/50 border border-blue-200">
              <label className="block text-xs font-bold text-blue-700">Tên bài Kata AO (Xanh)</label>
              <input
                type="text"
                value={match.aoKataName}
                onChange={(e) => updateMatch({ aoKataName: e.target.value })}
                className="w-full px-3 py-2 bg-white border border-blue-200 rounded-xl text-sm font-bold"
              />
              {match.kataSystem === 'flags' ? (
                <div>
                  <span className="block text-xs font-bold text-slate-600 mb-1.5">Số Cờ Xanh (0 - 5):</span>
                  <div className="flex gap-2">
                    {[0, 1, 2, 3, 4, 5].map((num) => (
                      <button
                        key={num}
                        onClick={() => updateMatch({ aoFlags: num, akaFlags: 5 - num, winner: num >= 3 ? 'AO' : 'AKA' })}
                        className={`flex-1 py-2.5 rounded-xl font-black text-sm ${
                          match.aoFlags === num ? 'bg-blue-600 text-white' : 'bg-white border text-slate-700'
                        }`}
                      >
                        {num}
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <div>
                  <span className="block text-xs font-bold text-slate-600 mb-1.5">Điểm 5 Trọng Tài (Tự bỏ Max/Min):</span>
                  <div className="grid grid-cols-5 gap-2">
                    {match.aoJudges.map((sc, i) => (
                      <input
                        key={i}
                        type="number"
                        step="0.1"
                        value={sc}
                        onChange={(e) => {
                          const next = [...match.aoJudges];
                          next[i] = e.target.value;
                          updateMatch({ aoJudges: next });
                        }}
                        className="px-2 py-2 bg-white border border-blue-200 rounded-lg text-center font-bold text-sm"
                      />
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* 6. MÀN HÌNH XEM TRƯỚC */}
      <div className="space-y-2">
        <p className="text-xs font-bold text-slate-500 uppercase tracking-wider">
          Màn hình xem trước (Live Preview — Giống 100% trên màn chiếu TV)
        </p>
        {renderPublicScoreboard(true)}
      </div>
    </div>
  );
}