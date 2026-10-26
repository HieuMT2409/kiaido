import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Users,
  Bell,
  CheckCircle2,
  Search,
  Plus,
  Calendar,
  Edit2,
  Trash2,
  ChevronLeft,
  ChevronRight,
  Download,
  Cloud,
  CloudOff,
  Settings,
  X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  getFirestore,
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  writeBatch
} from 'firebase/firestore';

// ==========================================
// 1. KHỞI TẠO FIREBASE TỪ FILE .ENV
// ==========================================
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

let db = null;
try {
  if (firebaseConfig.apiKey && firebaseConfig.projectId) {
    const app = initializeApp(firebaseConfig);
    db = getFirestore(app);
  }
} catch (error) {
  console.error('Lỗi khởi tạo Firebase:', error);
}

const INITIAL_COURSES = [];
const INITIAL_STUDENTS = [];
const INITIAL_PAYMENTS = {};

const formatVND = (amount) => {
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(Number(amount) || 0);
};

const formatDateTime = (isoStr) => {
  if (!isoStr) return '—';
  const date = new Date(isoStr);
  if (isNaN(date.getTime())) return isoStr;
  return date.toLocaleString('vi-VN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
};

export default function App() {
  const [courses, setCourses] = useState(INITIAL_COURSES);
  const [students, setStudents] = useState(INITIAL_STUDENTS);
  const [payments, setPayments] = useState(INITIAL_PAYMENTS);

  const [firebaseStatus, setFirebaseStatus] = useState(db ? 'connecting' : 'offline');
  const [firebaseErrorMsg, setFirebaseErrorMsg] = useState('');

  const [selectedMonth, setSelectedMonth] = useState('2026-10');
  const [activeTab, setActiveTab] = useState('tuition');
  const [searchQuery, setSearchQuery] = useState('');
  const [classFilter, setClassFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');

  const [showNotifPanel, setShowNotifPanel] = useState(false);
  
  // Modal thu tiền chi tiết (chọn phương thức, ngày giờ, ghi chú)
  const [paymentModal, setPaymentModal] = useState({ 
    open: false, 
    student: null, 
    paidDate: '', 
    method: 'Chuyển khoản', 
    note: '' 
  });

  const [studentModal, setStudentModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '' } });
  const [courseModal, setCourseModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' }, updateStudentsFee: true });
  const [toast, setToast] = useState(null);

  const showToast = (message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3500);
  };

  useEffect(() => {
    if (!db) {
      setFirebaseStatus('offline');
      return;
    }

    const unsubCourses = onSnapshot(
      collection(db, 'courses'),
      (snapshot) => {
        setFirebaseStatus('connected');
        setFirebaseErrorMsg('');
        const list = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
        setCourses(list);
      },
      (err) => {
        setFirebaseStatus('error');
        setFirebaseErrorMsg(err.message);
      }
    );

    const unsubStudents = onSnapshot(
      collection(db, 'students'),
      (snapshot) => {
        setFirebaseStatus('connected');
        const list = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
        list.sort((a, b) => a.id.localeCompare(b.id));
        setStudents(list);
      },
      (err) => console.error(err)
    );

    const unsubPayments = onSnapshot(
      collection(db, 'payments'),
      (snapshot) => {
        const payMap = {};
        snapshot.docs.forEach((d) => {
          payMap[d.id] = d.data();
        });
        setPayments(payMap);
      },
      (err) => console.error(err)
    );

    return () => {
      unsubCourses();
      unsubStudents();
      unsubPayments();
    };
  }, []);

  const currentMonthPayments = payments[selectedMonth] || {};

  const enrichedStudents = useMemo(() => {
    return students.map((student) => {
      const paymentInfo = currentMonthPayments[student.id] || { paid: false, paidDate: null, method: 'Chuyển khoản', note: '' };
      return {
        ...student,
        paid: Boolean(paymentInfo.paid),
        paidDate: paymentInfo.paidDate || null,
        method: paymentInfo.method || 'Chuyển khoản',
        note: paymentInfo.note || ''
      };
    });
  }, [students, currentMonthPayments]);

  const unpaidStudents = useMemo(() => enrichedStudents.filter((s) => !s.paid), [enrichedStudents]);
  const paidStudents = useMemo(() => enrichedStudents.filter((s) => s.paid), [enrichedStudents]);

  const stats = useMemo(() => {
    const expectedRevenue = enrichedStudents.reduce((sum, s) => sum + Number(s.tuitionFee || 0), 0);
    const collectedRevenue = paidStudents.reduce((sum, s) => sum + Number(s.tuitionFee || 0), 0);
    const pendingRevenue = expectedRevenue - collectedRevenue;
    const completionRate = enrichedStudents.length > 0 ? Math.round((paidStudents.length / enrichedStudents.length) * 100) : 0;
    return { expectedRevenue, collectedRevenue, pendingRevenue, completionRate };
  }, [enrichedStudents, paidStudents]);

  const filteredStudents = useMemo(() => {
    return enrichedStudents.filter((s) => {
      const matchSearch = s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.id.toLowerCase().includes(searchQuery.toLowerCase()) || (s.phone && s.phone.includes(searchQuery));
      const matchClass = classFilter === 'ALL' || s.className === classFilter;
      const matchStatus = statusFilter === 'ALL' || (statusFilter === 'PAID' && s.paid) || (statusFilter === 'UNPAID' && !s.paid);
      return matchSearch && matchClass && matchStatus;
    });
  }, [enrichedStudents, searchQuery, classFilter, statusFilter]);

  const changeMonth = (offset) => {
    const [year, month] = selectedMonth.split('-').map(Number);
    const date = new Date(year, month - 1 + offset, 1);
    setSelectedMonth(`${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`);
  };

  const saveMonthPaymentData = async (monthKey, updatedMonthObj) => {
    setPayments((prev) => ({ ...prev, [monthKey]: updatedMonthObj }));
    if (db) {
      try {
        await setDoc(doc(db, 'payments', monthKey), updatedMonthObj);
      } catch (err) {
        showToast('Lỗi lưu Firebase: ' + err.message, 'error');
      }
    }
  };

  // Mở modal thu tiền chi tiết để Admin chọn Phương thức & Ngày giờ đóng
  const openPaymentModal = (student) => {
    const nowIso = new Date().toISOString().slice(0, 16);
    setPaymentModal({
      open: true,
      student,
      paidDate: student.paidDate || nowIso,
      method: student.method || 'Chuyển khoản',
      note: student.note || ''
    });
  };

  // Lưu thông tin thu tiền từ Modal
  const handleSavePaymentModal = async (e) => {
    e.preventDefault();
    const { student, paidDate, method, note } = paymentModal;
    if (!student) return;

    const updatedMonth = {
      ...(payments[selectedMonth] || {}),
      [student.id]: {
        paid: true,
        paidDate: paidDate || new Date().toISOString().slice(0, 16),
        method,
        note,
        amount: Number(student.tuitionFee)
      }
    };
    await saveMonthPaymentData(selectedMonth, updatedMonth);
    setPaymentModal({ open: false, student: null, paidDate: '', method: 'Chuyển khoản', note: '' });
    showToast(`Đã ghi nhận ${student.name} đóng học phí (${method})!`);
  };

  const handleMarkUnpaid = async (student) => {
    const updatedMonth = {
      ...(payments[selectedMonth] || {}),
      [student.id]: { paid: false, paidDate: null, method: 'Chuyển khoản', note: '', amount: Number(student.tuitionFee) }
    };
    await saveMonthPaymentData(selectedMonth, updatedMonth);
    showToast(`Đã chuyển ${student.name} về trạng thái Chưa đóng.`, 'warning');
  };

  const openAddCourseModal = () => {
    setCourseModal({ open: true, mode: 'add', data: { id: `COURSE_${Date.now()}`, name: '', fee: 1500000, description: '' }, updateStudentsFee: true });
  };

  const openEditCourseModal = (course) => {
    setCourseModal({ open: true, mode: 'edit', oldName: course.name, data: { ...course }, updateStudentsFee: true });
  };

  const handleSaveCourse = async (e) => {
    e.preventDefault();
    const { mode, data, oldName, updateStudentsFee } = courseModal;
    const cleanCourse = { id: data.id, name: data.name.trim(), fee: Number(data.fee) || 0, description: data.description || '' };
    if (!cleanCourse.name) return;

    if (mode === 'add') {
      setCourses((prev) => [...prev, cleanCourse]);
    } else {
      setCourses((prev) => prev.map((c) => (c.id === cleanCourse.id ? cleanCourse : c)));
    }

    if (db) {
      await setDoc(doc(db, 'courses', cleanCourse.id), cleanCourse).catch((err) => showToast(err.message, 'error'));
    }

    if (mode === 'edit' && updateStudentsFee) {
      const affectedStudents = students.filter((s) => s.className === oldName || s.className === cleanCourse.name);
      if (affectedStudents.length > 0) {
        const updatedStudents = students.map((s) => (s.className === oldName || s.className === cleanCourse.name ? { ...s, className: cleanCourse.name, tuitionFee: cleanCourse.fee } : s));
        setStudents(updatedStudents);
        if (db) {
          const batch = writeBatch(db);
          affectedStudents.forEach((s) => batch.set(doc(db, 'students', s.id), { ...s, className: cleanCourse.name, tuitionFee: cleanCourse.fee }));
          await batch.commit().catch(console.error);
        }
      }
    }

    setCourseModal({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' }, updateStudentsFee: true });
    showToast(mode === 'add' ? `Đã thêm khóa học` : `Đã cập nhật khóa học`);
  };

  const handleDeleteCourse = async (course) => {
    const enrolledStudents = students.filter((s) => s.className === course.name);
    if (enrolledStudents.length > 0) {
      window.alert(`Không thể xóa! Khóa học "${course.name}" hiện đang có ${enrolledStudents.length} học sinh theo học.`);
      return;
    }
    if (!window.confirm(`Bạn có chắc muốn xóa khóa học "${course.name}"?`)) return;
    setCourses((prev) => prev.filter((c) => c.id !== course.id));
    if (db) await deleteDoc(doc(db, 'courses', course.id)).catch(console.error);
    showToast(`Đã xóa khóa học`, 'warning');
  };

  const openAddStudentModal = () => {
    const defaultCourse = courses[0] || { name: 'Lớp Cơ Bản', fee: 1200000 };
    setStudentModal({
      open: true,
      mode: 'add',
      data: { id: `HS${String(students.length + 1).padStart(3, '0')}`, name: '', className: defaultCourse.name, phone: '', parentName: '', tuitionFee: defaultCourse.fee }
    });
  };

  const openEditStudentModal = (student) => {
    setStudentModal({ open: true, mode: 'edit', data: { ...student } });
  };

  const handleCourseSelectInStudentModal = (selectedCourseName) => {
    const foundCourse = courses.find((c) => c.name === selectedCourseName);
    setStudentModal((prev) => ({
      ...prev,
      data: { ...prev.data, className: selectedCourseName, tuitionFee: foundCourse ? foundCourse.fee : prev.data.tuitionFee }
    }));
  };

  // Bắt buộc nhập số điện thoại khi thêm/sửa học sinh
  const handleSaveStudent = async (e) => {
    e.preventDefault();
    const { mode, data } = studentModal;
    if (!data.name.trim() || !data.id.trim()) return;

    if (!data.phone || !data.phone.trim()) {
      alert('Vui lòng nhập số điện thoại phụ huynh!');
      return;
    }

    const studentObj = {
      id: data.id.trim().toUpperCase(),
      name: data.name.trim(),
      className: data.className.trim(),
      phone: data.phone.trim(),
      parentName: (data.parentName || '').trim(),
      tuitionFee: Number(data.tuitionFee) || 0,
      status: 'active'
    };

    if (mode === 'add') {
      setStudents((prev) => [...prev, studentObj]);
    } else {
      setStudents((prev) => prev.map((s) => (s.id === studentObj.id ? studentObj : s)));
    }

    if (db) {
      await setDoc(doc(db, 'students', studentObj.id), studentObj).catch((err) => showToast(err.message, 'error'));
    }
    setStudentModal({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '' } });
    showToast('Đã lưu hồ sơ học sinh!');
  };

  const handleDeleteStudent = async (studentId) => {
    if (!window.confirm('Bạn có chắc chắn muốn xóa học sinh này?')) return;
    setStudents((prev) => prev.filter((s) => s.id !== studentId));
    if (db) await deleteDoc(doc(db, 'students', studentId)).catch(console.error);
    showToast('Đã xóa học sinh.', 'warning');
  };

  // Xuất file Excel kèm đúng phương thức đóng tiền
  const handleExportExcel = () => {
    const studentRows = enrichedStudents.map((s, idx) => ({
      'STT': idx + 1,
      'Mã HS': s.id,
      'Họ và Tên': s.name,
      'Khóa học': s.className,
      'Phụ huynh': s.parentName,
      'Số điện thoại': s.phone,
      'Học phí': Number(s.tuitionFee),
      [`Trạng thái (${selectedMonth})`]: s.paid ? 'Đã đóng' : 'Chưa đóng',
      'Ngày đóng': s.paidDate ? formatDateTime(s.paidDate) : '',
      'Phương thức': s.paid ? s.method : ''
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(studentRows), `DS_${selectedMonth}`);
    XLSX.writeFile(wb, `QuanLyHocPhi_${selectedMonth}.xlsx`);
    showToast('Xuất Excel thành công!');
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col">
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg text-white text-sm font-medium bg-slate-900">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast.message}</span>
        </div>
      )}

      {/* HEADER */}
      <header className="bg-black border-b border-slate-800 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-3 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 sm:w-16 sm:h-16 rounded-xl bg-indigo-600 flex items-center justify-center text-white shadow-sm font-bold text-lg overflow-hidden shrink-0">
              <img src="/logo.jpg" alt="Logo" className="w-full h-full object-cover" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h1 className="text-base sm:text-lg font-bold text-red-500">KIAIDO Manager</h1>
                {firebaseStatus === 'connected' ? (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] sm:text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                    <Cloud className="w-3 h-3" /> Firebase
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] sm:text-[11px] font-medium bg-amber-50 text-amber-700 border border-amber-200">
                    <CloudOff className="w-3 h-3" /> Offline
                  </span>
                )}
              </div>
              <p className="text-[11px] sm:text-xs text-slate-400">Hệ thống quản lý học sinh & học phí</p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto justify-between sm:justify-end">
            <div className="flex items-center bg-slate-100 rounded-xl p-1 border border-slate-200">
              <button onClick={() => changeMonth(-1)} className="p-1.5 hover:bg-white rounded-lg transition"><ChevronLeft className="w-4 h-4" /></button>
              <div className="flex items-center gap-1.5 px-2 sm:px-3 text-xs sm:text-sm font-semibold text-slate-800">
                <Calendar className="w-4 h-4 text-indigo-600 shrink-0" />
                <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="bg-transparent border-none focus:outline-none cursor-pointer" />
              </div>
              <button onClick={() => changeMonth(1)} className="p-1.5 hover:bg-white rounded-lg transition"><ChevronRight className="w-4 h-4" /></button>
            </div>

            <div className="relative">
              <button onClick={() => setShowNotifPanel(!showNotifPanel)} className={`relative p-2.5 rounded-xl border transition ${unpaidStudents.length > 0 ? 'bg-rose-50 border-rose-200 text-rose-600' : 'bg-slate-100 border-slate-200 text-slate-600'}`}>
                <Bell className="w-5 h-5" />
                {unpaidStudents.length > 0 && <span className="absolute -top-1.5 -right-1.5 bg-rose-600 text-white text-[11px] font-bold w-5 h-5 rounded-full flex items-center justify-center">{unpaidStudents.length}</span>}
              </button>
              {showNotifPanel && (
                <div className="absolute right-0 mt-2 w-72 sm:w-80 bg-white rounded-2xl shadow-xl border border-slate-200 z-50 overflow-hidden">
                  <div className="p-3 bg-rose-50 border-b border-rose-100 font-semibold text-sm text-rose-800">Chưa đóng: {unpaidStudents.length} học sinh</div>
                  <div className="max-h-60 overflow-y-auto divide-y divide-slate-100">
                    {unpaidStudents.map((s) => (
                      <div key={s.id} className="p-3 flex items-center justify-between text-sm">
                        <span className="font-medium">{s.name}</span>
                        <button onClick={() => openPaymentModal(s)} className="px-2.5 py-1 bg-emerald-600 text-white text-xs rounded-lg">Thu ngay</button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto w-full px-3 sm:px-6 py-4 sm:py-6 space-y-4 sm:space-y-6 flex-1">
        {/* STATS */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
          <div className="bg-white p-3.5 sm:p-4 rounded-2xl border border-slate-200">
            <p className="text-xs text-slate-500">Dự kiến</p>
            <p className="text-base sm:text-xl font-bold text-slate-900 mt-1">{formatVND(stats.expectedRevenue)}</p>
          </div>
          <div className="bg-white p-3.5 sm:p-4 rounded-2xl border border-emerald-200">
            <p className="text-xs text-emerald-700">Đã thu</p>
            <p className="text-base sm:text-xl font-bold text-emerald-600 mt-1">{formatVND(stats.collectedRevenue)}</p>
          </div>
          <div className="bg-white p-3.5 sm:p-4 rounded-2xl border border-rose-200">
            <p className="text-xs text-rose-700">Còn thiếu</p>
            <p className="text-base sm:text-xl font-bold text-rose-600 mt-1">{formatVND(stats.pendingRevenue)}</p>
          </div>
          <div className="bg-white p-3.5 sm:p-4 rounded-2xl border border-slate-200">
            <p className="text-xs text-slate-500">Hoàn thành</p>
            <p className="text-base sm:text-xl font-bold text-indigo-600 mt-1">{stats.completionRate}%</p>
          </div>
        </div>

        {/* TABS & ACTIONS */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 pb-3">
          <div className="grid grid-cols-3 sm:flex items-center gap-1.5 sm:gap-2">
            <button onClick={() => setActiveTab('tuition')} className={`px-2.5 sm:px-4 py-2 rounded-xl text-xs sm:text-sm font-semibold text-center ${activeTab === 'tuition' ? 'bg-indigo-600 text-white' : 'bg-white text-slate-600 border border-slate-200'}`}>Thu Học Phí</button>
            <button onClick={() => setActiveTab('students')} className={`px-2.5 sm:px-4 py-2 rounded-xl text-xs sm:text-sm font-semibold text-center ${activeTab === 'students' ? 'bg-indigo-600 text-white' : 'bg-white text-slate-600 border border-slate-200'}`}>Học Sinh ({students.length})</button>
            <button onClick={() => setActiveTab('courses')} className={`px-2.5 sm:px-4 py-2 rounded-xl text-xs sm:text-sm font-semibold text-center ${activeTab === 'courses' ? 'bg-indigo-600 text-white' : 'bg-white text-slate-600 border border-slate-200'}`}>Khóa Học ({courses.length})</button>
          </div>

          <div className="grid grid-cols-3 sm:flex items-center gap-1.5 sm:gap-2">
            <button onClick={handleExportExcel} className="px-2.5 sm:px-3 py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold flex items-center justify-center gap-1"><Download className="w-3.5 h-3.5 text-emerald-600 shrink-0" /> Xuất Excel</button>
            <button onClick={openAddCourseModal} className="px-2.5 sm:px-3.5 py-2 bg-amber-500 hover:bg-amber-600 text-white rounded-xl text-xs font-semibold flex items-center justify-center gap-1"><Settings className="w-3.5 h-3.5 shrink-0" /> + Khóa Học</button>
            <button onClick={openAddStudentModal} className="px-2.5 sm:px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold flex items-center justify-center gap-1"><Plus className="w-3.5 h-3.5 shrink-0" /> + Học Sinh</button>
          </div>
        </div>

        {/* TAB: COURSES */}
        {activeTab === 'courses' && (
          <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden">
            <div className="p-4 border-b border-slate-200 bg-slate-50 font-bold text-sm sm:text-base">Danh mục Khóa học & Học phí</div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                    <th className="p-3.5">Tên Lớp</th>
                    <th className="p-3.5">Học Phí</th>
                    <th className="p-3.5 text-right">Thao tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-sm">
                  {courses.map((c) => (
                    <tr key={c.id}>
                      <td className="p-3.5 font-semibold">{c.name}</td>
                      <td className="p-3.5 font-bold text-indigo-600">{formatVND(c.fee)}</td>
                      <td className="p-3.5 text-right space-x-1.5 whitespace-nowrap">
                        <button onClick={() => openEditCourseModal(c)} className="px-2.5 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-semibold">Sửa</button>
                        <button onClick={() => handleDeleteCourse(c)} className="px-2.5 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Xóa</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* TAB: TUITION & STUDENTS */}
        {(activeTab === 'tuition' || activeTab === 'students') && (
          <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden">
            <div className="p-3.5 sm:p-4 border-b border-slate-200 flex items-center gap-3 bg-slate-50">
              <Search className="w-4 h-4 text-slate-400 shrink-0" />
              <input type="text" placeholder="Tìm kiếm học sinh, mã HS, số điện thoại..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full bg-transparent border-none text-sm focus:outline-none" />
            </div>

            {/* MOBILE CARD VIEW */}
            <div className="divide-y divide-slate-200 md:hidden">
              {filteredStudents.length === 0 ? (
                <div className="p-8 text-center text-sm text-slate-400">Chưa có học viên nào. Hãy bấm "+ Học Sinh" để thêm mới.</div>
              ) : (
                filteredStudents.map((s) => (
                  <div key={s.id} className="p-4 space-y-3 hover:bg-slate-50">
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="font-mono text-xs font-bold px-2 py-0.5 bg-slate-100 text-slate-700 rounded">{s.id}</span>
                          <span className="px-2 py-0.5 bg-indigo-50 text-indigo-700 rounded text-xs font-medium">{s.className}</span>
                        </div>
                        <p className="font-bold text-base text-slate-900 mt-1.5">{s.name}</p>
                        <p className="text-xs text-slate-500 mt-0.5">
                          PH: {s.parentName || '—'} •{' '}
                          {s.phone ? (
                            <a href={`https://zalo.me/${s.phone.replace(/[^0-9]/g, '')}`} target="_blank" rel="noopener noreferrer" className="text-indigo-600 font-semibold underline">
                              {s.phone} (Zalo)
                            </a>
                          ) : '—'}
                        </p>
                      </div>

                      <div className="text-right shrink-0">
                        <p className="font-bold text-sm text-slate-900">{formatVND(s.tuitionFee)}</p>
                        <div className="mt-1">
                          {s.paid ? (
                            <div className="space-y-0.5">
                              <span className="inline-block px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">Đã đóng</span>
                              <p className="text-[10px] text-slate-500 font-medium">{s.method} • {formatDateTime(s.paidDate)}</p>
                            </div>
                          ) : (
                            <span className="inline-block px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">Chưa đóng</span>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center justify-between gap-2 pt-2 border-t border-slate-100">
                      {!s.paid ? (
                        <button onClick={() => openPaymentModal(s)} className="flex-1 py-2 bg-emerald-600 active:bg-emerald-700 text-white rounded-xl text-xs font-semibold text-center shadow-xs">
                          Thu tiền
                        </button>
                      ) : (
                        <button onClick={() => handleMarkUnpaid(s)} className="flex-1 py-2 bg-rose-50 active:bg-rose-100 text-rose-600 rounded-xl text-xs font-semibold text-center">
                          Hủy thu
                        </button>
                      )}

                      <button onClick={() => openEditStudentModal(s)} className="px-3.5 py-2 bg-slate-100 active:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold inline-flex items-center gap-1.5">
                        <Edit2 className="w-3.5 h-3.5 text-indigo-600" /> Sửa
                      </button>

                      <button onClick={() => handleDeleteStudent(s.id)} className="px-3.5 py-2 bg-rose-50 active:bg-rose-100 text-rose-600 rounded-xl text-xs font-semibold inline-flex items-center gap-1.5">
                        <Trash2 className="w-3.5 h-3.5" /> Xóa
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* DESKTOP TABLE VIEW */}
            <div className="hidden md:block overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                    <th className="p-3.5">Mã HS</th>
                    <th className="p-3.5">Họ Tên & Zalo Phụ Huynh</th>
                    <th className="p-3.5">Khóa Học</th>
                    <th className="p-3.5">Học Phí</th>
                    <th className="p-3.5">Trạng Thái & Ngày Đóng</th>
                    <th className="p-3.5 text-right">Thao Tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-sm">
                  {filteredStudents.map((s) => (
                    <tr key={s.id} className="hover:bg-slate-50">
                      <td className="p-3.5 font-mono text-xs font-bold">{s.id}</td>
                      <td className="p-3.5">
                        <p className="font-semibold">{s.name}</p>
                        <p className="text-xs text-slate-500">
                          PH: {s.parentName || '—'} •{' '}
                          {s.phone ? (
                            <a href={`https://zalo.me/${s.phone.replace(/[^0-9]/g, '')}`} target="_blank" rel="noopener noreferrer" className="text-indigo-600 font-medium hover:underline">
                              {s.phone} (Zalo)
                            </a>
                          ) : '—'}
                        </p>
                      </td>
                      <td className="p-3.5"><span className="px-2 py-1 bg-indigo-50 text-indigo-700 rounded text-xs">{s.className}</span></td>
                      <td className="p-3.5 font-bold">{formatVND(s.tuitionFee)}</td>
                      <td className="p-3.5">
                        {s.paid ? (
                          <div className="space-y-0.5">
                            <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700">Đã đóng</span>
                            <p className="text-[11px] text-slate-500 font-medium">{s.method} • {formatDateTime(s.paidDate)}</p>
                          </div>
                        ) : (
                          <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-rose-50 text-rose-700">Chưa đóng</span>
                        )}
                      </td>
                      <td className="p-3.5 text-right space-x-1.5 whitespace-nowrap">
                        {!s.paid ? (
                          <button onClick={() => openPaymentModal(s)} className="px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-semibold shadow-xs">Thu tiền</button>
                        ) : (
                          <button onClick={() => handleMarkUnpaid(s)} className="px-3 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Hủy thu</button>
                        )}
                        <button onClick={() => openEditStudentModal(s)} className="px-2.5 py-1.5 bg-slate-100 text-slate-700 hover:bg-indigo-50 hover:text-indigo-600 rounded-lg text-xs font-semibold inline-flex items-center gap-1">
                          <Edit2 className="w-3.5 h-3.5" /> Sửa
                        </button>
                        <button onClick={() => handleDeleteStudent(s.id)} className="px-2.5 py-1.5 bg-rose-50 text-rose-600 hover:bg-rose-100 rounded-lg text-xs font-semibold inline-flex items-center gap-1">
                          <Trash2 className="w-3.5 h-3.5" /> Xóa
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {/* MODAL CHỌN PHƯƠNG THỨC THU TIỀN & NGÀY GIỜ */}
      {paymentModal.open && paymentModal.student && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold text-slate-900">Xác Nhận Thu Học Phí — {paymentModal.student.name}</h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Tháng {selectedMonth} • Số tiền: <strong className="text-emerald-600">{formatVND(paymentModal.student.tuitionFee)}</strong>
            </p>
            <form onSubmit={handleSavePaymentModal} className="space-y-4 mt-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Phương thức thanh toán <span className="text-rose-500">*</span></label>
                <select
                  value={paymentModal.method}
                  onChange={(e) => setPaymentModal({ ...paymentModal, method: e.target.value })}
                  className="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-sm bg-white font-semibold text-indigo-600 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                >
                  <option value="Chuyển khoản">Chuyển khoản ngân hàng</option>
                  <option value="Tiền mặt">Tiền mặt tại trung tâm</option>
                  <option value="Ví điện tử">Ví MoMo / ZaloPay</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Ngày & Giờ đóng tiền</label>
                <input
                  type="datetime-local"
                  required
                  value={paymentModal.paidDate}
                  onChange={(e) => setPaymentModal({ ...paymentModal, paidDate: e.target.value })}
                  className="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Ghi chú (Tùy chọn)</label>
                <input
                  type="text"
                  placeholder="VD: Chuyển khoản qua VCB..."
                  value={paymentModal.note}
                  onChange={(e) => setPaymentModal({ ...paymentModal, note: e.target.value })}
                  className="w-full px-3.5 py-2.5 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setPaymentModal({ ...paymentModal, open: false })}
                  className="px-4 py-2.5 rounded-xl text-xs font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200"
                >
                  Hủy
                </button>
                <button type="submit" className="px-5 py-2.5 rounded-xl text-xs font-semibold bg-emerald-600 text-white hover:bg-emerald-700 shadow-sm">
                  Xác Nhận Đã Thu
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL COURSE */}
      {courseModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold mb-4">{courseModal.mode === 'add' ? 'Thêm Khóa Học Mới' : 'Sửa Khóa Học'}</h3>
            <form onSubmit={handleSaveCourse} className="space-y-3">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Tên Khóa Học / Lớp *</label>
                <input type="text" required placeholder="Tên lớp" value={courseModal.data.name} onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Mức Học Phí Mặc Định (VNĐ/tháng) *</label>
                <input type="number" required placeholder="Học phí" value={courseModal.data.fee} onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, fee: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setCourseModal({ ...courseModal, open: false })} className="px-4 py-2 text-xs text-slate-600">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs rounded-xl font-semibold">Lưu Khóa Học</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL STUDENT (BẮT BUỘC NHẬP SĐT) */}
      {studentModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200 my-auto">
            <h3 className="text-base font-bold mb-4">{studentModal.mode === 'add' ? 'Thêm Hồ Sơ Học Sinh' : 'Sửa Hồ Sơ Học Sinh'}</h3>
            <form onSubmit={handleSaveStudent} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Mã HS *</label>
                <input type="text" required placeholder="HS001" value={studentModal.data.id} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, id: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm uppercase font-mono" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Họ tên học sinh *</label>
                <input type="text" required placeholder="Họ tên" value={studentModal.data.name} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Khóa học / Lớp *</label>
                <select value={studentModal.data.className} onChange={(e) => handleCourseSelectInStudentModal(e.target.value)} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white">
                  {courses.map((c) => <option key={c.id} value={c.name}>{c.name} — {formatVND(c.fee)}</option>)}
                  {!courses.some((c) => c.name === studentModal.data.className) && studentModal.data.className && (
                    <option value={studentModal.data.className}>{studentModal.data.className}</option>
                  )}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Học phí áp dụng (VNĐ/tháng) *</label>
                <input type="number" required placeholder="Học phí" value={studentModal.data.tuitionFee} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, tuitionFee: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Họ tên phụ huynh</label>
                <input type="text" placeholder="Phụ huynh" value={studentModal.data.parentName || ''} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, parentName: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Số điện thoại phụ huynh (Bắt buộc) <span className="text-rose-500">*</span></label>
                <input type="tel" required placeholder="VD: 0932766613" value={studentModal.data.phone} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, phone: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm focus:ring-2 focus:ring-indigo-500" />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setStudentModal({ ...studentModal, open: false })} className="px-4 py-2 text-xs text-slate-600">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs rounded-xl font-semibold">Lưu Hồ Sơ</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}