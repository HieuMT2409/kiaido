import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Users,
  DollarSign,
  Bell,
  CheckCircle2,
  XCircle,
  Search,
  Plus,
  Calendar,
  TrendingUp,
  Edit2,
  Trash2,
  ChevronLeft,
  ChevronRight,
  AlertTriangle,
  Download,
  Upload,
  FileSpreadsheet,
  Printer,
  Check,
  X,
  CreditCard,
  Cloud,
  CloudOff,
  BookOpen,
  Settings,
  RefreshCw
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
  writeBatch,
  Firestore
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

let db: Firestore | null = null;
try {
  if (firebaseConfig.apiKey && firebaseConfig.projectId) {
    const app = initializeApp(firebaseConfig);
    db = getFirestore(app);
  }
} catch (error) {
  console.error('Lỗi khởi tạo Firebase:', error);
}

// ==========================================
// 2. DỮ LIỆU MẪU BAN ĐẦU
// ==========================================
const INITIAL_COURSES: any[] | (() => any[]) = [];

const INITIAL_STUDENTS: any[] | (() => any[]) = [];

const INITIAL_PAYMENTS = {};

// Hàm tiện ích định dạng tiền tệ VNĐ
const formatVND = (amount: string | number) => {
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(Number(amount) || 0);
};

// Hàm tiện ích hiển thị ngày giờ
const formatDateTime = (isoStr: string | number | Date) => {
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
  // States dữ liệu chính
  const [courses, setCourses] = useState(INITIAL_COURSES);
  const [students, setStudents] = useState(INITIAL_STUDENTS);
  const [payments, setPayments] = useState(INITIAL_PAYMENTS);

  // States trạng thái Firebase
  const [firebaseStatus, setFirebaseStatus] = useState(db ? 'connecting' : 'offline'); // 'connecting' | 'connected' | 'error' | 'offline'
  const [firebaseErrorMsg, setFirebaseErrorMsg] = useState('');

  // States bộ lọc và điều hướng
  const [selectedMonth, setSelectedMonth] = useState('2026-10');
  const [activeTab, setActiveTab] = useState('tuition'); // 'tuition' | 'students' | 'courses' | 'revenue'
  const [searchQuery, setSearchQuery] = useState('');
  const [classFilter, setClassFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL'); // 'ALL' | 'PAID' | 'UNPAID'

  // States Modals
  const [showNotifPanel, setShowNotifPanel] = useState(false);
  const [paymentModal, setPaymentModal] = useState({ open: false, student: null, paidDate: '', method: 'Chuyển khoản', note: '' });
  const [studentModal, setStudentModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '' } });
  const [courseModal, setCourseModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' }, updateStudentsFee: true });
  const [receiptModal, setReceiptModal] = useState({ open: false, student: null, payment: null });
  const [excelModal, setExcelModal] = useState({ open: false, previewRows: [], fileName: '' });
  const [toast, setToast] = useState(null);

  const fileInputRef = useRef(null);

  const showToast = (message: string, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3500);
  };

  // ==========================================
  // 3. ĐỒNG BỘ FIREBASE FIRESTORE TRỰC TIẾP
  // ==========================================
  useEffect(() => {
    if (!db) {
      setFirebaseStatus('offline');
      return;
    }

    // Lắng nghe Collection: courses
    const unsubCourses = onSnapshot(
      collection(db, 'courses'),
      (snapshot) => {
        setFirebaseStatus('connected');
        setFirebaseErrorMsg('');
        if (!snapshot.empty) {
          const list = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
          setCourses(list);
        } else {
          // Tự động đẩy khóa học mẫu lên Firebase trong lần đầu tiên
          INITIAL_COURSES.forEach((c) => {
            setDoc(doc(db, 'courses', c.id), c).catch(console.error);
          });
        }
      },
      (err) => {
        console.error('Lỗi Firestore (courses):', err);
        setFirebaseStatus('error');
        setFirebaseErrorMsg(err.message);
      }
    );

    // Lắng nghe Collection: students
    const unsubStudents = onSnapshot(
      collection(db, 'students'),
      (snapshot) => {
        setFirebaseStatus('connected');
        if (!snapshot.empty) {
          const list = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
          list.sort((a, b) => a.id.localeCompare(b.id));
          setStudents(list);
        } else {
          // Đẩy học sinh mẫu lên Firebase nếu database hoàn toàn trống
          INITIAL_STUDENTS.forEach((s) => {
            setDoc(doc(db, 'students', s.id), s).catch(console.error);
          });
        }
      },
      (err) => {
        console.error('Lỗi Firestore (students):', err);
        setFirebaseStatus('error');
        setFirebaseErrorMsg(err.message);
      }
    );

    // Lắng nghe Collection: payments
    const unsubPayments = onSnapshot(
      collection(db, 'payments'),
      (snapshot) => {
        if (!snapshot.empty) {
          const payMap = {};
          snapshot.docs.forEach((d) => {
            payMap[d.id] = d.data();
          });
          setPayments(payMap);
        } else {
          Object.entries(INITIAL_PAYMENTS).forEach(([monthKey, monthData]) => {
            setDoc(doc(db, 'payments', monthKey), monthData).catch(console.error);
          });
        }
      },
      (err) => console.error('Lỗi Firestore (payments):', err)
    );

    return () => {
      unsubCourses();
      unsubStudents();
      unsubPayments();
    };
  }, []);

  // Hàm đồng bộ thủ công toàn bộ dữ liệu hiện tại lên Firebase
  const handleForceSyncToFirebase = async () => {
    if (!db) {
      showToast('Chưa cấu hình file .env cho Firebase!', 'error');
      return;
    }
    try {
      const batch = writeBatch(db);
      courses.forEach((c) => batch.set(doc(db, 'courses', c.id), c));
      students.forEach((s) => batch.set(doc(db, 'students', s.id), s));
      Object.entries(payments).forEach(([mKey, mData]) => batch.set(doc(db, 'payments', mKey), mData));
      await batch.commit();
      showToast('Đã đẩy toàn bộ dữ liệu lên Firebase thành công!');
    } catch (error) {
      showToast('Lỗi ghi Firebase: ' + error.message, 'error');
    }
  };

  // ==========================================
  // 4. TÍNH TOÁN DỮ LIỆU & THỐNG KÊ
  // ==========================================
  const currentMonthPayments = payments[selectedMonth] || {};

  const enrichedStudents = useMemo(() => {
    return students.map((student) => {
      const paymentInfo = currentMonthPayments[student.id] || {
        paid: false,
        paidDate: null,
        method: '',
        note: ''
      };
      return {
        ...student,
        paid: Boolean(paymentInfo.paid),
        paidDate: paymentInfo.paidDate || null,
        method: paymentInfo.method || '',
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
    const completionRate = expectedStudentsCount(enrichedStudents)
      ? Math.round((paidStudents.length / enrichedStudents.length) * 100)
      : 0;
    return { expectedRevenue, collectedRevenue, pendingRevenue, completionRate };
  }, [enrichedStudents, paidStudents]);

  function expectedStudentsCount(arr: string | any[]) {
    return arr && arr.length > 0;
  }

  const filteredStudents = useMemo(() => {
    return enrichedStudents.filter((s) => {
      const matchSearch =
        s.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        s.id.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (s.phone && s.phone.includes(searchQuery));
      const matchClass = classFilter === 'ALL' || s.className === classFilter;
      const matchStatus =
        statusFilter === 'ALL' ||
        (statusFilter === 'PAID' && s.paid) ||
        (statusFilter === 'UNPAID' && !s.paid);
      return matchSearch && matchClass && matchStatus;
    });
  }, [enrichedStudents, searchQuery, classFilter, statusFilter]);

  // Chuyển tháng
  const changeMonth = (offset: number) => {
    const [year, month] = selectedMonth.split('-').map(Number);
    const date = new Date(year, month - 1 + offset, 1);
    const newYear = date.getFullYear();
    const newMonth = String(date.getMonth() + 1).padStart(2, '0');
    setSelectedMonth(`${newYear}-${newMonth}`);
  };

  // ==========================================
  // 5. XỬ LÝ THU HỌC PHÍ
  // ==========================================
  const saveMonthPaymentData = async (monthKey: string, updatedMonthObj: any) => {
    setPayments((prev) => ({ ...prev, [monthKey]: updatedMonthObj }));
    if (db) {
      try {
        await setDoc(doc(db, 'payments', monthKey), updatedMonthObj);
      } catch (err) {
        showToast('Lỗi lưu Firebase: ' + err.message, 'error');
      }
    }
  };

  const handleQuickMarkPaid = async (student: { id: any; tuitionFee: any; name: any; }) => {
    const nowIso = new Date().toISOString().slice(0, 16);
    const updatedMonth = {
      ...(payments[selectedMonth] || {}),
      [student.id]: {
        paid: true,
        paidDate: nowIso,
        method: 'Chuyển khoản',
        note: 'Admin xác nhận nhanh',
        amount: Number(student.tuitionFee)
      }
    };
    await saveMonthPaymentData(selectedMonth, updatedMonth);
    showToast(`Đã ghi nhận ${student.name} đóng học phí!`);
  };

  const handleMarkUnpaid = async (student: { id: any; tuitionFee: any; name: any; }) => {
    const updatedMonth = {
      ...(payments[selectedMonth] || {}),
      [student.id]: {
        paid: false,
        paidDate: null,
        method: '',
        note: '',
        amount: Number(student.tuitionFee)
      }
    };
    await saveMonthPaymentData(selectedMonth, updatedMonth);
    showToast(`Đã chuyển ${student.name} về trạng thái Chưa đóng.`, 'warning');
  };

  const openCustomPaymentModal = (student: { paidDate: any; method: any; note: any; }) => {
    const nowIso = new Date().toISOString().slice(0, 16);
    setPaymentModal({
      open: true,
      student,
      paidDate: student.paidDate || nowIso,
      method: student.method || 'Chuyển khoản',
      note: student.note || ''
    });
  };

  const handleSaveCustomPayment = async (e: { preventDefault: () => void; }) => {
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
    showToast(`Đã lưu thông tin thu tiền của ${student.name}`);
  };

  // ==========================================
  // 6. QUẢN LÝ KHÓA HỌC & GIÁ TIỀN (MỚI)
  // ==========================================
  const openAddCourseModal = () => {
    setCourseModal({
      open: true,
      mode: 'add',
      data: { id: `COURSE_${Date.now()}`, name: '', fee: 1500000, description: '' },
      updateStudentsFee: true
    });
  };

  const openEditCourseModal = (course: { name: any; id?: string; fee?: string; description?: string; }) => {
    setCourseModal({
      open: true,
      mode: 'edit',
      oldName: course.name,
      data: { ...course },
      updateStudentsFee: true
    });
  };

  const handleSaveCourse = async (e: { preventDefault: () => void; }) => {
    e.preventDefault();
    const { mode, data, oldName, updateStudentsFee } = courseModal;
    const cleanCourse = {
      id: data.id,
      name: data.name.trim(),
      fee: Number(data.fee) || 0,
      description: data.description || ''
    };

    if (!cleanCourse.name) return;

    // Cập nhật state courses
    if (mode === 'add') {
      setCourses((prev) => [...prev, cleanCourse]);
    } else {
      setCourses((prev) => prev.map((c) => (c.id === cleanCourse.id ? cleanCourse : c)));
    }

    // Lưu khóa học lên Firebase
    if (db) {
      try {
        await setDoc(doc(db, 'courses', cleanCourse.id), cleanCourse);
      } catch (err) {
        showToast('Lỗi lưu khóa học lên Firebase: ' + err.message, 'error');
      }
    }

    // Nếu sửa khóa học và chọn tự động cập nhật học phí/tên lớp cho học sinh đang theo học
    if (mode === 'edit' && updateStudentsFee) {
      const affectedStudents = students.filter((s) => s.className === oldName || s.className === cleanCourse.name);
      if (affectedStudents.length > 0) {
        const updatedStudents = students.map((s) => {
          if (s.className === oldName || s.className === cleanCourse.name) {
            return { ...s, className: cleanCourse.name, tuitionFee: cleanCourse.fee };
          }
          return s;
        });
        setStudents(updatedStudents);

        if (db) {
          const batch = writeBatch(db);
          affectedStudents.forEach((s) => {
            batch.set(doc(db, 'students', s.id), {
              ...s,
              className: cleanCourse.name,
              tuitionFee: cleanCourse.fee
            });
          });
          await batch.commit().catch(console.error);
        }
      }
    }

    setCourseModal({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' }, updateStudentsFee: true });
    showToast(mode === 'add' ? `Đã thêm khóa học "${cleanCourse.name}"` : `Đã cập nhật khóa học "${cleanCourse.name}"`);
  };

  const handleDeleteCourse = async (course: { name: any; id: string; }) => {
    if (!window.confirm(`Bạn có chắc muốn xóa khóa học "${course.name}"?`)) return;
    setCourses((prev) => prev.filter((c) => c.id !== course.id));
    if (db) {
      await deleteDoc(doc(db, 'courses', course.id)).catch(console.error);
    }
    showToast(`Đã xóa khóa học ${course.name}`, 'warning');
  };

  // ==========================================
  // 7. QUẢN LÝ HỌC SINH
  // ==========================================
  const openAddStudentModal = () => {
    const nextNum = students.length + 1;
    const defaultId = `HS${String(nextNum).padStart(3, '0')}`;
    const defaultCourse = courses[0] || { name: 'Lớp Cơ Bản', fee: 1200000 };
    setStudentModal({
      open: true,
      mode: 'add',
      data: {
        id: defaultId,
        name: '',
        className: defaultCourse.name,
        phone: '',
        parentName: '',
        tuitionFee: defaultCourse.fee
      }
    });
  };

  const openEditStudentModal = (student: { id: any; name: any; className: any; phone: any; parentName: any; tuitionFee: any; }) => {
    setStudentModal({
      open: true,
      mode: 'edit',
      data: {
        id: student.id,
        name: student.name,
        className: student.className,
        phone: student.phone,
        parentName: student.parentName,
        tuitionFee: student.tuitionFee
      }
    });
  };

  // Khi chọn khóa học trong form Học sinh -> Tự động điền giá tiền tương ứng
  const handleCourseSelectInStudentModal = (selectedCourseName: string) => {
    const foundCourse = courses.find((c) => c.name === selectedCourseName);
    setStudentModal((prev) => ({
      ...prev,
      data: {
        ...prev.data,
        className: selectedCourseName,
        tuitionFee: foundCourse ? foundCourse.fee : prev.data.tuitionFee
      }
    }));
  };

  const handleSaveStudent = async (e: { preventDefault: () => void; }) => {
    e.preventDefault();
    const { mode, data } = studentModal;
    if (!data.name.trim() || !data.id.trim()) return;

    const studentObj = {
      id: data.id.trim().toUpperCase(),
      name: data.name.trim(),
      className: data.className.trim(),
      phone: data.phone.trim(),
      parentName: data.parentName.trim(),
      tuitionFee: Number(data.tuitionFee) || 0,
      status: 'active'
    };

    if (mode === 'add') {
      if (students.some((s) => s.id === studentObj.id)) {
        showToast('Mã học sinh này đã tồn tại!', 'error');
        return;
      }
      setStudents((prev) => [...prev, studentObj]);
    } else {
      setStudents((prev) => prev.map((s) => (s.id === studentObj.id ? studentObj : s)));
    }

    if (db) {
      try {
        await setDoc(doc(db, 'students', studentObj.id), studentObj);
        showToast(mode === 'add' ? 'Đã thêm học sinh lên Firebase!' : 'Đã cập nhật học sinh trên Firebase!');
      } catch (err) {
        showToast('Lỗi ghi Firebase: ' + err.message, 'error');
      }
    } else {
      showToast(mode === 'add' ? 'Đã thêm học sinh (Lưu cục bộ)!' : 'Đã cập nhật thông tin học sinh!');
    }

    setStudentModal({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '' } });
  };

  const handleDeleteStudent = async (studentId: string) => {
    if (!window.confirm('Bạn có chắc chắn muốn xóa học sinh này?')) return;
    setStudents((prev) => prev.filter((s) => s.id !== studentId));
    if (db) {
      await deleteDoc(doc(db, 'students', studentId)).catch(console.error);
    }
    showToast('Đã xóa học sinh khỏi danh sách.', 'warning');
  };

  // ==========================================
  // 8. NHẬP / XUẤT EXCEL (.XLSX)
  // ==========================================
  const handleExportExcel = () => {
    const studentRows = enrichedStudents.map((s, idx) => ({
      'STT': idx + 1,
      'Mã HS': s.id,
      'Họ và Tên': s.name,
      'Khóa học / Lớp': s.className,
      'Phụ huynh': s.parentName,
      'Số điện thoại': s.phone,
      'Học phí (VNĐ)': Number(s.tuitionFee),
      [`Trạng thái (${selectedMonth})`]: s.paid ? 'Đã đóng' : 'Chưa đóng',
      'Ngày đóng tiền': s.paidDate ? formatDateTime(s.paidDate) : '',
      'Phương thức': s.method || '',
      'Ghi chú': s.note || ''
    }));

    const courseRows = courses.map((c, idx) => ({
      'STT': idx + 1,
      'Mã Khóa Học': c.id,
      'Tên Khóa Học': c.name,
      'Học Phí Mặc Định (VNĐ)': Number(c.fee),
      'Mô tả': c.description
    }));

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(studentRows);
    const ws2 = XLSX.utils.json_to_sheet(courseRows);
    XLSX.utils.book_append_sheet(wb, ws1, `DS_HocSinh_${selectedMonth}`);
    XLSX.utils.book_append_sheet(wb, ws2, 'DS_KhoaHoc');
    XLSX.writeFile(wb, `QuanLyHocPhi_${selectedMonth}.xlsx`);
    showToast('Đã xuất file Excel thành công!');
  };

  const handleDownloadTemplate = () => {
    const sampleRows = [
      {
        'Mã HS': 'HS010',
        'Họ và Tên': 'Nguyễn Văn Mẫu',
        'Khóa học / Lớp': courses[0]?.name || 'Toán 10 - Nâng cao',
        'Phụ huynh': 'Nguyễn Văn Bố',
        'Số điện thoại': '0909123456',
        'Học phí (VNĐ)': courses[0]?.fee || 1200000
      }
    ];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(sampleRows);
    XLSX.utils.book_append_sheet(wb, ws, 'Mau_Nhap_Hoc_Sinh');
    XLSX.writeFile(wb, 'File_Mau_Nhap_Hoc_Sinh.xlsx');
  };

  const handleFileUpload = (e: { target: { files: any[]; value: null; }; }) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(ws);

      const parsed = rawData
        .map((row, index) => {
          const id = String(row['Mã HS'] || row['Ma HS'] || `HS${String(students.length + index + 1).padStart(3, '0')}`).trim().toUpperCase();
          const name = String(row['Họ và Tên'] || row['Ho va Ten'] || row['Tên'] || '').trim();
          const className = String(row['Khóa học / Lớp'] || row['Lớp'] || courses[0]?.name || 'Chưa xếp lớp').trim();
          const matchedCourse = courses.find((c) => c.name.toLowerCase() === className.toLowerCase());
          const rawFee = row['Học phí (VNĐ)'] || row['Học phí'] || (matchedCourse ? matchedCourse.fee : 1200000);
          const tuitionFee = Number(String(rawFee).replace(/[^0-9]/g, '')) || 1200000;

          return {
            id,
            name,
            className,
            parentName: String(row['Phụ huynh'] || '').trim(),
            phone: String(row['Số điện thoại'] || row['SĐT'] || '').trim(),
            tuitionFee,
            status: 'active'
          };
        })
        .filter((item) => item.name !== '');

      setExcelModal({ open: true, previewRows: parsed, fileName: file.name });
      e.target.value = null;
    };
    reader.readAsBinaryString(file);
  };

  const confirmExcelImport = async () => {
    const { previewRows } = excelModal;
    if (!previewRows.length) return;

    const mapById = new Map(students.map((s) => [s.id, s]));
    previewRows.forEach((row) => mapById.set(row.id, row));
    const mergedList = Array.from(mapById.values());
    setStudents(mergedList);

    if (db) {
      const batch = writeBatch(db);
      previewRows.forEach((row) => batch.set(doc(db, 'students', row.id), row));
      await batch.commit().catch(console.error);
    }

    setExcelModal({ open: false, previewRows: [], fileName: '' });
    showToast(`Đã nhập thành công ${previewRows.length} học sinh!`);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col">
      {/* TOAST THÔNG BÁO */}
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg text-white text-sm font-medium bg-slate-900 animate-bounce">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast.message}</span>
        </div>
      )}

      {/* BANNER LỖI QUYỀN FIREBASE NẾU CÓ */}
      {firebaseStatus === 'error' && (
        <div className="bg-rose-600 text-white px-4 py-2 text-xs sm:text-sm flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>
              <strong>Lỗi kết nối Firebase Firestore:</strong> {firebaseErrorMsg}. Hãy kiểm tra lại mục <strong>Firestore Database &gt; Rules</strong> và đổi thành <code>allow read, write: if true;</code>
            </span>
          </div>
        </div>
      )}

      {/* HEADER CHÍNH */}
      <header className="bg-black border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3.5 flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-3">
              <img src="/logo.jpg" alt="Logo" className="w-16 h-16 object-contain" />
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-lg font-bold text-red-500">KIAIDO Manager</h1>
                {firebaseStatus === 'connected' ? (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                    <Cloud className="w-3 h-3" /> Đã kết nối Firebase
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-amber-50 text-amber-700 border border-amber-200">
                    <CloudOff className="w-3 h-3" /> Lưu cục bộ (Offline)
                  </span>
                )}
              </div>
              <p className="text-xs text-slate-500">Hệ thống quản lý học sinh, khóa học & thu học phí</p>
            </div>
          </div>

          {/* BỘ CHỌN THÁNG & CÔNG CỤ */}
          <div className="flex items-center gap-2.5 flex-wrap">
            {db && (
              <button
                onClick={handleForceSyncToFirebase}
                title="Đẩy toàn bộ dữ liệu hiện tại lên Firebase"
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-100 hover:bg-slate-200 text-slate-700 transition"
              >
                <RefreshCw className="w-3.5 h-3.5" /> Đồng bộ lên Cloud
              </button>
            )}

            <div className="flex items-center bg-slate-100 rounded-xl p-1 border border-slate-200">
              <button onClick={() => changeMonth(-1)} className="p-1.5 hover:bg-white rounded-lg transition">
                <ChevronLeft className="w-4 h-4" />
              </button>
              <div className="flex items-center gap-1.5 px-3 text-sm font-semibold text-slate-800">
                <Calendar className="w-4 h-4 text-indigo-600" />
                <input
                  type="month"
                  value={selectedMonth}
                  onChange={(e) => setSelectedMonth(e.target.value)}
                  className="bg-transparent border-none focus:outline-none cursor-pointer"
                />
              </div>
              <button onClick={() => changeMonth(1)} className="p-1.5 hover:bg-white rounded-lg transition">
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>

            {/* CHUÔNG NHẮC NHỞ */}
            <div className="relative">
              <button
                onClick={() => setShowNotifPanel(!showNotifPanel)}
                className={`relative p-2.5 rounded-xl border transition ${
                  unpaidStudents.length > 0
                    ? 'bg-rose-50 border-rose-200 text-rose-600 hover:bg-rose-100'
                    : 'bg-slate-100 border-slate-200 text-slate-600'
                }`}
              >
                <Bell className="w-5 h-5" />
                {unpaidStudents.length > 0 && (
                  <span className="absolute -top-1.5 -right-1.5 bg-rose-600 text-white text-[11px] font-bold w-5 h-5 rounded-full flex items-center justify-center">
                    {unpaidStudents.length}
                  </span>
                )}
              </button>

              {showNotifPanel && (
                <div className="absolute right-0 mt-2 w-80 sm:w-96 bg-white rounded-2xl shadow-xl border border-slate-200 z-50 overflow-hidden">
                  <div className="p-3.5 bg-rose-50 border-b border-rose-100 flex items-center justify-between">
                    <span className="font-semibold text-sm text-rose-800">
                      Nhắc nhở: {unpaidStudents.length} HS chưa đóng tháng {selectedMonth}
                    </span>
                    <button onClick={() => setShowNotifPanel(false)} className="text-slate-400 hover:text-slate-600">
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="max-h-72 overflow-y-auto divide-y divide-slate-100">
                    {unpaidStudents.length === 0 ? (
                      <div className="p-6 text-center text-sm text-slate-500">Tất cả học sinh đã hoàn thành học phí!</div>
                    ) : (
                      unpaidStudents.map((s) => (
                        <div key={s.id} className="p-3 flex items-center justify-between hover:bg-slate-50">
                          <div>
                            <p className="text-sm font-semibold text-slate-800">{s.name}</p>
                            <p className="text-xs text-slate-500">{s.className} • {formatVND(s.tuitionFee)}</p>
                          </div>
                          <button
                            onClick={() => {
                              handleQuickMarkPaid(s);
                              setShowNotifPanel(false);
                            }}
                            className="px-2.5 py-1 bg-emerald-600 text-white text-xs font-medium rounded-lg hover:bg-emerald-700"
                          >
                            Thu ngay
                          </button>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto w-full px-4 sm:px-6 py-6 space-y-6 flex-1">
        {/* BANNER NHẮC NHỞ ADMIN */}
        {unpaidStudents.length > 0 && (
          <div className="bg-gradient-to-r from-amber-50 via-orange-50 to-rose-50 border border-amber-200/80 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              <div className="p-2.5 bg-amber-500 text-white rounded-xl shrink-0 mt-0.5">
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div>
                <h2 className="text-sm sm:text-base font-bold text-slate-900">
                  Cần thu học phí: Còn {unpaidStudents.length} học sinh chưa đóng tiền tháng {selectedMonth}
                </h2>
                <p className="text-xs sm:text-sm text-slate-600 mt-0.5">
                  Tổng số tiền cần thu thêm: <strong className="text-rose-600">{formatVND(stats.pendingRevenue)}</strong>
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <button
                onClick={() => {
                  setActiveTab('tuition');
                  setStatusFilter('UNPAID');
                }}
                className="px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-semibold rounded-xl transition"
              >
                Lọc DS chưa đóng ({unpaidStudents.length})
              </button>
            </div>
          </div>
        )}

        {/* 4 THẺ THỐNG KÊ DOANH THU */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
            <p className="text-xs font-medium text-slate-500">Doanh thu dự kiến ({selectedMonth})</p>
            <p className="text-xl font-bold text-slate-900 mt-1">{formatVND(stats.expectedRevenue)}</p>
            <p className="text-xs text-slate-400 mt-1">Tổng {students.length} học sinh</p>
          </div>
          <div className="bg-white p-4 rounded-2xl border border-emerald-200 shadow-sm">
            <p className="text-xs font-medium text-emerald-700">Thực thu (Đã đóng)</p>
            <p className="text-xl font-bold text-emerald-600 mt-1">{formatVND(stats.collectedRevenue)}</p>
            <p className="text-xs text-emerald-600 mt-1">{paidStudents.length} học sinh đã hoàn thành</p>
          </div>
          <div className="bg-white p-4 rounded-2xl border border-rose-200 shadow-sm">
            <p className="text-xs font-medium text-rose-700">Còn thiếu (Chưa đóng)</p>
            <p className="text-xl font-bold text-rose-600 mt-1">{formatVND(stats.pendingRevenue)}</p>
            <p className="text-xs text-rose-500 mt-1">{unpaidStudents.length} học sinh cần nhắc nhở</p>
          </div>
          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
            <div className="flex justify-between text-xs font-medium text-slate-500">
              <span>Tiến độ thu tháng {selectedMonth}</span>
              <span className="font-bold text-indigo-600">{stats.completionRate}%</span>
            </div>
            <div className="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden mt-3">
              <div className="h-full bg-indigo-600 rounded-full transition-all" style={{ width: `${stats.completionRate}%` }} />
            </div>
            <p className="text-xs text-slate-400 mt-2">Đang có {courses.length} khóa học hoạt động</p>
          </div>
        </div>

        {/* THANH ĐIỀU HƯỚNG TABS & NÚT HÀNH ĐỘNG */}
        <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 pb-3">
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setActiveTab('tuition')}
              className={`flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === 'tuition' ? 'bg-indigo-600 text-white shadow-sm' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
              }`}
            >
              <CheckCircle2 className="w-4 h-4" /> Thu Học Phí Theo Tháng
            </button>
            <button
              onClick={() => setActiveTab('students')}
              className={`flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === 'students' ? 'bg-indigo-600 text-white shadow-sm' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
              }`}
            >
              <Users className="w-4 h-4" /> Hồ Sơ Học Sinh ({students.length})
            </button>
            <button
              onClick={() => setActiveTab('courses')}
              className={`flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-semibold transition ${
                activeTab === 'courses' ? 'bg-indigo-600 text-white shadow-sm' : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
              }`}
            >
              <BookOpen className="w-4 h-4" /> Khóa Học & Giá Tiền ({courses.length})
            </button>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <input type="file" accept=".xlsx,.xls,.csv" ref={fileInputRef} onChange={handleFileUpload} className="hidden" />
            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1.5 px-3 py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold"
            >
              <Upload className="w-4 h-4 text-indigo-600" /> Nhập Excel
            </button>
            <button
              onClick={handleExportExcel}
              className="flex items-center gap-1.5 px-3 py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold"
            >
              <Download className="w-4 h-4 text-emerald-600" /> Xuất Excel
            </button>
            <button
              onClick={openAddCourseModal}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-amber-500 hover:bg-amber-600 text-white rounded-xl text-xs font-semibold shadow-sm"
            >
              <Settings className="w-4 h-4" /> + Thêm Khóa Học
            </button>
            <button
              onClick={openAddStudentModal}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold shadow-sm"
            >
              <Plus className="w-4 h-4" /> + Thêm Học Sinh
            </button>
          </div>
        </div>

        {/* ====================================================
            TAB 3: QUẢN LÝ KHÓA HỌC & GIÁ TIỀN (MỚI)
        ==================================================== */}
        {activeTab === 'courses' && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-200 flex items-center justify-between bg-slate-50/60">
              <div>
                <h3 className="font-bold text-slate-800">Danh mục Khóa học & Định mức Học phí</h3>
                <p className="text-xs text-slate-500">
                  Khi chỉnh sửa giá tiền khóa học, bạn có thể tự động cập nhật mức học phí mới cho toàn bộ học sinh đang học lớp đó.
                </p>
              </div>
              <button
                onClick={openAddCourseModal}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold"
              >
                <Plus className="w-4 h-4" /> Thêm khóa học mới
              </button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                    <th className="py-3.5 px-4">Tên Khóa Học / Lớp</th>
                    <th className="py-3.5 px-4">Học Phí Mặc Định / Tháng</th>
                    <th className="py-3.5 px-4">Số HS Đang Học</th>
                    <th className="py-3.5 px-4">Mô tả ghi chú</th>
                    <th className="py-3.5 px-4 text-right">Chỉnh sửa / Xóa</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-sm">
                  {courses.map((course) => {
                    const studentCount = students.filter((s) => s.className === course.name).length;
                    return (
                      <tr key={course.id} className="hover:bg-slate-50">
                        <td className="py-3.5 px-4 font-semibold text-slate-900">{course.name}</td>
                        <td className="py-3.5 px-4 font-bold text-indigo-600">{formatVND(course.fee)}</td>
                        <td className="py-3.5 px-4">
                          <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-700">
                            {studentCount} học sinh
                          </span>
                        </td>
                        <td className="py-3.5 px-4 text-slate-500 text-xs">{course.description || '—'}</td>
                        <td className="py-3.5 px-4 text-right space-x-2">
                          <button
                            onClick={() => openEditCourseModal(course)}
                            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-indigo-50 text-indigo-700 hover:bg-indigo-100"
                          >
                            <Edit2 className="w-3.5 h-3.5" /> Sửa giá / Tên
                          </button>
                          <button
                            onClick={() => handleDeleteCourse(course)}
                            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-rose-50 text-rose-600 hover:bg-rose-100"
                          >
                            <Trash2 className="w-3.5 h-3.5" /> Xóa
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ====================================================
            TAB 1 & TAB 2: THU HỌC PHÍ VÀ DANH SÁCH HỌC SINH
        ==================================================== */}
        {(activeTab === 'tuition' || activeTab === 'students') && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            {/* BỘ LỌC TÌM KIẾM */}
            <div className="p-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 bg-slate-50/50">
              <div className="relative flex-1 min-w-[240px]">
                <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Tìm theo tên học sinh, mã HS, số điện thoại..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 bg-white border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div className="flex items-center gap-2 flex-wrap">
                <select
                  value={classFilter}
                  onChange={(e) => setClassFilter(e.target.value)}
                  className="px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-700"
                >
                  <option value="ALL">Tất cả khóa học ({courses.length})</option>
                  {courses.map((c) => (
                    <option key={c.id} value={c.name}>
                      {c.name} ({formatVND(c.fee)})
                    </option>
                  ))}
                </select>

                <div className="flex bg-slate-200/70 p-1 rounded-xl text-xs font-semibold">
                  <button
                    onClick={() => setStatusFilter('ALL')}
                    className={`px-3 py-1.5 rounded-lg transition ${statusFilter === 'ALL' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600'}`}
                  >
                    Tất cả ({enrichedStudents.length})
                  </button>
                  <button
                    onClick={() => setStatusFilter('PAID')}
                    className={`px-3 py-1.5 rounded-lg transition ${statusFilter === 'PAID' ? 'bg-emerald-600 text-white shadow-sm' : 'text-slate-600'}`}
                  >
                    Đã đóng ({paidStudents.length})
                  </button>
                  <button
                    onClick={() => setStatusFilter('UNPAID')}
                    className={`px-3 py-1.5 rounded-lg transition ${statusFilter === 'UNPAID' ? 'bg-rose-600 text-white shadow-sm' : 'text-slate-600'}`}
                  >
                    Chưa đóng ({unpaidStudents.length})
                  </button>
                </div>
              </div>
            </div>

            {/* BẢNG DỮ LIỆU */}
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                    <th className="py-3.5 px-4">Mã HS</th>
                    <th className="py-3.5 px-4">Họ và Tên</th>
                    <th className="py-3.5 px-4">Khóa học / Lớp</th>
                    <th className="py-3.5 px-4">Học phí</th>
                    <th className="py-3.5 px-4">Trạng thái ({selectedMonth})</th>
                    <th className="py-3.5 px-4">Ngày đóng tiền</th>
                    <th className="py-3.5 px-4 text-right">Thao tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 text-sm">
                  {filteredStudents.map((student) => (
                    <tr key={student.id} className="hover:bg-slate-50/80 transition">
                      <td className="py-3.5 px-4 font-mono text-xs font-bold text-slate-500">{student.id}</td>
                      <td className="py-3.5 px-4">
                        <p className="font-semibold text-slate-900">{student.name}</p>
                        <p className="text-xs text-slate-500">
                          PH: {student.parentName || '—'} •{' '}
                          {student.phone ? (
                            <a
                              href={`https://zalo.me/${student.phone.replace(/[^0-9]/g, '')}`}
                              target="_blank"
                              rel="noopener noreferrer"
                              title="Bấm để nhắn tin / gọi Zalo phụ huynh"
                              className="text-indigo-600 font-medium hover:underline inline-flex items-center gap-0.5"
                            >{student.phone} (Zalo)</a>) : ('—')}
                        </p>
                      </td>
                      <td className="py-3.5 px-4">
                        <span className="px-2.5 py-1 rounded-lg text-xs font-medium bg-indigo-50 text-indigo-700">
                          {student.className}
                        </span>
                      </td>
                      <td className="py-3.5 px-4 font-bold text-slate-800">{formatVND(student.tuitionFee)}</td>
                      <td className="py-3.5 px-4">
                        {student.paid ? (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                            <CheckCircle2 className="w-3.5 h-3.5" /> Đã đóng
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                            <XCircle className="w-3.5 h-3.5" /> Chưa đóng
                          </span>
                        )}
                      </td>
                      <td className="py-3.5 px-4 text-xs text-slate-600">
                        {student.paid ? (
                          <div>
                            <p className="font-medium text-slate-800">{formatDateTime(student.paidDate)}</p>
                            <p className="text-[11px] text-slate-400">{student.method} {student.note ? `• ${student.note}` : ''}</p>
                          </div>
                        ) : (
                          <span className="text-slate-400 italic">Chưa ghi nhận</span>
                        )}
                      </td>
                      <td className="py-3.5 px-4 text-right">
                        <div className="flex items-center justify-end gap-1.5">
                          {!student.paid ? (
                            <>
                              <button
                                onClick={() => handleQuickMarkPaid(student)}
                                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold rounded-lg shadow-sm"
                              >
                                Đánh dấu Đã đóng
                              </button>
                              <button
                                onClick={() => openCustomPaymentModal(student)}
                                title="Chọn ngày đóng cụ thể"
                                className="p-1.5 text-slate-500 hover:bg-slate-100 rounded-lg border border-slate-200"
                              >
                                <Calendar className="w-4 h-4" />
                              </button>
                            </>
                          ) : (
                            <>
                              <button
                                onClick={() => openCustomPaymentModal(student)}
                                className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded-lg"
                              >
                                Sửa ngày
                              </button>
                              <button
                                onClick={() => handleMarkUnpaid(student)}
                                className="px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 text-xs font-medium rounded-lg"
                              >
                                Hủy thu
                              </button>
                            </>
                          )}
                          <button
                            onClick={() => openEditStudentModal(student)}
                            title="Sửa học sinh / học phí"
                            className="p-1.5 text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg"
                          >
                            <Edit2 className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => handleDeleteStudent(student.id)}
                            title="Xóa học sinh"
                            className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {/* ====================================================
          MODAL 1: THÊM / SỬA KHÓA HỌC & GIÁ TIỀN
      ==================================================== */}
      {courseModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-base font-bold text-slate-900">
                {courseModal.mode === 'add' ? 'Thêm Khóa Học Mới' : 'Chỉnh Sửa Khóa Học & Học Phí'}
              </h3>
              <button onClick={() => setCourseModal({ ...courseModal, open: false })} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>
            <form onSubmit={handleSaveCourse} className="space-y-4 mt-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Tên Khóa Học / Lớp *</label>
                <input
                  type="text"
                  required
                  placeholder="VD: Toán 10 - Nâng cao"
                  value={courseModal.data.name}
                  onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, name: e.target.value } })}
                  className="w-full px-3.5 py-2 border border-slate-300 rounded-xl text-sm"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Mức Học Phí Mặc Định (VNĐ/tháng) *</label>
                <input
                  type="number"
                  required
                  step="10000"
                  value={courseModal.data.fee}
                  onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, fee: e.target.value } })}
                  className="w-full px-3.5 py-2 border border-slate-300 rounded-xl text-sm font-bold text-indigo-600"
                />
                <p className="text-[11px] text-slate-500 mt-1">Hiển thị: {formatVND(courseModal.data.fee)}</p>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Mô tả / Ghi chú</label>
                <input
                  type="text"
                  placeholder="Lịch học, giáo viên phụ trách..."
                  value={courseModal.data.description}
                  onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, description: e.target.value } })}
                  className="w-full px-3.5 py-2 border border-slate-300 rounded-xl text-sm"
                />
              </div>

              {courseModal.mode === 'edit' && (
                <label className="flex items-start gap-2.5 p-3 rounded-xl bg-amber-50 border border-amber-200 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={courseModal.updateStudentsFee}
                    onChange={(e) => setCourseModal({ ...courseModal, updateStudentsFee: e.target.checked })}
                    className="mt-1 rounded text-indigo-600"
                  />
                  <span className="text-xs text-amber-900">
                    <strong>Đồng bộ giá mới cho học sinh:</strong> Tự động cập nhật mức học phí mới ({formatVND(courseModal.data.fee)}) cho tất cả học sinh đang thuộc lớp này.
                  </span>
                </label>
              )}

              <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setCourseModal({ ...courseModal, open: false })}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100"
                >
                  Hủy
                </button>
                <button type="submit" className="px-4 py-2 rounded-xl text-xs font-semibold bg-indigo-600 text-white hover:bg-indigo-700">
                  Lưu Khóa Học
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ====================================================
          MODAL 2: THÊM / SỬA HỌC SINH (CHỌN KHÓA HỌC TỰ NHẢY GIÁ)
      ==================================================== */}
      {studentModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-base font-bold text-slate-900">
                {studentModal.mode === 'add' ? 'Thêm Học Sinh Mới' : 'Cập Nhật Hồ Sơ Học Sinh'}
              </h3>
              <button onClick={() => setStudentModal({ ...studentModal, open: false })} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>
            <form onSubmit={handleSaveStudent} className="space-y-3.5 mt-4">
              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Mã HS *</label>
                  <input
                    type="text"
                    required
                    disabled={studentModal.mode === 'edit'}
                    value={studentModal.data.id}
                    onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, id: e.target.value } })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm font-mono uppercase disabled:bg-slate-100"
                  />
                </div>
                <div className="col-span-2">
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Họ và Tên học sinh *</label>
                  <input
                    type="text"
                    required
                    value={studentModal.data.name}
                    onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, name: e.target.value } })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm"
                  />
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-semibold text-slate-700">Chọn Khóa Học / Lớp *</label>
                  <button
                    type="button"
                    onClick={() => {
                      setStudentModal({ ...studentModal, open: false });
                      openAddCourseModal();
                    }}
                    className="text-[11px] font-semibold text-indigo-600 hover:underline"
                  >
                    + Tạo khóa học mới
                  </button>
                </div>
                <select
                  value={studentModal.data.className}
                  onChange={(e) => handleCourseSelectInStudentModal(e.target.value)}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm bg-white"
                >
                  {courses.map((c) => (
                    <option key={c.id} value={c.name}>
                      {c.name} — {formatVND(c.fee)}
                    </option>
                  ))}
                  {!courses.some((c) => c.name === studentModal.data.className) && studentModal.data.className && (
                    <option value={studentModal.data.className}>{studentModal.data.className}</option>
                  )}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Học phí áp dụng cho HS này (VNĐ/tháng) *
                </label>
                <input
                  type="number"
                  required
                  step="10000"
                  value={studentModal.data.tuitionFee}
                  onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, tuitionFee: e.target.value } })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm font-bold text-indigo-600"
                />
                <p className="text-[11px] text-slate-400 mt-0.5">
                  Tự động lấy theo giá khóa học, hoặc bạn có thể sửa riêng nếu học sinh được giảm giá.
                </p>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Họ tên Phụ huynh</label>
                  <input
                    type="text"
                    value={studentModal.data.parentName}
                    onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, parentName: e.target.value } })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">Số điện thoại</label>
                  <input
                    type="text"
                    value={studentModal.data.phone}
                    onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, phone: e.target.value } })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm"
                  />
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setStudentModal({ ...studentModal, open: false })}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100"
                >
                  Hủy
                </button>
                <button type="submit" className="px-4 py-2 rounded-xl text-xs font-semibold bg-indigo-600 text-white hover:bg-indigo-700">
                  {studentModal.mode === 'add' ? 'Thêm Học Sinh' : 'Lưu Thay Đổi'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ====================================================
          MODAL 3: GHI NHẬN NGÀY ĐÓNG TIỀN CHI TIẾT
      ==================================================== */}
      {paymentModal.open && paymentModal.student && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold text-slate-900">Ghi Nhận Đóng Học Phí — {paymentModal.student.name}</h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Tháng {selectedMonth} • Số tiền: <strong className="text-emerald-600">{formatVND(paymentModal.student.tuitionFee)}</strong>
            </p>
            <form onSubmit={handleSaveCustomPayment} className="space-y-4 mt-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Ngày & Giờ đóng tiền</label>
                <input
                  type="datetime-local"
                  required
                  value={paymentModal.paidDate}
                  onChange={(e) => setPaymentModal({ ...paymentModal, paidDate: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Phương thức thanh toán</label>
                <select
                  value={paymentModal.method}
                  onChange={(e) => setPaymentModal({ ...paymentModal, method: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm bg-white"
                >
                  <option value="Chuyển khoản">Chuyển khoản ngân hàng</option>
                  <option value="Tiền mặt">Tiền mặt tại quầy</option>
                  <option value="Ví điện tử">Ví MoMo / ZaloPay</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Ghi chú biên thu</label>
                <input
                  type="text"
                  placeholder="VD: Mẹ bé chuyển khoản..."
                  value={paymentModal.note}
                  onChange={(e) => setPaymentModal({ ...paymentModal, note: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm"
                />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setPaymentModal({ ...paymentModal, open: false })}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100"
                >
                  Đóng
                </button>
                <button type="submit" className="px-4 py-2 rounded-xl text-xs font-semibold bg-emerald-600 text-white hover:bg-emerald-700">
                  Xác Nhận Đã Thu
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ====================================================
          MODAL 4: XEM TRƯỚC IMPORT EXCEL
      ==================================================== */}
      {excelModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900">Xem trước dữ liệu nhập từ Excel</h3>
                <p className="text-xs text-slate-500">
                  File: {excelModal.fileName} • Phát hiện {excelModal.previewRows.length} học sinh hợp lệ
                </p>
              </div>
              <button onClick={handleDownloadTemplate} className="text-xs font-semibold text-indigo-600 hover:underline">
                Tải file Excel mẫu
              </button>
            </div>
            <div className="max-h-64 overflow-y-auto my-4 border border-slate-200 rounded-xl">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 border-b border-slate-200 font-semibold text-slate-600">
                  <tr>
                    <th className="p-2.5">Mã HS</th>
                    <th className="p-2.5">Họ Tên</th>
                    <th className="p-2.5">Khóa học</th>
                    <th className="p-2.5">Học phí</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {excelModal.previewRows.map((r) => (
                    <tr key={r.id}>
                      <td className="p-2.5 font-mono font-bold">{r.id}</td>
                      <td className="p-2.5 font-medium">{r.name}</td>
                      <td className="p-2.5">{r.className}</td>
                      <td className="p-2.5 font-bold text-indigo-600">{formatVND(r.tuitionFee)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setExcelModal({ open: false, previewRows: [], fileName: '' })}
                className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100"
              >
                Hủy
              </button>
              <button
                onClick={confirmExcelImport}
                className="px-4 py-2 rounded-xl text-xs font-semibold bg-indigo-600 text-white hover:bg-indigo-700"
              >
                Xác nhận lưu {excelModal.previewRows.length} học sinh
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}