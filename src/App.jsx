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
  Upload,
  Cloud,
  CloudOff,
  Settings,
  X,
  Package,
  ShoppingCart,
  DollarSign,
  Store,
  FileSpreadsheet,
  UserX,
  UserCheck,
  TrendingUp,
  BookOpen,
  Menu,
  Filter,
  FileText,
  AlertCircle,
  BookPlus,
  UserPlus2,
  ClipboardCheck,
  Check,
  RotateCcw,
  PhoneCall,
  Eye,
  Award
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

const formatVND = (amount) => {
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(Number(amount) || 0);
};

const formatDateOnly = (dateStr) => {
  if (!dateStr) return '—';
  const cleanDate = String(dateStr).slice(0, 10);
  const parts = cleanDate.split('-');
  if (parts.length === 3) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return dateStr;
  return date.toLocaleDateString('vi-VN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
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

const validatePhone10Digits = (phone) => {
  const cleanPhone = String(phone || '').trim().replace(/[^0-9]/g, '');
  const regex = /^0\d{9}$/;
  return regex.test(cleanPhone);
};

const BELT_RANKS = [
  { name: 'Đai Trắng', bg: 'bg-slate-100', text: 'text-slate-700', border: 'border-slate-300' },
  { name: 'Đai Vàng', bg: 'bg-yellow-50', text: 'text-yellow-700', border: 'border-yellow-300' },
  { name: 'Đai Cam', bg: 'bg-orange-50', text: 'text-orange-700', border: 'border-orange-300' },
  { name: 'Đai Xanh Lá', bg: 'bg-emerald-50', text: 'text-emerald-700', border: 'border-emerald-300' },
  { name: 'Đai Xanh Dương', bg: 'bg-blue-50', text: 'text-blue-700', border: 'border-blue-300' },
  { name: 'Đai Đỏ', bg: 'bg-rose-50', text: 'text-rose-700', border: 'border-rose-300' },
  { name: 'Đai Tím', bg: 'bg-purple-50', text: 'text-purple-700', border: 'border-purple-300' },
  { name: 'Đai Nâu Cấp 2', bg: 'bg-amber-100', text: 'text-amber-900', border: 'border-amber-400' },
  { name: 'Đai Nâu Cấp 1', bg: 'bg-amber-900', text: 'text-white', border: 'border-amber-300' },
  { name: 'Đai Đen', bg: 'bg-slate-900', text: 'text-white', border: 'border-slate-900' }
];

const getNextBelt = (currentBelt) => {
  const idx = BELT_RANKS.findIndex((b) => b.name === (currentBelt || 'Đai Trắng'));
  if (idx === -1 || idx >= BELT_RANKS.length - 1) return BELT_RANKS[BELT_RANKS.length - 1].name;
  return BELT_RANKS[idx + 1].name;
};

const getBeltStyle = (beltName) => {
  return BELT_RANKS.find((b) => b.name === beltName) || BELT_RANKS[0];
};

export default function App() {
  const [courses, setCourses] = useState([]);
  const [students, setStudents] = useState([]);
  const [payments, setPayments] = useState({});
  const [products, setProducts] = useState([]);
  const [shopOrders, setShopOrders] = useState([]);
  const [attendance, setAttendance] = useState({});

  const [firebaseStatus, setFirebaseStatus] = useState(db ? 'connecting' : 'offline');
  const [selectedMonth, setSelectedMonth] = useState('2026-10');
  const [attendanceDate, setAttendanceDate] = useState(() => new Date().toISOString().slice(0, 10));

  const [activeTab, setActiveTab] = useState('tuition');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  const [searchQuery, setSearchQuery] = useState('');
  const [classFilter, setClassFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');

  const [selectedAttendanceIds, setSelectedAttendanceIds] = useState([]);
  const [showLowAttendanceOnly, setShowLowAttendanceOnly] = useState(false);

  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);

  const [showNotifPanel, setShowNotifPanel] = useState(false);
  const [paymentModal, setPaymentModal] = useState({ open: false, student: null, paidDate: '', method: 'Chuyển khoản', note: '' });
  const [studentModal, setStudentModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '', status: 'active' } });
  const [courseModal, setCourseModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' } });

  const [productModal, setProductModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', costPrice: '', price: '', stock: '', category: 'Đồng phục' } });
  const [sellModal, setSellModal] = useState({ open: false, productId: '', quantity: 1, buyerName: '', note: '' });

  const [excelChoiceModal, setExcelChoiceModal] = useState(false);
  const [excelPreviewModal, setExcelPreviewModal] = useState({ open: false, previewRows: [], fileName: '', missingCourses: [] });
  const fileInputRef = useRef(null);

  const [studentDetailModal, setStudentDetailModal] = useState({ open: false, student: null });

  const [toast, setToast] = useState(null);

  const [beltFilter, setBeltFilter] = useState('ALL');
  const [selectedExamIds, setSelectedExamIds] = useState([]);

  const showToast = (message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3500);
  };

  useEffect(() => {
    if (!db) {
      setFirebaseStatus('offline');
      return;
    }

    const unsubCourses = onSnapshot(collection(db, 'courses'), (snap) => {
      setFirebaseStatus('connected');
      setCourses(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    }, () => setFirebaseStatus('error'));

    const unsubStudents = onSnapshot(collection(db, 'students'), (snap) => {
      const list = snap.docs.map((d) => ({ id: d.id, status: 'active', ...d.data() }));
      list.sort((a, b) => a.id.localeCompare(b.id));
      setStudents(list);
    });

    const unsubPayments = onSnapshot(collection(db, 'payments'), (snap) => {
      const payMap = {};
      snap.docs.forEach((d) => { payMap[d.id] = d.data(); });
      setPayments(payMap);
    });

    const unsubProducts = onSnapshot(collection(db, 'products'), (snap) => {
      setProducts(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    const unsubShopOrders = onSnapshot(collection(db, 'shopOrders'), (snap) => {
      const list = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      list.sort((a, b) => new Date(b.date) - new Date(a.date));
      setShopOrders(list);
    });

    const unsubAttendance = onSnapshot(collection(db, 'attendance'), (snap) => {
      const attMap = {};
      snap.docs.forEach((d) => { attMap[d.id] = d.data(); });
      setAttendance(attMap);
    });

    return () => {
      unsubCourses();
      unsubStudents();
      unsubPayments();
      unsubProducts();
      unsubShopOrders();
      unsubAttendance();
    };
  }, []);

  useEffect(() => {
    setCurrentPage(1);
    setSelectedAttendanceIds([]);
  }, [activeTab, searchQuery, statusFilter, classFilter, rowsPerPage, attendanceDate, showLowAttendanceOnly]);

  useEffect(() => {
    if (!attendanceDate.startsWith(selectedMonth)) {
      setAttendanceDate(`${selectedMonth}-01`);
    }
  }, [selectedMonth]);

  const currentMonthPayments = payments[selectedMonth] || {};

  const activeStudentsList = useMemo(() => students.filter((s) => s.status !== 'inactive'), [students]);
  const inactiveStudentsList = useMemo(() => students.filter((s) => s.status === 'inactive'), [students]);

  const lowStockProducts = useMemo(() => {
    return products.filter((p) => Number(p.stock) <= 1);
  }, [products]);

  const currentMonthAttendance = attendance[selectedMonth] || {};

  // Kiểm tra xem học sinh có đi học buổi nào trong tháng đang chọn không
  const hasAttendedInSelectedMonth = (studentId) => {
    return Object.values(currentMonthAttendance).some(
      (dayMap) => dayMap && dayMap[studentId] === true
    );
  };

  // Kiểm tra học sinh có thuộc danh sách của tháng đang chọn (selectedMonth) hay không
  const isStudentVisibleInMonth = (student) => {
    const paymentInfo = currentMonthPayments[student.id];
    const hasPaidInThisMonth = Boolean(paymentInfo && paymentInfo.paid);
    const hasAttended = hasAttendedInSelectedMonth(student.id);

    // Nếu trong tháng đó có đóng học phí hoặc có điểm danh đi học -> Luôn hiển thị
    if (hasPaidInThisMonth || hasAttended) return true;

    // Nếu học sinh đã nghỉ: Các tháng trước tháng bắt đầu nghỉ (selectedMonth < inactiveFromMonth) vẫn hiển thị
    if (student.status === 'inactive') {
      if (student.inactiveFromMonth && selectedMonth < student.inactiveFromMonth) {
        return true;
      }
      return false;
    }

    return true;
  };

  const enrichedStudents = useMemo(() => {
    return students
      .filter((student) => isStudentVisibleInMonth(student))
      .map((student) => {
        const paymentInfo = currentMonthPayments[student.id] || { paid: false, paidDate: null, method: 'Chuyển khoản', note: '' };
        const effectiveFee = paymentInfo.paid && paymentInfo.amount !== undefined
          ? Number(paymentInfo.amount)
          : Number(student.tuitionFee || 0);

        return {
          ...student,
          tuitionFee: effectiveFee,
          paid: Boolean(paymentInfo.paid),
          paidDate: paymentInfo.paidDate || null,
          method: paymentInfo.method || 'Chuyển khoản',
          note: paymentInfo.note || ''
        };
      });
  }, [students, currentMonthPayments, currentMonthAttendance, selectedMonth]);

  const unpaidStudents = useMemo(() => enrichedStudents.filter((s) => !s.paid), [enrichedStudents]);
  const paidStudents = useMemo(() => enrichedStudents.filter((s) => s.paid), [enrichedStudents]);

  const totalNotificationsCount = unpaidStudents.length + lowStockProducts.length;

  const stats = useMemo(() => {
    let collectedRevenue = 0;
    let cashRevenue = 0;
    let transferRevenue = 0;

    Object.values(currentMonthPayments).forEach((p) => {
      if (p && p.paid) {
        const amt = Number(p.amount || 0);
        collectedRevenue += amt;
        if (p.method === 'Tiền mặt') {
          cashRevenue += amt;
        } else {
          transferRevenue += amt;
        }
      }
    });

    const pendingRevenue = unpaidStudents.reduce((sum, s) => sum + Number(s.tuitionFee || 0), 0);
    const expectedRevenue = collectedRevenue + pendingRevenue;
    const completionRate = enrichedStudents.length > 0 ? Math.round((paidStudents.length / enrichedStudents.length) * 100) : 0;

    return { expectedRevenue, collectedRevenue, cashRevenue, transferRevenue, pendingRevenue, completionRate };
  }, [currentMonthPayments, unpaidStudents, enrichedStudents, paidStudents]);

  const shopStats = useMemo(() => {
    const totalShopRevenue = shopOrders.reduce((sum, o) => sum + Number(o.totalAmount || 0), 0);
    const totalItemsSold = shopOrders.reduce((sum, o) => sum + Number(o.quantity || 0), 0);
    const totalShopProfit = shopOrders.reduce((sum, o) => {
      if (o.profit !== undefined) return sum + Number(o.profit);
      const prod = (products || []).find((p) => p.id === o.productId);
      const cost = o.costPrice !== undefined ? Number(o.costPrice) : Number(prod?.costPrice || 0);
      const qty = Number(o.quantity || 0);
      const rev = Number(o.totalAmount || 0);
      return sum + (rev - cost * qty);
    }, 0);
    return { totalShopRevenue, totalItemsSold, totalShopProfit };
  }, [shopOrders, products]);

  const displayStudentsSource = enrichedStudents;

  const filteredStudents = useMemo(() => {
    const list = displayStudentsSource.filter((s) => {
      const matchSearch = s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.id.toLowerCase().includes(searchQuery.toLowerCase()) || (s.phone && s.phone.includes(searchQuery));
      const matchClass = classFilter === 'ALL' || s.className === classFilter;
      const matchStatus = statusFilter === 'ALL' || (statusFilter === 'PAID' && s.paid) || (statusFilter === 'UNPAID' && !s.paid);
      return matchSearch && matchClass && matchStatus;
    });

    const collator = new Intl.Collator('vi', { sensitivity: 'base' });

    return list.sort((a, b) => {
      const partsA = a.name.trim().split(/\s+/);
      const partsB = b.name.trim().split(/\s+/);
      const givenA = partsA[partsA.length - 1];
      const givenB = partsB[partsB.length - 1];
      const byGiven = collator.compare(givenA, givenB);
      if (byGiven !== 0) return byGiven;
      return collator.compare(a.name, b.name);
    });
  }, [displayStudentsSource, searchQuery, classFilter, statusFilter]);

  const totalStudentPages = Math.ceil(filteredStudents.length / rowsPerPage) || 1;
  const paginatedStudents = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredStudents.slice(start, start + rowsPerPage);
  }, [filteredStudents, currentPage, rowsPerPage]);

  const totalShopPages = Math.ceil(shopOrders.length / rowsPerPage) || 1;
  const paginatedShopOrders = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return shopOrders.slice(start, start + rowsPerPage);
  }, [shopOrders, currentPage, rowsPerPage]);

  const totalProductPages = Math.ceil(products.length / rowsPerPage) || 1;
  const paginatedProducts = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return products.slice(start, start + rowsPerPage);
  }, [products, currentPage, rowsPerPage]);

  // ==========================================
  // ĐIỂM DANH & SẮP XẾP LUÔN THEO ALPHABET TÊN
  // ==========================================
  const currentDayAttendance = currentMonthAttendance[attendanceDate] || {};

  const monthlyAttendanceCounts = useMemo(() => {
    const counts = {};
    Object.values(currentMonthAttendance).forEach((dayMap) => {
      if (dayMap && typeof dayMap === 'object') {
        Object.entries(dayMap).forEach(([studentId, isPresent]) => {
          if (isPresent === true) {
            counts[studentId] = (counts[studentId] || 0) + 1;
          }
        });
      }
    });
    return counts;
  }, [currentMonthAttendance]);

  const getStudentAttendanceDetail = (studentId) => {
    const presentDates = [];
    Object.entries(currentMonthAttendance).forEach(([dateStr, dayMap]) => {
      if (dayMap && dayMap[studentId] === true) {
        presentDates.push(dateStr);
      }
    });
    presentDates.sort();
    return presentDates;
  };

  const lowAttendanceStudents = useMemo(() => {
    return enrichedStudents.filter((s) => (monthlyAttendanceCounts[s.id] || 0) < 3);
  }, [enrichedStudents, monthlyAttendanceCounts]);

  const collator = useMemo(() => new Intl.Collator('vi', { sensitivity: 'variant' }), []);

  const filteredAttendanceStudents = useMemo(() => {
    const list = enrichedStudents.filter((s) => {
      const matchSearch = s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.id.toLowerCase().includes(searchQuery.toLowerCase()) || (s.phone && s.phone.includes(searchQuery));
      const matchClass = classFilter === 'ALL' || s.className === classFilter;
      const matchLow = !showLowAttendanceOnly || (monthlyAttendanceCounts[s.id] || 0) < 3;
      return matchSearch && matchClass && matchLow;
    });

    const baseCollator = new Intl.Collator('vi', { sensitivity: 'base' });

    return list.sort((a, b) => {
      const partsA = a.name.trim().split(/\s+/);
      const partsB = b.name.trim().split(/\s+/);
      const givenA = partsA[partsA.length - 1];
      const givenB = partsB[partsB.length - 1];
      const byGiven = baseCollator.compare(givenA, givenB);
      if (byGiven !== 0) return byGiven;
      return baseCollator.compare(a.name, b.name);
    });
  }, [enrichedStudents, searchQuery, classFilter, showLowAttendanceOnly, monthlyAttendanceCounts]);

  const totalAttendancePages = Math.ceil(filteredAttendanceStudents.length / rowsPerPage) || 1;
  const paginatedAttendanceStudents = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredAttendanceStudents.slice(start, start + rowsPerPage);
  }, [filteredAttendanceStudents, currentPage, rowsPerPage]);

  const presentTodayCount = useMemo(() => {
    return filteredAttendanceStudents.filter(s => currentDayAttendance[s.id] === true).length;
  }, [filteredAttendanceStudents, currentDayAttendance]);

  const handleToggleAttendance = async (studentId) => {
    const monthKey = attendanceDate.slice(0, 7);
    const prevMonthObj = attendance[monthKey] || {};
    const prevDayObj = prevMonthObj[attendanceDate] || {};
    const isCurrentlyPresent = prevDayObj[studentId] === true;
    const newStatus = !isCurrentlyPresent;

    const updatedMonthObj = {
      ...prevMonthObj,
      [attendanceDate]: {
        ...prevDayObj,
        [studentId]: newStatus
      }
    };

    setAttendance((prev) => ({ ...prev, [monthKey]: updatedMonthObj }));
    if (db) {
      await setDoc(doc(db, 'attendance', monthKey), updatedMonthObj).catch((err) => showToast(err.message, 'error'));
    }
  };

  const handleSelectStudentCheckbox = (studentId) => {
    setSelectedAttendanceIds((prev) =>
      prev.includes(studentId) ? prev.filter((id) => id !== studentId) : [...prev, studentId]
    );
  };

  const handleSelectAllOnPage = (checked) => {
    const pageIds = paginatedAttendanceStudents.map((s) => s.id);
    if (checked) {
      setSelectedAttendanceIds((prev) => Array.from(new Set([...prev, ...pageIds])));
    } else {
      setSelectedAttendanceIds((prev) => prev.filter((id) => !pageIds.includes(id)));
    }
  };

  const handleBulkMarkPresent = async () => {
    if (selectedAttendanceIds.length === 0) {
      alert('Vui lòng tích chọn ít nhất 1 học sinh trong danh sách để điểm danh hàng loạt!');
      return;
    }

    const monthKey = attendanceDate.slice(0, 7);
    const prevMonthObj = attendance[monthKey] || {};
    const prevDayObj = { ...(prevMonthObj[attendanceDate] || {}) };

    selectedAttendanceIds.forEach((id) => {
      prevDayObj[id] = true;
    });

    const updatedMonthObj = {
      ...prevMonthObj,
      [attendanceDate]: prevDayObj
    };

    setAttendance((prev) => ({ ...prev, [monthKey]: updatedMonthObj }));
    if (db) {
      await setDoc(doc(db, 'attendance', monthKey), updatedMonthObj).catch((err) => showToast(err.message, 'error'));
    }
    showToast(`Đã điểm danh hàng loạt cho ${selectedAttendanceIds.length} học sinh!`);
    setSelectedAttendanceIds([]);
  };

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

  const openPaymentModal = (student) => {
    const todayDateOnly = new Date().toISOString().slice(0, 10);
    const existingDateOnly = student.paidDate ? String(student.paidDate).slice(0, 10) : todayDateOnly;
    const validMethod = student.method === 'Tiền mặt' ? 'Tiền mặt' : 'Chuyển khoản';
    setPaymentModal({ open: true, student, paidDate: existingDateOnly, method: validMethod, note: student.note || '' });
  };

  const handleSavePaymentModal = async (e) => {
    e.preventDefault();
    const { student, paidDate, method, note } = paymentModal;
    if (!student) return;

    const updatedMonth = {
      ...(payments[selectedMonth] || {}),
      [student.id]: {
        paid: true,
        paidDate: paidDate || new Date().toISOString().slice(0, 10),
        method,
        note,
        amount: Number(student.tuitionFee),
        studentName: student.name,
        className: student.className
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

  const handleDeactivateStudent = async (student) => {
    if (!window.confirm(`Cho học sinh ${student.name} nghỉ học từ tháng ${selectedMonth}? (Các tháng trước tháng ${selectedMonth} vẫn giữ nguyên danh sách và điểm danh)`)) return;
    const updated = { ...student, status: 'inactive', inactiveFromMonth: selectedMonth };
    setStudents((prev) => prev.map((s) => s.id === student.id ? updated : s));
    if (db) await setDoc(doc(db, 'students', student.id), updated).catch(console.error);
    showToast(`Đã ghi nhận ${student.name} nghỉ học từ tháng ${selectedMonth}.`, 'warning');
  };

  const handleRestoreStudent = async (student) => {
    const updated = { ...student, status: 'active', inactiveFromMonth: null };
    setStudents((prev) => prev.map((s) => s.id === student.id ? updated : s));
    if (db) await setDoc(doc(db, 'students', student.id), updated).catch(console.error);
    showToast(`Đã cho học sinh ${student.name} quay lại lớp học!`);
  };

  const handleDeleteStudentPermanent = async (studentId) => {
    if (!window.confirm('CẢNH BÁO: Thao tác này sẽ xóa vĩnh viễn hồ sơ và không thể khôi phục.')) return;
    setStudents((prev) => prev.filter((s) => s.id !== studentId));
    if (db) await deleteDoc(doc(db, 'students', studentId)).catch(console.error);
    showToast('Đã xóa vĩnh viễn học sinh.', 'warning');
  };

  const handleDownloadTemplate = () => {
    const templateData = [
      { 'Họ và Tên': 'Nguyễn Văn A', 'Khóa học': 'Karate Cơ Bản', 'Phụ huynh': 'Nguyễn Văn B', 'Số điện thoại': '0901234567' },
      { 'Họ và Tên': 'Trần Thị C', 'Khóa học': 'Karate Nâng Cao', 'Phụ huynh': 'Trần Văn D', 'Số điện thoại': '0987654321' }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(templateData), 'MauNhapHocSinh');
    XLSX.writeFile(wb, 'FileMau_NhapDanhSachHocSinh.xlsx');
    showToast('Đã tải xuống file mẫu Excel!');
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        const missingCoursesSet = new Set();

        let maxSeq = students.reduce((max, s) => {
          const match = String(s.id).match(/HS(\d+)/i);
          if (match) {
            const num = parseInt(match[1], 10);
            return num > max ? num : max;
          }
          return max;
        }, 0);

        const mappedRows = json.map((row) => {
          const name = String(row['Họ và Tên'] || row['Ho va Ten'] || row['Tên'] || '').trim();
          const className = String(row['Khóa học'] || row['Lớp'] || '').trim();
          const parentName = String(row['Phụ huynh'] || row['Phu huynh'] || '').trim();
          const rawPhone = String(row['Số điện thoại'] || row['SĐT'] || row['Phone'] || '').trim();

          const matchedCourse = courses.find(c => c.name.toLowerCase().trim() === className.toLowerCase());
          if (className && !matchedCourse) {
            missingCoursesSet.add(className);
          }

          maxSeq++;
          const generatedId = `HS${String(maxSeq).padStart(3, '0')}`;

          return {
            id: generatedId,
            name,
            className,
            parentName,
            phone: rawPhone,
            tuitionFee: matchedCourse ? Number(matchedCourse.fee) : 0,
            status: 'active',
            isValidPhone: validatePhone10Digits(rawPhone),
            isValidCourse: Boolean(matchedCourse)
          };
        }).filter(r => r.name);

        if (mappedRows.length === 0) {
          alert('Không tìm thấy dữ liệu hợp lệ trong file Excel.');
          return;
        }

        setExcelChoiceModal(false);
        setExcelPreviewModal({
          open: true,
          previewRows: mappedRows,
          fileName: file.name,
          missingCourses: Array.from(missingCoursesSet)
        });
      } catch (err) {
        alert('Lỗi đọc file Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = null;
  };

  const handleConfirmImport = async () => {
    const { previewRows, missingCourses } = excelPreviewModal;
    if (previewRows.length === 0) return;

    if (missingCourses.length > 0) {
      alert(`Không thể import! Có khóa học chưa tồn tại trên hệ thống:\n- ${missingCourses.join('\n- ')}`);
      return;
    }

    const invalidPhoneRow = previewRows.find(r => !r.isValidPhone);
    if (invalidPhoneRow) {
      alert(`Lỗi: Học sinh "${invalidPhoneRow.name}" có số điện thoại "${invalidPhoneRow.phone}" không hợp lệ!`);
      return;
    }

    const batch = db ? writeBatch(db) : null;

    previewRows.forEach((st) => {
      setStudents((prev) => [...prev, st]);
      if (batch && db) {
        batch.set(doc(db, 'students', st.id), st);
      }
    });

    if (batch) {
      await batch.commit().catch(console.error);
    }

    setExcelPreviewModal({ open: false, previewRows: [], fileName: '', missingCourses: [] });
    showToast(`Đã import thành công ${previewRows.length} học sinh!`);
  };

  const handleSaveProduct = async (e) => {
    e.preventDefault();
    const { mode, data } = productModal;
    if (!data.name.trim() || !data.price) return;

    const prodId = data.id || `PROD_${Date.now()}`;
    const prodObj = {
      id: prodId,
      name: data.name.trim(),
      costPrice: Number(data.costPrice) || 0,
      price: Number(data.price) || 0,
      stock: Number(data.stock) || 0,
      category: data.category || 'Đồng phục'
    };

    setProducts((prev) => mode === 'add' ? [...prev, prodObj] : prev.map((p) => p.id === prodId ? prodObj : p));
    if (db) {
      await setDoc(doc(db, 'products', prodId), prodObj).catch((err) => showToast(err.message, 'error'));
    }
    setProductModal({ open: false, mode: 'add', data: { id: '', name: '', costPrice: '', price: '', stock: '', category: 'Đồng phục' } });
    showToast('Đã lưu sản phẩm vào kho!');
  };

  const handleDeleteProduct = async (prodId) => {
    if (!window.confirm('Bạn có chắc muốn xóa sản phẩm này khỏi kho?')) return;
    setProducts((prev) => prev.filter((p) => p.id !== prodId));
    if (db) await deleteDoc(doc(db, 'products', prodId)).catch(console.error);
    showToast('Đã xóa sản phẩm.', 'warning');
  };

  const handleExecuteSale = async (e) => {
    e.preventDefault();
    const { productId, quantity, buyerName, note } = sellModal;
    const product = products.find((p) => p.id === productId);
    if (!product) return;

    const qty = Number(quantity) || 1;
    if (product.stock < qty) {
      alert(`Số lượng trong kho không đủ! Chỉ còn lại ${product.stock} sản phẩm.`);
      return;
    }

    const newStock = product.stock - qty;
    const costPrice = Number(product.costPrice || 0);
    const sellPrice = Number(product.price || 0);
    const totalAmount = sellPrice * qty;
    const profit = (sellPrice - costPrice) * qty;

    const orderId = `ORDER_${Date.now()}`;
    const orderObj = {
      id: orderId,
      productId: product.id,
      productName: product.name,
      costPrice,
      price: sellPrice,
      quantity: qty,
      totalAmount,
      profit,
      buyerName: buyerName.trim() || 'Khách lẻ / Võ sinh',
      note: note.trim(),
      date: new Date().toISOString()
    };

    const updatedProd = { ...product, stock: newStock };
    setProducts((prev) => prev.map((p) => p.id === product.id ? updatedProd : p));
    setShopOrders((prev) => [orderObj, ...prev]);

    if (db) {
      const batch = writeBatch(db);
      batch.set(doc(db, 'products', product.id), updatedProd);
      batch.set(doc(db, 'shopOrders', orderId), orderObj);
      await batch.commit().catch((err) => showToast(err.message, 'error'));
    }

    setSellModal({ open: false, productId: '', quantity: 1, buyerName: '', note: '' });
    if (newStock <= 2) {
      showToast(`Bán thành công! Cảnh báo: "${product.name}" trong kho chỉ còn ${newStock} sản phẩm!`, 'warning');
    } else {
      showToast(`Bán thành công ${qty} ${product.name}! Thu về ${formatVND(totalAmount)}`);
    }
  };

  const handleExportTuitionExcel = () => {
    const studentRows = enrichedStudents.map((s, idx) => ({
      'STT': idx + 1,
      'Mã HS': s.id,
      'Họ và Tên': s.name,
      'Khóa học': s.className,
      'Phụ huynh': s.parentName,
      'Số điện thoại': s.phone,
      'Học phí (VNĐ)': Number(s.tuitionFee),
      [`Trạng thái (${selectedMonth})`]: s.paid ? 'Đã đóng' : 'Chưa đóng',
      'Ngày đóng': s.paidDate ? formatDateOnly(s.paidDate) : '',
      'Tình trạng theo học': s.status === 'inactive' ? 'Đã nghỉ' : 'Đang học'
    }));

    // Thêm dòng Tổng thực thu ở cuối bảng Excel (không tách tiền mặt/chuyển khoản)
    studentRows.push({
      'STT': '',
      'Mã HS': '',
      'Họ và Tên': 'TỔNG DOANH THU ĐÃ THU',
      'Khóa học': '',
      'Phụ huynh': '',
      'Số điện thoại': '',
      'Học phí (VNĐ)': stats.collectedRevenue,
      [`Trạng thái (${selectedMonth})`]: '',
      'Ngày đóng': '',
      'Tình trạng theo học': ''
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(studentRows), `HocPhi_${selectedMonth}`);
    XLSX.writeFile(wb, `BaoCaoHocPhi_${selectedMonth}.xlsx`);
    showToast('Xuất báo cáo học phí thành công!');
  };

  const handleExportStudentListExcel = () => {
    const sortedStudents = [...activeStudentsList].sort((a, b) => {
      const partsA = a.name.trim().split(/\s+/);
      const partsB = b.name.trim().split(/\s+/);

      const givenA = partsA[partsA.length - 1];
      const givenB = partsB[partsB.length - 1];

      const byGiven = collator.compare(givenA, givenB);
      if (byGiven !== 0) return byGiven;

      return collator.compare(a.name, b.name);
    });

    const rows = sortedStudents.map((s, idx) => ({
      'STT': idx + 1,
      'Họ và Tên': s.name,
      'Lớp / Khóa học': s.className,
      'Số điện thoại phụ huynh': s.phone
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'DanhSachHocSinh');
    XLSX.writeFile(wb, 'DanhSachHocSinh.xlsx');
    showToast('Xuất danh sách học sinh thành công!');
  };

  const handleExportShopExcel = () => {
    const shopRows = shopOrders.map((o, idx) => ({
      'STT': idx + 1,
      'Mã Giao Dịch': o.id,
      'Thời Gian': formatDateTime(o.date),
      'Tên Sản Phẩm': o.productName,
      'Số Lượng': o.quantity,
      'Giá Bán (VNĐ)': Number(o.price),
      'Thành Tiền (VNĐ)': Number(o.totalAmount),
      'Người Mua': o.buyerName,
      'Ghi Chú': o.note || ''
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(shopRows), 'DoanhThuBanHang');
    XLSX.writeFile(wb, 'BaoCaoDoanhThuBanHang_Shop.xlsx');
    showToast('Xuất báo cáo doanh thu bán hàng thành công!');
  };

  // ==========================================
  // LOGIC THI LÊN ĐAI (BELT EXAM)
  // ==========================================
  const filteredExamStudents = useMemo(() => {
    const list = activeStudentsList.filter((s) => {
      const currentBelt = s.belt || 'Đai Trắng';
      const matchSearch = s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.id.toLowerCase().includes(searchQuery.toLowerCase());
      const matchClass = classFilter === 'ALL' || s.className === classFilter;
      const matchBelt = beltFilter === 'ALL' || currentBelt === beltFilter;
      return matchSearch && matchClass && matchBelt;
    });

    const baseCollator = new Intl.Collator('vi', { sensitivity: 'base' });
    return list.sort((a, b) => {
      const partsA = a.name.trim().split(/\s+/);
      const partsB = b.name.trim().split(/\s+/);
      const byGiven = baseCollator.compare(partsA[partsA.length - 1], partsB[partsB.length - 1]);
      return byGiven !== 0 ? byGiven : baseCollator.compare(a.name, b.name);
    });
  }, [activeStudentsList, searchQuery, classFilter, beltFilter]);

  // Đổi màu đai hiện tại trực tiếp trên bảng
  const handleQuickChangeBelt = async (student, newBelt) => {
    const updated = { ...student, belt: newBelt };
    setStudents((prev) => prev.map((s) => (s.id === student.id ? updated : s)));
    if (db) await setDoc(doc(db, 'students', student.id), updated).catch(console.error);
    showToast(`Đã cập nhật ${student.name} sang ${newBelt}!`);
  };

  // Chọn nhanh tất cả học sinh đủ điều kiện chuyên cần (>= 4 buổi trong tháng)
  const handleSelectEligibleStudents = () => {
    const eligibleIds = filteredExamStudents
      .filter((s) => (monthlyAttendanceCounts[s.id] || 0) >= 4)
      .map((s) => s.id);
    if (eligibleIds.length === 0) {
      alert('Chưa có học sinh nào trong danh sách lọc đạt từ 4 buổi học trở lên trong tháng này!');
      return;
    }
    setSelectedExamIds(eligibleIds);
    showToast(`Đã tích chọn ${eligibleIds.length} võ sinh đủ điều kiện chuyên cần!`);
  };

  // Thăng đai hàng loạt cho các võ sinh đã chọn sau khi thi đậu
  const handleBulkPromoteBelt = async () => {
    if (selectedExamIds.length === 0) {
      alert('Vui lòng tích chọn các võ sinh đã thi đậu để thăng đai!');
      return;
    }
    if (!window.confirm(`Xác nhận nâng lên 1 cấp đai mới cho ${selectedExamIds.length} võ sinh đã chọn?`)) return;

    const batch = db ? writeBatch(db) : null;
    const updatedList = students.map((s) => {
      if (selectedExamIds.includes(s.id)) {
        const nextBelt = getNextBelt(s.belt || 'Đai Trắng');
        const updatedStudent = { ...s, belt: nextBelt };
        if (batch && db) batch.set(doc(db, 'students', s.id), updatedStudent);
        return updatedStudent;
      }
      return s;
    });

    setStudents(updatedList);
    if (batch) await batch.commit().catch(console.error);
    showToast(`Đã thăng cấp đai thành công cho ${selectedExamIds.length} võ sinh!`);
    setSelectedExamIds([]);
  };

  // Xuất file Excel Danh sách Thi Lên Đai (Tự động chia Sheet theo từng màu đai)
  const handleExportBeltExamExcel = () => {
    const targetList = selectedExamIds.length > 0
      ? filteredExamStudents.filter((s) => selectedExamIds.includes(s.id))
      : filteredExamStudents;

    if (targetList.length === 0) {
      alert('Không có võ sinh nào để xuất danh sách!');
      return;
    }

    const wb = XLSX.utils.book_new();

    // 1. Sheet Tổng hợp tất cả thí sinh
    const allRows = targetList.map((s, idx) => ({
      'STT': idx + 1,
      'Mã HS': s.id,
      'Họ và Tên': s.name,
      'Lớp / Khóa học': s.className,
      'Đai Hiện Tại': s.belt || 'Đai Trắng',
      'Đai Dự Thi': getNextBelt(s.belt || 'Đai Trắng'),
      'SĐT Phụ Huynh': s.phone,
      'Điểm Kỹ Thuật': '',
      'Điểm Đối Kháng': '',
      'Kết Quả (Đạt / Chưa đạt)': ''
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allRows), 'TongHop_ThiLenDai');

    // 2. Tự động gom nhóm và tạo từng Sheet riêng cho mỗi Màu Đai Hiện Tại
    BELT_RANKS.forEach((beltObj) => {
      const group = targetList.filter((s) => (s.belt || 'Đai Trắng') === beltObj.name);
      if (group.length > 0) {
        const sheetRows = group.map((s, idx) => ({
          'STT': idx + 1,
          'Mã HS': s.id,
          'Họ và Tên': s.name,
          'Lớp': s.className,
          'Đai Hiện Tại': s.belt || 'Đai Trắng',
          'Thi Lên Đai': getNextBelt(s.belt || 'Đai Trắng'),
          'Điểm Kỹ Thuật': '',
          'Điểm Đối Kháng': '',
          'Kết Quả': '',
          'Ghi Chú': ''
        }));
        const safeSheetName = `ThiLen_${getNextBelt(beltObj.name)}`.replace(/\s+/g, '').slice(0, 31);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheetRows), safeSheetName);
      }
    });

    XLSX.writeFile(wb, `DanhSachThiLenDai_${selectedMonth}.xlsx`);
    showToast(`Đã xuất danh sách thi lên đai (${targetList.length} võ sinh)!`);
  };

  const allOnPageSelected = paginatedAttendanceStudents.length > 0 && paginatedAttendanceStudents.every((s) => selectedAttendanceIds.includes(s.id));

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 flex flex-col md:flex-row relative">
      {toast && (
        <div className="fixed bottom-4 right-4 left-4 sm:left-auto z-50 flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg text-white text-sm font-medium bg-slate-900">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast.message}</span>
        </div>
      )}

      <input type="file" ref={fileInputRef} onChange={handleFileChange} accept=".xlsx, .xls, .csv" className="hidden" />

      {/* ================= SIDEBAR (CỐ ĐỊNH TRÊN TABLET/DESKTOP, DRAWER TRÊN MOBILE) ================= */}
      {isMobileMenuOpen && (
        <div 
          className="fixed inset-0 bg-black/60 z-40 md:hidden backdrop-blur-xs transition-opacity"
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}

      <aside className={`fixed md:sticky md:top-0 md:h-screen inset-y-0 left-0 z-50 w-64 bg-black text-slate-300 flex flex-col justify-between border-r border-slate-800 shrink-0 transform transition-transform duration-300 ease-in-out ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div className="overflow-y-auto">
          <div className="p-4 sm:p-5 flex items-center justify-between border-b border-slate-800">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center text-white font-bold overflow-hidden shrink-0">
                <img src="/logo.jpg" alt="Logo" className="w-full h-full object-cover" />
              </div>
              <div>
                <h1 className="text-base font-bold text-red-500 tracking-wide">KIAIDO Manager</h1>
                <p className="text-[10px] text-slate-400">Hệ thống quản lý trung tâm</p>
              </div>
            </div>
            <button 
              onClick={() => setIsMobileMenuOpen(false)}
              className="md:hidden text-slate-400 hover:text-white p-1 rounded-lg"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          <nav className="p-3 space-y-1">
            <p className="px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-500">Quản Lý Trung Tâm</p>
            <button
              onClick={() => { setActiveTab('tuition'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'tuition' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <DollarSign className="w-4 h-4 shrink-0" /> Thu Học Phí
            </button>
            <button
              onClick={() => { setActiveTab('attendance'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'attendance' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <ClipboardCheck className="w-4 h-4 shrink-0" /> Điểm Danh Học Sinh
            </button>
            <button
              onClick={() => { setActiveTab('belt_exam'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'belt_exam' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <Award className="w-4 h-4 shrink-0" /> Thi Lên Đai
            </button>
            <button
              onClick={() => { setActiveTab('revenue'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'revenue' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <TrendingUp className="w-4 h-4 shrink-0" /> Báo Cáo Doanh Thu
            </button>
            <button
              onClick={() => { setActiveTab('students'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'students' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <Users className="w-4 h-4 shrink-0" /> Danh sách Học Sinh ({activeStudentsList.length})
            </button>
            <button
              onClick={() => { setActiveTab('inactive_students'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'inactive_students' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <UserX className="w-4 h-4 shrink-0" /> Học Sinh Đã Nghỉ ({inactiveStudentsList.length})
            </button>
            <button
              onClick={() => { setActiveTab('courses'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'courses' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <BookOpen className="w-4 h-4 shrink-0" /> Khóa Học & Lớp ({courses.length})
            </button>

            <p className="px-3 pt-4 pb-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-500">Kho Quần Áo & Dụng Cụ</p>
            <button
              onClick={() => { setActiveTab('inventory'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'inventory' ? 'bg-indigo-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <span className="flex items-center gap-3">
                <Package className="w-4 h-4 shrink-0" /> Quản Lý Kho ({products.length})
              </span>
              {lowStockProducts.length > 0 && (
                <span className="bg-rose-600 text-white text-[10px] font-bold px-1.5 py-0.5 rounded-full">
                  ! {lowStockProducts.length}
                </span>
              )}
            </button>
          </nav>
        </div>

        <div className="p-4 border-t border-slate-800 text-xs">
          {firebaseStatus === 'connected' ? (
            <span className="flex items-center gap-1.5 text-emerald-400 font-medium">
              <Cloud className="w-3.5 h-3.5" /> Đã kết nối Firebase
            </span>
          ) : (
            <span className="flex items-center gap-1.5 text-amber-400 font-medium">
              <CloudOff className="w-3.5 h-3.5" /> Chế độ Offline
            </span>
          )}
        </div>
      </aside>

      {/* ================= MAIN CONTENT AREA ================= */}
      <div className="flex-1 flex flex-col min-w-0">
        <header className="bg-white border-b border-slate-200 sticky top-0 z-20 px-3 sm:px-6 py-3 flex items-center justify-between gap-2 shadow-xs">
          <div className="flex items-center gap-2.5 min-w-0">
            <button
              onClick={() => setIsMobileMenuOpen(true)}
              className="md:hidden p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 transition shrink-0"
              title="Mở menu"
            >
              <Menu className="w-5 h-5" />
            </button>

            <span className="text-xs sm:text-sm font-bold text-slate-800 uppercase tracking-wide truncate">
              {activeTab === 'tuition' && 'Quản Lý Thu Học Phí'}
              {activeTab === 'attendance' && 'Điểm Danh & Chuyên Cần'}
              {activeTab === 'belt_exam' && 'Lập Danh Sách Thi Lên Đai'}
              {activeTab === 'revenue' && 'Báo Cáo Doanh Thu'}
              {activeTab === 'students' && 'Danh Sách Học Sinh'}
              {activeTab === 'inactive_students' && 'Học Sinh Đã Nghỉ'}
              {activeTab === 'courses' && 'Quản Lý Khóa Học'}
              {activeTab === 'inventory' && 'Quản Lý Kho Hàng'}
            </span>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <div className="flex items-center bg-slate-100 rounded-xl p-1 border border-slate-200">
              <button onClick={() => changeMonth(-1)} className="p-1 hover:bg-white rounded-lg transition"><ChevronLeft className="w-4 h-4" /></button>
              <div className="flex items-center gap-1 px-1.5 sm:px-2 text-xs font-semibold text-slate-800">
                <Calendar className="w-3.5 h-3.5 text-indigo-600 hidden sm:inline-block shrink-0" />
                <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="bg-transparent border-none focus:outline-none cursor-pointer text-[11px] sm:text-xs w-[105px] sm:w-auto" />
              </div>
              <button onClick={() => changeMonth(1)} className="p-1 hover:bg-white rounded-lg transition"><ChevronRight className="w-4 h-4" /></button>
            </div>

            <div className="relative">
              <button onClick={() => setShowNotifPanel(!showNotifPanel)} className={`relative p-2 rounded-xl border transition ${totalNotificationsCount > 0 ? 'bg-rose-50 border-rose-200 text-rose-600' : 'bg-slate-100 border-slate-200 text-slate-600'}`}>
                <Bell className="w-4 h-4" />
                {totalNotificationsCount > 0 && <span className="absolute -top-1 -right-1 bg-rose-600 text-white text-[10px] font-bold w-4 h-4 rounded-full flex items-center justify-center">{totalNotificationsCount}</span>}
              </button>
              {showNotifPanel && (
                <div className="absolute right-0 mt-2 w-80 sm:w-96 max-w-[calc(100vw-24px)] bg-white rounded-2xl shadow-xl border border-slate-200 z-50 overflow-hidden">
                  {lowStockProducts.length > 0 && (
                    <div className="border-b border-slate-200">
                      <div className="p-3 bg-amber-50 border-b border-amber-100 font-bold text-xs text-amber-900 flex items-center justify-between">
                        <span className="flex items-center gap-1.5">
                          <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" /> Cảnh báo kho sắp hết (≤ 2): {lowStockProducts.length} SP
                        </span>
                        <button
                          onClick={() => { setActiveTab('inventory'); setShowNotifPanel(false); }}
                          className="text-[11px] text-indigo-600 underline font-semibold shrink-0"
                        >
                          Mở Kho
                        </button>
                      </div>
                      <div className="max-h-40 overflow-y-auto divide-y divide-slate-100">
                        {lowStockProducts.map((p) => (
                          <div key={p.id} className="p-2.5 flex items-center justify-between text-xs bg-amber-50/20">
                            <div>
                              <p className="font-semibold text-slate-800">{p.name}</p>
                              <p className="text-[11px] text-rose-600 font-bold">Chỉ còn: {p.stock} sản phẩm</p>
                            </div>
                            <button
                              onClick={() => {
                                setShowNotifPanel(false);
                                setActiveTab('inventory');
                                setProductModal({ open: true, mode: 'edit', data: { costPrice: 0, ...p } });
                              }}
                              className="px-2.5 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-xl font-semibold"
                            >
                              Bổ sung
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="p-3 bg-rose-50 border-b border-rose-100 font-semibold text-xs text-rose-800">
                    Chưa đóng học phí ({selectedMonth}): {unpaidStudents.length} học sinh
                  </div>
                  <div className="max-h-56 overflow-y-auto divide-y divide-slate-100">
                    {unpaidStudents.length === 0 ? (
                      <div className="p-4 text-center text-xs text-slate-400">Tất cả học sinh đã hoàn thành học phí!</div>
                    ) : (
                      unpaidStudents.map((s) => (
                        <div key={s.id} className="p-2.5 flex items-center justify-between text-xs">
                          <span className="font-medium">{s.name}</span>
                          <button
                            onClick={() => {
                              setShowNotifPanel(false);
                              openPaymentModal(s);
                            }}
                            className="px-2.5 py-1 bg-emerald-600 text-white rounded-lg"
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
        </header>

        <main className="max-w-7xl mx-auto w-full px-3 sm:px-6 py-4 sm:py-6 space-y-4 sm:space-y-6">
          {lowStockProducts.length > 0 && (
            <div className="p-3.5 sm:p-4 rounded-2xl bg-rose-50 border border-rose-200 flex items-center justify-between flex-wrap gap-3 shadow-xs">
              <div className="flex items-center gap-2.5 text-xs sm:text-sm text-rose-800">
                <AlertCircle className="w-5 h-5 text-rose-600 shrink-0" />
                <div>
                  <span className="font-bold">Cảnh báo tồn kho thấp (≤ 1): </span>
                  <span>
                    Có <strong>{lowStockProducts.length} sản phẩm</strong> sắp hết hoặc đã hết hàng (
                    {lowStockProducts.map((p) => `${p.name}: còn ${p.stock}`).join(', ')}). Vui lòng kiểm tra và nhập bổ sung kịp thời!
                  </span>
                </div>
              </div>
              {activeTab !== 'inventory' && (
                <button
                  onClick={() => setActiveTab('inventory')}
                  className="px-3.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold shrink-0"
                >
                  Kiểm tra kho ngay
                </button>
              )}
            </div>
          )}

          {/* TOOLBAR TỐI ƯU GRID TRÊN MOBILE, FLEX TRÊN TABLET/DESKTOP */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
            <div className="grid grid-cols-2 sm:flex items-center gap-2">
              {activeTab === 'tuition' || activeTab === 'students' ? (
                <>
                  <button onClick={handleExportStudentListExcel} className="px-3 py-2.5 sm:px-4 sm:py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 shadow-xs" title="Xuất danh sách học sinh theo alphabet">
                    <Download className="w-4 h-4 text-emerald-600 shrink-0" /> <span className="truncate">Xuất DS Học Sinh</span>
                  </button>
                  <button onClick={() => setExcelChoiceModal(true)} className="px-3 py-2.5 sm:px-4 sm:py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 shadow-xs">
                    <Upload className="w-4 h-4 text-indigo-600 shrink-0" /> <span className="truncate">Nhập DS Học Sinh</span>
                  </button>
                </>
              ) : null}
            </div>

            <div className="grid grid-cols-3 sm:flex items-center gap-2">
              <button
                onClick={() => setSellModal({ open: true, productId: products[0]?.id || '', quantity: 1, buyerName: '', note: '' })}
                className="px-2.5 py-2.5 sm:px-3.5 sm:py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 shadow-xs"
              >
                <ShoppingCart className="w-3.5 h-3.5 shrink-0" /> <span className="truncate">Bán hàng</span>
              </button>
              <button onClick={() => setCourseModal({ open: true, mode: 'add', data: { id: `COURSE_${Date.now()}`, name: '', fee: '', description: '' } })} className="px-2.5 py-2.5 sm:px-3.5 sm:py-2 bg-amber-500 hover:bg-amber-600 text-white rounded-xl text-xs font-semibold flex items-center justify-center gap-1 shadow-xs">
                <BookPlus className="w-3.5 h-3.5 shrink-0" /> <span className="truncate">Khóa Học</span>
              </button>
              <button onClick={() => setStudentModal({ open: true, mode: 'add', data: { id: `HS${String(students.length + 1).padStart(3, '0')}`, name: '', className: courses[0]?.name || '', phone: '', parentName: '', tuitionFee: courses[0]?.fee || 1200000, status: 'active' } })} className="px-2.5 py-2.5 sm:px-3.5 sm:py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold flex items-center justify-center gap-1 shadow-xs">
                <UserPlus2 className="w-3.5 h-3.5 shrink-0" /> <span className="truncate">Học Sinh</span>
              </button>
            </div>
          </div>

          {/* ====================================================
              TAB: ĐIỂM DANH HỌC SINH (TỐI ƯU CARD MOBILE & TABLE DESKTOP)
          ==================================================== */}
          {activeTab === 'attendance' && (
            <div className="space-y-4 sm:space-y-5">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4">
                <div className="bg-white p-4 rounded-2xl border border-emerald-200 shadow-xs flex items-center justify-between">
                  <div>
                    <p className="text-xs text-slate-500 font-medium">Học sinh đang theo học</p>
                    <p className="text-2xl font-extrabold text-emerald-600 mt-1">{activeStudentsList.length} <span className="text-xs font-normal text-slate-500">học sinh</span></p>
                    <p className="text-[11px] text-slate-500 mt-1">Đi học ngày {formatDateOnly(attendanceDate)}: <strong>{presentTodayCount} em</strong></p>
                  </div>
                  <div className="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0">
                    <UserCheck className="w-5 h-5" />
                  </div>
                </div>

                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs flex items-center justify-between">
                  <div>
                    <p className="text-xs text-slate-500 font-medium">Học sinh đã nghỉ / Bảo lưu</p>
                    <p className="text-2xl font-extrabold text-slate-700 mt-1">{inactiveStudentsList.length} <span className="text-xs font-normal text-slate-500">học sinh</span></p>
                    <button
                      onClick={() => setActiveTab('inactive_students')}
                      className="text-[11px] text-indigo-600 font-semibold hover:underline mt-1 inline-block"
                    >
                      Xem danh sách nghỉ →
                    </button>
                  </div>
                  <div className="w-10 h-10 rounded-xl bg-slate-100 text-slate-600 flex items-center justify-center shrink-0">
                    <UserX className="w-5 h-5" />
                  </div>
                </div>

                <div className={`p-4 rounded-2xl border shadow-xs flex items-center justify-between transition ${showLowAttendanceOnly ? 'bg-amber-50 border-amber-400' : 'bg-white border-amber-200'}`}>
                  <div>
                    <p className="text-xs text-amber-800 font-bold">Đi học ít trong tháng (&lt; 3 buổi)</p>
                    <p className="text-2xl font-extrabold text-amber-600 mt-1">{lowAttendanceStudents.length} <span className="text-xs font-normal text-amber-700">học sinh</span></p>
                    <button
                      onClick={() => setShowLowAttendanceOnly(!showLowAttendanceOnly)}
                      className="mt-1 px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-bold inline-flex items-center gap-1"
                    >
                      <PhoneCall className="w-3 h-3" />
                      {showLowAttendanceOnly ? 'Đang lọc: Xem tất cả' : 'Lọc DS gọi hỏi thăm'}
                    </button>
                  </div>
                  <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                    <PhoneCall className="w-5 h-5" />
                  </div>
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                <div className="p-3.5 sm:p-4 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-sm sm:text-base text-slate-900 flex items-center gap-2">
                      <ClipboardCheck className="w-5 h-5 text-indigo-600 shrink-0" /> Điểm Danh Học Sinh (Tháng {selectedMonth})
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Tích chọn nhiều học sinh để <strong>Điểm danh hàng loạt</strong> hoặc bấm <strong>Điểm danh</strong> ở từng em.
                    </p>
                  </div>

                  <div className="flex items-center justify-between sm:justify-end gap-2 flex-wrap">
                    <div className="flex items-center gap-1.5 bg-white px-3 py-2 rounded-xl border border-slate-300 text-xs font-semibold">
                      <span>Ngày:</span>
                      <input
                        type="date"
                        value={attendanceDate}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val) {
                            setAttendanceDate(val);
                            setSelectedMonth(val.slice(0, 7));
                          }
                        }}
                        className="bg-transparent text-indigo-600 font-bold focus:outline-none cursor-pointer"
                      />
                    </div>

                    <button
                      onClick={handleBulkMarkPresent}
                      disabled={selectedAttendanceIds.length === 0}
                      className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-200 disabled:text-slate-400 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-xs"
                    >
                      <Check className="w-4 h-4 shrink-0" /> Điểm danh loạt ({selectedAttendanceIds.length})
                    </button>
                  </div>
                </div>

                <div className="p-3.5 border-b border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5 bg-white">
                  <div className="flex items-center gap-2 bg-slate-50 px-3.5 py-2 rounded-xl border border-slate-200 flex-1">
                    <Search className="w-4 h-4 text-slate-400 shrink-0" />
                    <input
                      type="text"
                      placeholder="Tìm nhanh tên học sinh, mã HS hoặc SĐT..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full bg-transparent border-none text-sm focus:outline-none"
                    />
                  </div>
                  <div className="flex items-center gap-2">
                    <label className="flex md:hidden items-center gap-1.5 text-xs font-semibold text-slate-700 bg-slate-100 px-3 py-2 rounded-xl border border-slate-200 cursor-pointer shrink-0">
                      <input
                        type="checkbox"
                        checked={allOnPageSelected}
                        onChange={(e) => handleSelectAllOnPage(e.target.checked)}
                        className="w-4 h-4 rounded text-indigo-600"
                      />
                      Chọn trang
                    </label>
                    <select
                      value={classFilter}
                      onChange={(e) => setClassFilter(e.target.value)}
                      className="flex-1 sm:flex-initial px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs sm:text-sm font-semibold text-slate-700"
                    >
                      <option value="ALL">Tất cả các lớp</option>
                      {courses.map((c) => (
                        <option key={c.id} value={c.name}>{c.name}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* MOBILE CARD VIEW CHO ĐIỂM DANH */}
                <div className="divide-y divide-slate-200 md:hidden">
                  {paginatedAttendanceStudents.length === 0 ? (
                    <div className="p-8 text-center text-sm text-slate-400">Không có học sinh nào phù hợp.</div>
                  ) : (
                    paginatedAttendanceStudents.map((s) => {
                      const isPresent = currentDayAttendance[s.id] === true;
                      const totalDaysInMonth = monthlyAttendanceCounts[s.id] || 0;
                      const isChecked = selectedAttendanceIds.includes(s.id);
                      const isLowAttendance = totalDaysInMonth < 3;

                      return (
                        <div key={s.id} className={`p-3.5 space-y-2.5 ${isPresent ? 'bg-emerald-50/40' : ''}`}>
                          <div className="flex items-start justify-between gap-2">
                            <div className="flex items-start gap-2.5">
                              <input
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => handleSelectStudentCheckbox(s.id)}
                                className="mt-1 w-4 h-4 rounded text-indigo-600 cursor-pointer shrink-0"
                              />
                              <div>
                                <div className="flex items-center gap-1.5 flex-wrap">
                                  <span className="font-mono text-[11px] font-bold px-1.5 py-0.5 bg-slate-100 text-slate-700 rounded">{s.id}</span>
                                  <span className="px-2 py-0.5 bg-indigo-50 text-indigo-700 rounded text-[11px] font-medium">{s.className}</span>
                                </div>
                                <button
                                  onClick={() => setStudentDetailModal({ open: true, student: s })}
                                  className="font-bold text-sm text-slate-900 hover:text-indigo-600 underline text-left inline-flex items-center gap-1 mt-1"
                                >
                                  {s.name} <Eye className="w-3.5 h-3.5 text-indigo-500 shrink-0" />
                                </button>
                                <p className="text-xs text-slate-500 mt-0.5">
                                  PH: {s.parentName || '—'} •{' '}
                                  {s.phone ? (
                                    <a href={`https://zalo.me/${s.phone.replace(/[^0-9]/g, '')}`} target="_blank" rel="noopener noreferrer" className="text-indigo-600 font-semibold underline">
                                      {s.phone}
                                    </a>
                                  ) : '—'}
                                </p>
                              </div>
                            </div>

                            <button
                              onClick={() => setStudentDetailModal({ open: true, student: s })}
                              className={`px-2 py-1 rounded-full text-[11px] font-bold border shrink-0 ${
                                isLowAttendance ? 'bg-amber-50 text-amber-700 border-amber-200' : 'bg-blue-50 text-blue-700 border-blue-200'
                              }`}
                            >
                              {totalDaysInMonth} buổi
                            </button>
                          </div>

                          <div className="flex items-center justify-end pt-1">
                            <button
                              onClick={() => handleToggleAttendance(s.id)}
                              className={`w-full py-2 rounded-xl text-xs font-bold transition shadow-xs ${
                                isPresent
                                  ? 'bg-emerald-600 text-white'
                                  : 'bg-emerald-500 text-white border border-emerald-200'
                              }`}
                            >
                              {isPresent ? '✓ Đã điểm danh' : 'Điểm danh'}
                            </button>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>

                {/* DESKTOP TABLE VIEW CHO ĐIỂM DANH */}
                <div className="hidden md:block overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                        <th className="p-3.5 w-10 text-center">
                          <input
                            type="checkbox"
                            checked={allOnPageSelected}
                            onChange={(e) => handleSelectAllOnPage(e.target.checked)}
                            className="w-4 h-4 rounded text-indigo-600 cursor-pointer"
                            title="Chọn tất cả trên trang"
                          />
                        </th>
                        <th className="p-3.5">Mã HS</th>
                        <th className="p-3.5">Họ Tên & SĐT Liên Hệ</th>
                        <th className="p-3.5">Lớp / Khóa Học</th>
                        <th className="p-3.5 text-center">Tổng buổi đi học ({selectedMonth})</th>
                        <th className="p-3.5 text-right">Trạng thái ngày {formatDateOnly(attendanceDate)}</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-sm">
                      {paginatedAttendanceStudents.length === 0 ? (
                        <tr><td colSpan="6" className="p-8 text-center text-slate-400">Không có học sinh nào phù hợp.</td></tr>
                      ) : (
                        paginatedAttendanceStudents.map((s) => {
                          const isPresent = currentDayAttendance[s.id] === true;
                          const totalDaysInMonth = monthlyAttendanceCounts[s.id] || 0;
                          const isChecked = selectedAttendanceIds.includes(s.id);
                          const isLowAttendance = totalDaysInMonth < 3;

                          return (
                            <tr key={s.id} className={isPresent ? 'bg-emerald-50/40 hover:bg-emerald-50/70' : 'hover:bg-slate-50'}>
                              <td className="p-3.5 text-center">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => handleSelectStudentCheckbox(s.id)}
                                  className="w-4 h-4 rounded text-indigo-600 cursor-pointer"
                                />
                              </td>
                              <td className="p-3.5 font-mono text-xs font-bold">{s.id}</td>
                              <td className="p-3.5">
                                <button
                                  onClick={() => setStudentDetailModal({ open: true, student: s })}
                                  className="font-semibold text-slate-900 hover:text-indigo-600 underline text-left inline-flex items-center gap-1"
                                >
                                  {s.name} <Eye className="w-3.5 h-3.5 text-indigo-500" />
                                </button>
                                <p className="text-xs text-slate-500">
                                  PH: {s.parentName || '—'} •{' '}
                                  {s.phone ? (
                                    <a
                                      href={`https://zalo.me/${s.phone.replace(/[^0-9]/g, '')}`}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="text-indigo-600 font-semibold underline"
                                    >
                                      {s.phone} (Gọi Zalo)
                                    </a>
                                  ) : '—'}
                                </p>
                              </td>
                              <td className="p-3.5">
                                <span className="px-2 py-1 bg-indigo-50 text-indigo-700 rounded text-xs">{s.className}</span>
                              </td>
                              <td className="p-3.5 text-center">
                                <div className="inline-flex flex-col items-center">
                                  <button
                                    onClick={() => setStudentDetailModal({ open: true, student: s })}
                                    className={`px-2.5 py-1 rounded-full text-xs font-bold border underline transition ${
                                      isLowAttendance
                                        ? 'bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100'
                                        : 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                                    }`}
                                    title="Bấm vào để xem chi tiết các ngày đi học"
                                  >
                                    {totalDaysInMonth} buổi (Xem chi tiết)
                                  </button>
                                  {isLowAttendance && (
                                    <span className="text-[10px] text-amber-600 font-semibold mt-0.5">Đi học ít (&lt;3)</span>
                                  )}
                                </div>
                              </td>
                              <td className="p-3.5 text-right">
                                <button
                                  onClick={() => handleToggleAttendance(s.id)}
                                  className={`px-4 py-2 rounded-xl text-xs font-bold transition shadow-xs ${
                                    isPresent
                                      ? 'bg-emerald-600 text-white hover:bg-emerald-700'
                                      : 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200'
                                  }`}
                                >
                                  {isPresent ? '✓ Đã điểm danh' : 'Điểm danh'}
                                </button>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {filteredAttendanceStudents.length > 0 && (
                  <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between flex-wrap gap-3">
                    <div className="flex items-center gap-2 text-xs text-slate-600">
                      <span>Hiển thị</span>
                      <select
                        value={rowsPerPage}
                        onChange={(e) => setRowsPerPage(Number(e.target.value))}
                        className="px-2 py-1 bg-white border border-slate-200 rounded-lg font-semibold"
                      >
                        <option value={10}>10</option>
                        <option value={20}>20</option>
                        <option value={50}>50</option>
                        <option value={100}>100</option>
                      </select>
                      <span>học sinh / trang (Tổng: {filteredAttendanceStudents.length})</span>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                        disabled={currentPage === 1}
                        className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                      >
                        Trước
                      </button>
                      <span className="text-xs font-semibold">Trang {currentPage} / {totalAttendancePages}</span>
                      <button
                        onClick={() => setCurrentPage(p => Math.min(p + 1, totalAttendancePages))}
                        disabled={currentPage === totalAttendancePages}
                        className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                      >
                        Sau
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ====================================================
              TAB: THI LÊN ĐAI (BELT PROMOTION EXAM)
          ==================================================== */}
          {activeTab === 'belt_exam' && (
            <div className="space-y-4 sm:space-y-5">
              {/* Thống kê số lượng võ sinh theo từng cấp đai */}
              <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2.5">
                {BELT_RANKS.map((b) => {
                  const count = activeStudentsList.filter((s) => (s.belt || 'Đai Trắng') === b.name).length;
                  const isSelected = beltFilter === b.name;
                  return (
                    <button
                      key={b.name}
                      onClick={() => setBeltFilter(isSelected ? 'ALL' : b.name)}
                      className={`p-3 rounded-2xl border text-left transition shadow-2xs ${
                        isSelected ? 'ring-2 ring-indigo-600 bg-white border-indigo-600' : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <span className={`inline-block px-2 py-0.5 rounded-full text-[10px] font-bold border ${b.bg} ${b.text} ${b.border}`}>
                        {b.name}
                      </span>
                      <p className="text-xl font-extrabold text-slate-800 mt-1.5">{count} <span className="text-[11px] font-normal text-slate-500">võ sinh</span></p>
                    </button>
                  );
                })}
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                {/* Header công cụ Thi lên đai */}
                <div className="p-4 border-b border-slate-200 bg-slate-50 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-sm sm:text-base text-slate-900 flex items-center gap-2">
                      <Award className="w-5 h-5 text-amber-500 shrink-0" /> Xét Duyệt & Xuất Danh Sách Thi Lên Đai
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Tích chọn các võ sinh đủ điều kiện dự thi. Khi xuất Excel, file sẽ tự động chia từng Sheet theo từng màu đai dự thi.
                    </p>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      onClick={handleSelectEligibleStudents}
                      className="px-3 py-2 bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-semibold shadow-2xs"
                    >
                      ✓ Chọn nhanh HS đủ ĐK (≥4 buổi)
                    </button>
                    <button
                      onClick={handleBulkPromoteBelt}
                      disabled={selectedExamIds.length === 0}
                      className="px-3.5 py-2 bg-amber-500 hover:bg-amber-600 disabled:bg-slate-200 disabled:text-slate-400 text-white rounded-xl text-xs font-bold shadow-xs transition"
                    >
                      Thăng Đai Hàng Loạt ({selectedExamIds.length})
                    </button>
                    <button
                      onClick={handleExportBeltExamExcel}
                      className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                      {selectedExamIds.length > 0 ? `Xuất DS Dự Thi (${selectedExamIds.length} em)` : 'Xuất Toàn Bộ DS Lọc'}
                    </button>
                  </div>
                </div>

                {/* Bộ lọc Tìm kiếm + Lớp + Màu Đai */}
                <div className="p-3.5 border-b border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5 bg-white">
                  <div className="flex items-center gap-2 bg-slate-50 px-3.5 py-2 rounded-xl border border-slate-200 flex-1">
                    <Search className="w-4 h-4 text-slate-400 shrink-0" />
                    <input
                      type="text"
                      placeholder="Tìm võ sinh theo tên hoặc mã HS..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full bg-transparent border-none text-sm focus:outline-none"
                    />
                  </div>
                  <div className="grid grid-cols-2 sm:flex items-center gap-2">
                    <select
                      value={classFilter}
                      onChange={(e) => setClassFilter(e.target.value)}
                      className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs sm:text-sm font-semibold text-slate-700"
                    >
                      <option value="ALL">Tất cả các lớp</option>
                      {courses.map((c) => (
                        <option key={c.id} value={c.name}>{c.name}</option>
                      ))}
                    </select>

                    <select
                      value={beltFilter}
                      onChange={(e) => setBeltFilter(e.target.value)}
                      className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs sm:text-sm font-semibold text-slate-700"
                    >
                      <option value="ALL">Tất cả màu đai</option>
                      {BELT_RANKS.map((b) => (
                        <option key={b.name} value={b.name}>{b.name}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Bảng danh sách xét duyệt thi lên đai */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                        <th className="p-3.5 w-10 text-center">
                          <input
                            type="checkbox"
                            checked={filteredExamStudents.length > 0 && filteredExamStudents.every((s) => selectedExamIds.includes(s.id))}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setSelectedExamIds(filteredExamStudents.map((s) => s.id));
                              } else {
                                setSelectedExamIds([]);
                              }
                            }}
                            className="w-4 h-4 rounded text-indigo-600 cursor-pointer"
                            title="Chọn tất cả danh sách đang lọc"
                          />
                        </th>
                        <th className="p-3.5">Mã HS</th>
                        <th className="p-3.5">Họ và Tên Võ Sinh</th>
                        <th className="p-3.5">Lớp</th>
                        <th className="p-3.5">Đai Hiện Tại (Đổi nhanh)</th>
                        <th className="p-3.5">Đai Dự Thi Tiếp Theo</th>
                        <th className="p-3.5 text-center">Chuyên Cần ({selectedMonth})</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-sm">
                      {filteredExamStudents.length === 0 ? (
                        <tr><td colSpan="7" className="p-8 text-center text-slate-400">Không tìm thấy võ sinh nào phù hợp.</td></tr>
                      ) : (
                        filteredExamStudents.map((s) => {
                          const currentBelt = s.belt || 'Đai Trắng';
                          const nextBelt = getNextBelt(currentBelt);
                          const currentStyle = getBeltStyle(currentBelt);
                          const nextStyle = getBeltStyle(nextBelt);
                          const attendedDays = monthlyAttendanceCounts[s.id] || 0;
                          const isChecked = selectedExamIds.includes(s.id);

                          return (
                            <tr key={s.id} className={isChecked ? 'bg-indigo-50/40 hover:bg-indigo-50/70' : 'hover:bg-slate-50'}>
                              <td className="p-3.5 text-center">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => {
                                    setSelectedExamIds((prev) =>
                                      prev.includes(s.id) ? prev.filter((id) => id !== s.id) : [...prev, s.id]
                                    );
                                  }}
                                  className="w-4 h-4 rounded text-indigo-600 cursor-pointer"
                                />
                              </td>
                              <td className="p-3.5 font-mono text-xs font-bold">{s.id}</td>
                              <td className="p-3.5">
                                <p className="font-semibold text-slate-900">{s.name}</p>
                                <p className="text-xs text-slate-500">SĐT: {s.phone || '—'}</p>
                              </td>
                              <td className="p-3.5">
                                <span className="px-2 py-1 bg-slate-100 text-slate-700 rounded text-xs">{s.className}</span>
                              </td>
                              <td className="p-3.5">
                                <select
                                  value={currentBelt}
                                  onChange={(e) => handleQuickChangeBelt(s, e.target.value)}
                                  className={`px-2.5 py-1 rounded-full text-xs font-bold border cursor-pointer focus:outline-none ${currentStyle.bg} ${currentStyle.text} ${currentStyle.border}`}
                                >
                                  {BELT_RANKS.map((b) => (
                                    <option key={b.name} value={b.name} className="bg-white text-slate-800 font-semibold">
                                      {b.name}
                                    </option>
                                  ))}
                                </select>
                              </td>
                              <td className="p-3.5">
                                <span className={`px-2.5 py-1 rounded-full text-xs font-bold border ${nextStyle.bg} ${nextStyle.text} ${nextStyle.border}`}>
                                  → {nextBelt}
                                </span>
                              </td>
                              <td className="p-3.5 text-center">
                                <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                                  attendedDays >= 4 ? 'bg-emerald-50 text-emerald-700' : 'bg-amber-50 text-amber-700'
                                }`}>
                                  {attendedDays} buổi {attendedDays >= 4 ? '(Đủ ĐK)' : '(Ít buổi)'}
                                </span>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ====================================================
              TAB: BÁO CÁO DOANH THU
          ==================================================== */}
          {activeTab === 'revenue' && (
            <div className="space-y-6">
              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                <div className="p-4 sm:p-5 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-sm sm:text-base text-slate-900 flex items-center gap-2"><DollarSign className="w-5 h-5 text-emerald-600 shrink-0" /> Doanh Thu Học Phí (Tháng {selectedMonth})</h3>
                    <p className="text-xs text-slate-500">Quản lý quỹ học phí học viên đóng hàng tháng</p>
                  </div>
                  <button onClick={handleExportTuitionExcel} className="w-full sm:w-auto justify-center px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs">
                    <FileSpreadsheet className="w-4 h-4" /> Xuất File Doanh Thu
                  </button>
                </div>
                <div className="p-4 sm:p-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 sm:gap-4 bg-emerald-50/30">
                  <div className="p-4 bg-white rounded-xl border border-emerald-100">
                    <p className="text-xs text-slate-500">Dự kiến thu</p>
                    <p className="text-lg font-bold text-slate-800 mt-1">{formatVND(stats.expectedRevenue)}</p>
                  </div>
                  <div className="p-4 bg-white rounded-xl border border-emerald-100">
                    <p className="text-xs text-emerald-700 font-medium">Tổng thực thu (Đã đóng)</p>
                    <p className="text-lg font-bold text-emerald-600 mt-1">{formatVND(stats.collectedRevenue)}</p>
                  </div>
                  <div className="p-4 bg-white rounded-xl border border-indigo-100">
                    <p className="text-xs text-indigo-700 font-medium">Thu Chuyển khoản</p>
                    <p className="text-lg font-bold text-indigo-600 mt-1">{formatVND(stats.transferRevenue)}</p>
                  </div>
                  <div className="p-4 bg-white rounded-xl border border-amber-100">
                    <p className="text-xs text-amber-700 font-medium">Thu Tiền mặt</p>
                    <p className="text-lg font-bold text-amber-600 mt-1">{formatVND(stats.cashRevenue)}</p>
                  </div>
                  <div className="p-4 bg-white rounded-xl border border-rose-100 sm:col-span-2 lg:col-span-1">
                    <p className="text-xs text-rose-700 font-medium">Còn thiếu (Chưa đóng)</p>
                    <p className="text-lg font-bold text-rose-600 mt-1">{formatVND(stats.pendingRevenue)}</p>
                  </div>
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                <div className="p-4 sm:p-5 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-sm sm:text-base text-slate-900 flex items-center gap-2"><Store className="w-5 h-5 text-indigo-600 shrink-0" /> Doanh Thu Bán Quần Áo & Dụng Cụ (Shop)</h3>
                    <p className="text-xs text-slate-500">Quỹ doanh thu bán đồ tập tách biệt hoàn toàn khỏi học phí</p>
                  </div>
                  <button onClick={handleExportShopExcel} className="w-full sm:w-auto justify-center px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs">
                    <FileSpreadsheet className="w-4 h-4" /> Xuất File Doanh Thu Shop Riêng
                  </button>
                </div>

                <div className="p-4 sm:p-5 bg-indigo-50/30 flex items-center justify-between border-b border-slate-100 flex-wrap gap-3">
                  <div className="space-y-1">
                    <p className="text-xs text-slate-500">Tổng doanh thu bán hàng tích lũy</p>
                    <p className="text-2xl font-extrabold text-indigo-700">{formatVND(shopStats.totalShopRevenue)}</p>
                    <p className="text-xs text-slate-600">
                      Số lượng đã bán: <strong className="text-slate-900">{shopStats.totalItemsSold} sản phẩm</strong>
                    </p>
                    <p className="text-xs text-emerald-700 font-semibold">
                      Số tiền lãi thu được: <strong className="text-sm text-emerald-600">{formatVND(shopStats.totalShopProfit)}</strong>
                    </p>
                  </div>
                  <button onClick={() => setSellModal({ open: true, productId: products[0]?.id || '', quantity: 1, buyerName: '', note: '' })} className="px-4 py-2.5 bg-indigo-600 text-white rounded-xl text-xs font-bold shadow-sm">
                    + Bán đồ ngay
                  </button>
                </div>

                {/* MOBILE CARD VIEW CHO LỊCH SỬ BÁN HÀNG */}
                <div className="divide-y divide-slate-200 md:hidden">
                  {shopOrders.length === 0 ? (
                    <div className="p-8 text-center text-sm text-slate-400">Chưa có giao dịch bán hàng nào.</div>
                  ) : (
                    paginatedShopOrders.map((o) => {
                      const prod = products.find((p) => p.id === o.productId);
                      const cost = o.costPrice !== undefined ? Number(o.costPrice) : Number(prod?.costPrice || 0);
                      const itemProfit = o.profit !== undefined ? Number(o.profit) : Number(o.totalAmount || 0) - cost * Number(o.quantity || 0);
                      return (
                        <div key={o.id} className="p-4 space-y-2 hover:bg-slate-50">
                          <div className="flex items-start justify-between gap-2">
                            <div>
                              <p className="font-bold text-sm text-slate-900">{o.productName} <span className="text-indigo-600">x{o.quantity}</span></p>
                              <p className="text-xs text-slate-500">{formatDateTime(o.date)}</p>
                            </div>
                            <div className="text-right">
                              <p className="font-bold text-sm text-indigo-600">{formatVND(o.totalAmount)}</p>
                              <p className="text-[11px] font-semibold text-emerald-600">Lãi: +{formatVND(itemProfit)}</p>
                            </div>
                          </div>
                          <p className="text-xs text-slate-600">
                            Người mua: <strong className="text-slate-800">{o.buyerName}</strong> {o.note && `• ${o.note}`}
                          </p>
                        </div>
                      );
                    })
                  )}
                </div>

                {/* DESKTOP TABLE CHO LỊCH SỬ BÁN HÀNG */}
                <div className="hidden md:block overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                        <th className="p-3.5">Thời Gian</th>
                        <th className="p-3.5">Sản Phẩm</th>
                        <th className="p-3.5">Số Lượng</th>
                        <th className="p-3.5">Đơn Giá</th>
                        <th className="p-3.5">Thành Tiền</th>
                        <th className="p-3.5">Tiền Lãi</th>
                        <th className="p-3.5">Người Mua / Ghi chú</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-sm">
                      {shopOrders.length === 0 ? (
                        <tr><td colSpan="7" className="p-8 text-center text-slate-400">Chưa có giao dịch bán hàng nào.</td></tr>
                      ) : (
                        paginatedShopOrders.map((o) => {
                          const prod = products.find((p) => p.id === o.productId);
                          const cost = o.costPrice !== undefined ? Number(o.costPrice) : Number(prod?.costPrice || 0);
                          const itemProfit = o.profit !== undefined ? Number(o.profit) : Number(o.totalAmount || 0) - cost * Number(o.quantity || 0);
                          return (
                            <tr key={o.id} className="hover:bg-slate-50">
                              <td className="p-3.5 text-xs text-slate-500 font-medium">{formatDateTime(o.date)}</td>
                              <td className="p-3.5 font-semibold text-slate-900">{o.productName}</td>
                              <td className="p-3.5 font-bold text-slate-700">{o.quantity}</td>
                              <td className="p-3.5">{formatVND(o.price)}</td>
                              <td className="p-3.5 font-bold text-indigo-600">{formatVND(o.totalAmount)}</td>
                              <td className="p-3.5 font-bold text-emerald-600">+{formatVND(itemProfit)}</td>
                              <td className="p-3.5 text-xs text-slate-600">
                                <p className="font-semibold text-slate-800">{o.buyerName}</p>
                                {o.note && <p className="text-[11px] text-slate-400">{o.note}</p>}
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {shopOrders.length > 0 && (
                  <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between flex-wrap gap-3">
                    <div className="flex items-center gap-2 text-xs text-slate-600">
                      <span>Hiển thị</span>
                      <select
                        value={rowsPerPage}
                        onChange={(e) => setRowsPerPage(Number(e.target.value))}
                        className="px-2 py-1 bg-white border border-slate-200 rounded-lg font-semibold"
                      >
                        <option value={10}>10</option>
                        <option value={20}>20</option>
                        <option value={50}>50</option>
                      </select>
                      <span>giao dịch / trang (Tổng: {shopOrders.length})</span>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                        disabled={currentPage === 1}
                        className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                      >
                        Trước
                      </button>
                      <span className="text-xs font-semibold">Trang {currentPage} / {totalShopPages}</span>
                      <button
                        onClick={() => setCurrentPage(p => Math.min(p + 1, totalShopPages))}
                        disabled={currentPage === totalShopPages}
                        className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                      >
                        Sau
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* ====================================================
              TAB: HỌC SINH ĐANG HỌC & THU HỌC PHÍ (THÊM BỘ LỌC LỚP ĐẦY ĐỦ)
          ==================================================== */}
          {(activeTab === 'tuition' || activeTab === 'students') && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-3.5 sm:p-4 border-b border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5 bg-slate-50">
                <div className="flex items-center gap-2.5 bg-white px-3.5 py-2 rounded-xl border border-slate-200 flex-1">
                  <Search className="w-4 h-4 text-slate-400 shrink-0" />
                  <input type="text" placeholder="Tìm kiếm học sinh theo tên, mã HS, số điện thoại..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full bg-transparent border-none text-sm focus:outline-none" />
                </div>

                <div className="grid grid-cols-2 sm:flex items-center gap-2 shrink-0">
                  <select
                    value={classFilter}
                    onChange={(e) => setClassFilter(e.target.value)}
                    className="px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs sm:text-sm font-semibold text-slate-700 focus:outline-none"
                  >
                    <option value="ALL">Tất cả các lớp</option>
                    {courses.map((c) => (
                      <option key={c.id} value={c.name}>{c.name}</option>
                    ))}
                  </select>

                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                    className="px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs sm:text-sm font-semibold text-slate-700 focus:outline-none"
                  >
                    <option value="ALL">Tất cả trạng thái</option>
                    <option value="PAID">Đã đóng học phí</option>
                    <option value="UNPAID">Chưa đóng học phí</option>
                  </select>
                </div>
              </div>

              {/* MOBILE CARD VIEW */}
              <div className="divide-y divide-slate-200 md:hidden">
                {paginatedStudents.length === 0 ? (
                  <div className="p-8 text-center text-sm text-slate-400">Không tìm thấy học sinh phù hợp.</div>
                ) : (
                  paginatedStudents.map((s) => (
                    <div key={s.id} className="p-4 space-y-3 hover:bg-slate-50">
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-mono text-xs font-bold px-2 py-0.5 bg-slate-100 text-slate-700 rounded">{s.id}</span>
                            <span className="px-2 py-0.5 bg-indigo-50 text-indigo-700 rounded text-xs font-medium">{s.className}</span>
                            {s.status === 'inactive' && (
                              <span className="px-2 py-0.5 bg-amber-50 text-amber-700 rounded text-[10px] font-semibold border border-amber-200">Đã nghỉ</span>
                            )}
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
                                <p className="text-[10px] text-slate-500 font-medium">{s.method} • {formatDateOnly(s.paidDate)}</p>
                              </div>
                            ) : (
                              <span className="inline-block px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">Chưa đóng</span>
                            )}
                          </div>
                        </div>
                      </div>

                      <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 flex-wrap">
                        {activeTab === 'tuition' && (
                          !s.paid ? (
                            <button onClick={() => openPaymentModal(s)} className="flex-1 py-2 bg-emerald-600 active:bg-emerald-700 text-white rounded-xl text-xs font-semibold text-center shadow-xs">
                              Thu học phí
                            </button>
                          ) : (
                            <button onClick={() => handleMarkUnpaid(s)} className="flex-1 py-2 bg-rose-50 active:bg-rose-100 text-rose-600 rounded-xl text-xs font-semibold text-center">
                              Hủy thu
                            </button>
                          )
                        )}
                        <button onClick={() => setStudentModal({ open: true, mode: 'edit', data: { ...s } })} className={`${activeTab === 'students' ? 'flex-1' : 'px-3'} py-2 bg-slate-100 text-slate-700 rounded-xl text-xs font-semibold text-center`}>Sửa</button>
                        {s.status !== 'inactive' && (
                          <button onClick={() => handleDeactivateStudent(s)} className={`${activeTab === 'students' ? 'flex-1' : 'px-3'} py-2 bg-amber-50 text-amber-700 rounded-xl text-xs font-semibold text-center`} title="Nghỉ">Nghỉ</button>
                        )}
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
                    {paginatedStudents.length === 0 ? (
                      <tr><td colSpan="6" className="p-8 text-center text-slate-400">Không tìm thấy học sinh phù hợp.</td></tr>
                    ) : (
                      paginatedStudents.map((s) => (
                        <tr key={s.id} className="hover:bg-slate-50">
                          <td className="p-3.5 font-mono text-xs font-bold">{s.id}</td>
                          <td className="p-3.5">
                            <div className="flex items-center gap-2">
                              <p className="font-semibold text-slate-900">{s.name}</p>
                              {s.status === 'inactive' && (
                                <span className="px-2 py-0.5 bg-amber-50 text-amber-700 rounded text-[10px] font-semibold border border-amber-200">Đã nghỉ</span>
                              )}
                            </div>
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
                                <p className="text-[11px] text-slate-500 font-medium">{s.method} • {formatDateOnly(s.paidDate)}</p>
                              </div>
                            ) : (
                              <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-rose-50 text-rose-700">Chưa đóng</span>
                            )}
                          </td>
                          <td className="p-3.5 text-right space-x-1.5 whitespace-nowrap">
                            {activeTab === 'tuition' && (
                              !s.paid ? (
                                <button onClick={() => openPaymentModal(s)} className="px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-semibold">Thu tiền</button>
                              ) : (
                                <button onClick={() => handleMarkUnpaid(s)} className="px-3 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Hủy thu</button>
                              )
                            )}
                            <button onClick={() => setStudentModal({ open: true, mode: 'edit', data: { ...s } })} className="px-2.5 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-semibold">Sửa</button>
                            {s.status !== 'inactive' && (
                              <button onClick={() => handleDeactivateStudent(s)} className="px-2.5 py-1.5 bg-amber-50 text-amber-700 hover:bg-amber-100 rounded-lg text-xs font-semibold">Nghỉ</button>
                            )}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {filteredStudents.length > 0 && (
                <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between flex-wrap gap-3">
                  <div className="flex items-center gap-2 text-xs text-slate-600">
                    <span>Hiển thị</span>
                    <select
                      value={rowsPerPage}
                      onChange={(e) => setRowsPerPage(Number(e.target.value))}
                      className="px-2 py-1 bg-white border border-slate-200 rounded-lg font-semibold"
                    >
                      <option value={10}>10</option>
                      <option value={20}>20</option>
                      <option value={50}>50</option>
                      <option value={100}>100</option>
                    </select>
                    <span>học sinh / trang (Tổng: {filteredStudents.length})</span>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                      disabled={currentPage === 1}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                    >
                      Trước
                    </button>
                    <span className="text-xs font-semibold">Trang {currentPage} / {totalStudentPages}</span>
                    <button
                      onClick={() => setCurrentPage(p => Math.min(p + 1, totalStudentPages))}
                      disabled={currentPage === totalStudentPages}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                    >
                      Sau
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ====================================================
              TAB: HỌC SINH ĐÃ NGHỈ (CÓ CARD MOBILE & BẢNG DESKTOP)
          ==================================================== */}
          {activeTab === 'inactive_students' && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-4 border-b border-slate-200 bg-amber-50/50">
                <h3 className="font-bold text-sm sm:text-base text-slate-900">Danh Sách Học Sinh Đã Nghỉ / Bảo Lưu</h3>
                <p className="text-xs text-slate-500">Dữ liệu bảo lưu an toàn. Có thể cho đi học lại bất cứ lúc nào.</p>
              </div>

              {/* MOBILE CARD CHO HỌC SINH ĐÃ NGHỈ */}
              <div className="divide-y divide-slate-200 md:hidden">
                {inactiveStudentsList.length === 0 ? (
                  <div className="p-8 text-center text-sm text-slate-400">Không có học sinh nào trong danh sách nghỉ học.</div>
                ) : (
                  inactiveStudentsList.map((s) => (
                    <div key={s.id} className="p-4 space-y-3">
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <span className="font-mono text-xs font-bold px-2 py-0.5 bg-slate-100 text-slate-600 rounded">{s.id}</span>
                          <span className="ml-2 px-2 py-0.5 bg-slate-100 rounded text-xs">{s.className}</span>
                          <p className="font-bold text-base text-slate-900 mt-1.5">{s.name}</p>
                          <p className="text-xs text-slate-500">{s.phone} • PH: {s.parentName || '—'}</p>
                        </div>
                        <span className="px-2.5 py-1 rounded-full text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200 shrink-0">Đã nghỉ</span>
                      </div>
                      <div className="flex items-center gap-2 pt-2 border-t border-slate-100">
                        <button onClick={() => handleRestoreStudent(s)} className="flex-1 py-2 bg-emerald-600 text-white rounded-xl text-xs font-semibold flex items-center justify-center gap-1">
                          <UserCheck className="w-3.5 h-3.5" /> Đi học lại
                        </button>
                        <button onClick={() => handleDeleteStudentPermanent(s.id)} className="px-3.5 py-2 bg-rose-50 text-rose-600 rounded-xl text-xs font-semibold flex items-center justify-center gap-1">
                          <Trash2 className="w-3.5 h-3.5" /> Xóa vĩnh viễn
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* DESKTOP TABLE CHO HỌC SINH ĐÃ NGHỈ */}
              <div className="hidden md:block overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                      <th className="p-3.5">Mã HS</th>
                      <th className="p-3.5">Họ Tên & SĐT</th>
                      <th className="p-3.5">Khóa Học Cũ</th>
                      <th className="p-3.5">Trạng Thái</th>
                      <th className="p-3.5 text-right">Khôi Phục / Xóa Vĩnh Viễn</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-sm">
                    {inactiveStudentsList.length === 0 ? (
                      <tr><td colSpan="5" className="p-8 text-center text-slate-400">Không có học sinh nào trong danh sách nghỉ học.</td></tr>
                    ) : (
                      inactiveStudentsList.map((s) => (
                        <tr key={s.id} className="hover:bg-slate-50">
                          <td className="p-3.5 font-mono text-xs font-bold text-slate-500">{s.id}</td>
                          <td className="p-3.5">
                            <p className="font-semibold text-slate-900">{s.name}</p>
                            <p className="text-xs text-slate-500">{s.phone} • PH: {s.parentName || '—'}</p>
                          </td>
                          <td className="p-3.5"><span className="px-2 py-0.5 bg-slate-100 rounded text-xs">{s.className}</span></td>
                          <td className="p-3.5"><span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-amber-50 text-amber-700 border border-amber-200">Đã nghỉ học</span></td>
                          <td className="p-3.5 text-right space-x-2">
                            <button onClick={() => handleRestoreStudent(s)} className="px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-semibold inline-flex items-center gap-1 shadow-xs">
                              <UserCheck className="w-3.5 h-3.5" /> Đi học lại
                            </button>
                            <button onClick={() => handleDeleteStudentPermanent(s.id)} className="px-3 py-1.5 bg-rose-50 text-rose-600 hover:bg-rose-100 rounded-lg text-xs font-semibold inline-flex items-center gap-1">
                              <Trash2 className="w-3.5 h-3.5" /> Xóa vĩnh viễn
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ====================================================
              TAB: QUẢN LÝ KHO (CÓ CARD MOBILE & TABLE DESKTOP)
          ==================================================== */}
          {activeTab === 'inventory' && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between flex-wrap gap-3">
                <div>
                  <h3 className="font-bold text-sm sm:text-base text-slate-900">Quản Lý Kho Quần Áo & Dụng Cụ Võ Thuật</h3>
                  <p className="text-xs text-slate-500">Hệ thống tự động cảnh báo đỏ khi số lượng tồn kho còn từ 2 sản phẩm trở xuống</p>
                </div>
                <div className="flex items-center gap-2 w-full sm:w-auto">
                  <button onClick={() => setSellModal({ open: true, productId: products[0]?.id || '', quantity: 1, buyerName: '', note: '' })} className="flex-1 sm:flex-initial justify-center px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs">
                    <ShoppingCart className="w-4 h-4" /> Bán hàng ngay
                  </button>
                  <button onClick={() => setProductModal({ open: true, mode: 'add', data: { id: `PROD_${Date.now()}`, name: '', costPrice: '', price: '', stock: '', category: 'Đồng phục' } })} className="flex-1 sm:flex-initial justify-center px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5">
                    <Plus className="w-4 h-4" /> Thêm sản phẩm
                  </button>
                </div>
              </div>

              {/* MOBILE CARD VIEW CHO KHO HÀNG */}
              <div className="divide-y divide-slate-200 md:hidden">
                {products.length === 0 ? (
                  <div className="p-8 text-center text-sm text-slate-400">Chưa có sản phẩm nào trong kho.</div>
                ) : (
                  paginatedProducts.map((p) => (
                    <div key={p.id} className={`p-4 space-y-3 ${Number(p.stock) <= 2 ? 'bg-rose-50/40' : ''}`}>
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <span className="px-2 py-0.5 bg-slate-100 rounded text-[11px] font-medium text-slate-700">{p.category}</span>
                          <p className="font-bold text-base text-slate-900 mt-1">{p.name}</p>
                          <p className="text-xs text-slate-500 mt-0.5">Giá vốn: <strong>{formatVND(p.costPrice || 0)}</strong></p>
                        </div>
                        <div className="text-right">
                          <p className="font-bold text-sm text-indigo-600">{formatVND(p.price)}</p>
                          <div className="mt-1">
                            {Number(p.stock) <= 2 ? (
                              <span className="px-2 py-0.5 rounded-full text-[11px] font-bold bg-rose-100 text-rose-700 border border-rose-300 inline-flex items-center gap-1">
                                Còn {p.stock} (Nhập thêm!)
                              </span>
                            ) : (
                              <span className="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                                Còn {p.stock} SP
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                        <button onClick={() => setProductModal({ open: true, mode: 'edit', data: { costPrice: 0, ...p } })} className="flex-1 py-2 bg-indigo-50 text-indigo-700 rounded-xl text-xs font-semibold">Sửa / Nhập thêm</button>
                        <button onClick={() => handleDeleteProduct(p.id)} className="px-4 py-2 bg-rose-50 text-rose-600 rounded-xl text-xs font-semibold">Xóa</button>
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* DESKTOP TABLE CHO KHO HÀNG */}
              <div className="hidden md:block overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                      <th className="p-3.5">Mã SP</th>
                      <th className="p-3.5">Tên Sản Phẩm</th>
                      <th className="p-3.5">Danh Mục</th>
                      <th className="p-3.5">Giá Vốn</th>
                      <th className="p-3.5">Giá Bán</th>
                      <th className="p-3.5">Tồn Kho</th>
                      <th className="p-3.5 text-right">Thao Tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-sm">
                    {products.length === 0 ? (
                      <tr><td colSpan="7" className="p-8 text-center text-slate-400">Chưa có sản phẩm nào trong kho.</td></tr>
                    ) : (
                      paginatedProducts.map((p) => (
                        <tr key={p.id} className={Number(p.stock) <= 2 ? 'bg-rose-50/40 hover:bg-rose-50/70' : 'hover:bg-slate-50'}>
                          <td className="p-3.5 font-mono text-xs font-bold text-slate-500">{p.id}</td>
                          <td className="p-3.5 font-semibold text-slate-900">{p.name}</td>
                          <td className="p-3.5"><span className="px-2 py-0.5 bg-slate-100 rounded text-xs font-medium text-slate-700">{p.category}</span></td>
                          <td className="p-3.5 font-semibold text-slate-600">{formatVND(p.costPrice || 0)}</td>
                          <td className="p-3.5 font-bold text-indigo-600">{formatVND(p.price)}</td>
                          <td className="p-3.5">
                            {Number(p.stock) <= 2 ? (
                              <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-rose-100 text-rose-700 border border-rose-300 inline-flex items-center gap-1">
                                <AlertCircle className="w-3.5 h-3.5" /> Còn {p.stock} (Cần nhập thêm!)
                              </span>
                            ) : (
                              <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                                Còn {p.stock} sản phẩm
                              </span>
                            )}
                          </td>
                          <td className="p-3.5 text-right space-x-1.5">
                            <button onClick={() => setProductModal({ open: true, mode: 'edit', data: { costPrice: 0, ...p } })} className="px-2.5 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-semibold">Sửa / Nhập thêm</button>
                            <button onClick={() => handleDeleteProduct(p.id)} className="px-2.5 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Xóa</button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {products.length > 0 && (
                <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between flex-wrap gap-3">
                  <div className="flex items-center gap-2 text-xs text-slate-600">
                    <span>Hiển thị</span>
                    <select
                      value={rowsPerPage}
                      onChange={(e) => setRowsPerPage(Number(e.target.value))}
                      className="px-2 py-1 bg-white border border-slate-200 rounded-lg font-semibold"
                    >
                      <option value={10}>10</option>
                      <option value={20}>20</option>
                      <option value={50}>50</option>
                    </select>
                    <span>sản phẩm / trang (Tổng: {products.length})</span>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                      disabled={currentPage === 1}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                    >
                      Trước
                    </button>
                    <span className="text-xs font-semibold">Trang {currentPage} / {totalProductPages}</span>
                    <button
                      onClick={() => setCurrentPage(p => Math.min(p + 1, totalProductPages))}
                      disabled={currentPage === totalProductPages}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-semibold disabled:opacity-50"
                    >
                      Sau
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ====================================================
              TAB: KHÓA HỌC
          ==================================================== */}
          {activeTab === 'courses' && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-4 border-b border-slate-200 bg-slate-50 font-bold text-sm">Danh mục Khóa học & Học phí</div>
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
                    {courses.length === 0 ? (
                      <tr><td colSpan="3" className="p-8 text-center text-slate-400">Chưa có khóa học nào.</td></tr>
                    ) : (
                      courses.map((c) => (
                        <tr key={c.id} className="hover:bg-slate-50">
                          <td className="p-3.5 font-semibold">{c.name}</td>
                          <td className="p-3.5 font-bold text-indigo-600">{formatVND(c.fee)}</td>
                          <td className="p-3.5 text-right space-x-1.5 whitespace-nowrap">
                            <button onClick={() => setCourseModal({ open: true, mode: 'edit', data: { ...c } })} className="px-2.5 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-semibold">Sửa</button>
                            <button onClick={async () => {
                              if (!window.confirm(`Xóa khóa học "${c.name}"?`)) return;
                              setCourses((prev) => prev.filter((item) => item.id !== c.id));
                              if (db) await deleteDoc(doc(db, 'courses', c.id)).catch(console.error);
                              showToast('Đã xóa khóa học', 'warning');
                            }} className="px-2.5 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Xóa</button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </main>
      </div>

      {/* ====================================================
          MODALS (ĐÃ BỔ SUNG LABEL RÕ RÀNG CHO MỌI Ô INPUT)
      ==================================================== */}
      {studentDetailModal.open && studentDetailModal.student && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 sm:p-6 shadow-xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between pb-3 border-b">
              <div>
                <h3 className="text-base font-bold text-slate-900">{studentDetailModal.student.name}</h3>
                <p className="text-xs text-slate-500">Mã: {studentDetailModal.student.id} • Lớp: {studentDetailModal.student.className}</p>
              </div>
              <button onClick={() => setStudentDetailModal({ open: false, student: null })} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-2">
              <p className="text-xs font-bold text-slate-700">
                Lịch sử đi học trong tháng {selectedMonth}: <span className="text-indigo-600 font-extrabold">{monthlyAttendanceCounts[studentDetailModal.student.id] || 0} buổi</span>
              </p>
              
              <div className="max-h-60 overflow-y-auto border border-slate-200 rounded-xl p-3 space-y-1.5 bg-slate-50">
                {getStudentAttendanceDetail(studentDetailModal.student.id).length === 0 ? (
                  <p className="text-xs text-slate-400 text-center py-4">Học sinh chưa đi học ngày nào trong tháng này.</p>
                ) : (
                  getStudentAttendanceDetail(studentDetailModal.student.id).map((dateStr, idx) => (
                    <div key={idx} className="flex items-center justify-between text-xs bg-white px-3 py-2 rounded-lg border border-slate-100 shadow-2xs">
                      <span className="font-semibold text-slate-700">Buổi {idx + 1}: Ngày {formatDateOnly(dateStr)}</span>
                      <span className="px-2 py-0.5 bg-emerald-50 text-emerald-700 rounded-full text-[10px] font-bold">✓ Có mặt</span>
                    </div>
                  ))
                )}
              </div>
            </div>

            <div className="flex justify-end pt-2 border-t">
              <button
                onClick={() => setStudentDetailModal({ open: false, student: null })}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {excelChoiceModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-xl w-full p-5 sm:p-6 shadow-xl border border-slate-200 space-y-5">
            <div className="flex items-center justify-between pb-3 border-b">
              <h3 className="text-base font-bold text-slate-900">Nhập Danh Sách Học Sinh Bằng Excel</h3>
              <button onClick={() => setExcelChoiceModal(false)} className="text-slate-400 hover:text-slate-600"><X className="w-5 h-5" /></button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div className="p-5 rounded-2xl border-2 border-dashed border-slate-200 bg-slate-50 flex flex-col items-center text-center justify-between space-y-3">
                <div className="space-y-1">
                  <div className="w-10 h-10 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center mx-auto mb-2">
                    <FileText className="w-5 h-5" />
                  </div>
                  <h4 className="font-bold text-sm text-slate-800">Bạn chưa có file mẫu?</h4>
                  <p className="text-xs text-slate-500">Tải file mẫu Excel chuẩn để điền thông tin.</p>
                </div>
                <button
                  onClick={handleDownloadTemplate}
                  className="w-full py-2.5 bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-semibold shadow-xs transition"
                >
                  Tải File Mẫu Excel
                </button>
              </div>

              <div className="p-5 rounded-2xl border-2 border-dashed border-indigo-200 bg-indigo-50/40 flex flex-col items-center text-center justify-between space-y-3">
                <div className="space-y-1">
                  <div className="w-10 h-10 rounded-xl bg-indigo-100 text-indigo-600 flex items-center justify-center mx-auto mb-2">
                    <Upload className="w-5 h-5" />
                  </div>
                  <h4 className="font-bold text-sm text-indigo-900">Bạn đã có file?</h4>
                  <p className="text-xs text-indigo-600/80">Chọn file Excel để hệ thống tự động kiểm tra.</p>
                </div>
                <button
                  onClick={() => {
                    setExcelChoiceModal(false);
                    fileInputRef.current?.click();
                  }}
                  className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold shadow-sm transition"
                >
                  Chọn & Tải File Lên
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {excelPreviewModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-xl w-full p-5 sm:p-6 shadow-xl border border-slate-200 space-y-4">
            <h3 className="text-base font-bold text-slate-900">Kiểm tra dữ liệu Import ({excelPreviewModal.previewRows.length} học sinh)</h3>
            <p className="text-xs text-slate-500">File: <span className="font-semibold text-slate-700">{excelPreviewModal.fileName}</span></p>

            {excelPreviewModal.missingCourses.length > 0 && (
              <div className="p-3 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 space-y-1">
                <p className="font-bold flex items-center gap-1"><AlertCircle className="w-4 h-4" /> Khóa học chưa tồn tại trên hệ thống:</p>
                <p className="font-semibold">{excelPreviewModal.missingCourses.join(', ')}</p>
                <p>Vui lòng tạo khóa học trước khi import!</p>
              </div>
            )}

            <div className="max-h-60 overflow-y-auto border border-slate-200 rounded-xl">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 border-b">
                  <tr>
                    <th className="p-2">Mã HS (Sinh)</th>
                    <th className="p-2">Họ Tên</th>
                    <th className="p-2">Khóa Học</th>
                    <th className="p-2">Học Phí (Match)</th>
                    <th className="p-2">SĐT</th>
                  </tr>
                </thead>
                <tbody className="divide-y">
                  {excelPreviewModal.previewRows.map((r, i) => (
                    <tr key={i} className={!r.isValidPhone || !r.isValidCourse ? 'bg-rose-50 text-rose-700 font-semibold' : ''}>
                      <td className="p-2 font-mono font-bold text-indigo-600">{r.id}</td>
                      <td className="p-2 font-semibold">{r.name}</td>
                      <td className="p-2">
                        {r.className} {!r.isValidCourse && <span className="text-[10px] bg-rose-200 text-rose-800 px-1 rounded ml-1">Chưa có lớp</span>}
                      </td>
                      <td className="p-2 font-bold">{formatVND(r.tuitionFee)}</td>
                      <td className="p-2">
                        {r.phone} {!r.isValidPhone && <span className="text-[10px] bg-rose-200 text-rose-800 px-1 rounded ml-1">Sai SĐT</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button onClick={() => setExcelPreviewModal({ open: false, previewRows: [], fileName: '', missingCourses: [] })} className="px-4 py-2 bg-slate-100 text-slate-600 rounded-xl text-xs">Hủy</button>
              <button
                onClick={handleConfirmImport}
                disabled={excelPreviewModal.missingCourses.length > 0}
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold"
              >
                Xác Nhận Import
              </button>
            </div>
          </div>
        </div>
      )}

      {sellModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 sm:p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2"><ShoppingCart className="w-5 h-5 text-indigo-600" /> Bán Quần Áo / Dụng Cụ</h3>
            <form onSubmit={handleExecuteSale} className="space-y-4 mt-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Chọn sản phẩm *</label>
                <select value={sellModal.productId} onChange={(e) => setSellModal({ ...sellModal, productId: e.target.value })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white font-semibold">
                  {products.map((p) => <option key={p.id} value={p.id}>{p.name} — {formatVND(p.price)} (Còn: {p.stock})</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Số lượng *</label>
                <input type="number" required min="1" value={sellModal.quantity} onChange={(e) => setSellModal({ ...sellModal, quantity: e.target.value })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Tên võ sinh / người mua</label>
                <input type="text" placeholder="VD: Nguyễn Minh Anh" value={sellModal.buyerName} onChange={(e) => setSellModal({ ...sellModal, buyerName: e.target.value })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div className="flex justify-end gap-2 pt-3 border-t">
                <button type="button" onClick={() => setSellModal({ ...sellModal, open: false })} className="px-4 py-2 text-xs text-slate-600 bg-slate-100 rounded-xl">Hủy</button>
                <button type="submit" className="px-5 py-2 bg-indigo-600 text-white text-xs font-semibold rounded-xl">Xác Nhận Bán</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {productModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 sm:p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold mb-4">{productModal.mode === 'add' ? 'Thêm Sản Phẩm Mới' : 'Sửa Sản Phẩm'}</h3>
            <form onSubmit={handleSaveProduct} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Tên sản phẩm *</label>
                <input type="text" required placeholder="Tên sản phẩm" value={productModal.data.name} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Danh mục *</label>
                <select value={productModal.data.category} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, category: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white">
                  <option value="Đồng phục">Đồng phục</option>
                  <option value="Đai võ">Đai võ</option>
                  <option value="Bảo hộ">Bảo hộ</option>
                  <option value="Khác">Khác</option>
                </select>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Giá vốn (VNĐ) *</label>
                  <input type="number" required step="5000" placeholder="VD: 100000" value={productModal.data.costPrice} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, costPrice: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-semibold text-slate-700" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Giá bán (VNĐ) *</label>
                  <input type="number" required step="5000" placeholder="VD: 200000" value={productModal.data.price} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, price: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
                </div>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Số lượng tồn kho *</label>
                <input type="number" required placeholder="Tồn kho" value={productModal.data.stock} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, stock: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold" />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setProductModal({ ...productModal, open: false })} className="px-4 py-2 text-xs text-slate-600 bg-slate-100 rounded-xl">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs font-semibold rounded-xl">Lưu</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {paymentModal.open && paymentModal.student && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 sm:p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold text-slate-900">Xác Nhận Thu Học Phí — {paymentModal.student.name}</h3>
            <form onSubmit={handleSavePaymentModal} className="space-y-4 mt-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Phương thức thanh toán *</label>
                <select value={paymentModal.method} onChange={(e) => setPaymentModal({ ...paymentModal, method: e.target.value })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white font-semibold text-indigo-600">
                  <option value="Chuyển khoản">Chuyển khoản ngân hàng</option>
                  <option value="Tiền mặt">Tiền mặt tại trung tâm</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Ngày đóng tiền (Ngày/Tháng/Năm) *</label>
                <input
                  type="date"
                  required
                  value={paymentModal.paidDate}
                  onChange={(e) => setPaymentModal({ ...paymentModal, paidDate: e.target.value })}
                  className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-semibold text-slate-700"
                />
              </div>
              <div className="flex justify-end gap-2 pt-3 border-t">
                <button type="button" onClick={() => setPaymentModal({ ...paymentModal, open: false })} className="px-4 py-2.5 rounded-xl text-xs font-semibold bg-slate-100">Hủy</button>
                <button type="submit" className="px-5 py-2.5 rounded-xl text-xs font-semibold bg-emerald-600 text-white">Xác Nhận Thu</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL KHÓA HỌC (ĐÃ BỔ SUNG LABEL RÕ RÀNG) */}
      {courseModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 sm:p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold mb-4">{courseModal.mode === 'add' ? 'Thêm Khóa Học Mới' : 'Sửa Khóa Học'}</h3>
            <form onSubmit={async (e) => {
              e.preventDefault();
              const { mode, data } = courseModal;
              const clean = { id: data.id || `COURSE_${Date.now()}`, name: data.name.trim(), fee: Number(data.fee) || 0, description: data.description || '' };
              if (!clean.name) return;
              setCourses((prev) => mode === 'add' ? [...prev, clean] : prev.map((c) => c.id === clean.id ? clean : c));
              if (db) await setDoc(doc(db, 'courses', clean.id), clean).catch(console.error);
              setCourseModal({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' } });
              showToast('Đã lưu khóa học!');
            }} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Tên Lớp / Khóa Học *</label>
                <input type="text" required placeholder="VD: Karate Cơ Bản" value={courseModal.data.name} onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Mức Học Phí Mặc Định (VNĐ) *</label>
                <input type="number" required placeholder="VD: 500000" value={courseModal.data.fee} onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, fee: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setCourseModal({ ...courseModal, open: false })} className="px-4 py-2 text-xs text-slate-600 bg-slate-100 rounded-xl">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs font-semibold rounded-xl">Lưu</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL HỌC SINH (ĐÃ BỔ SUNG LABEL RÕ RÀNG CHO TỪNG Ô) */}
      {studentModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 sm:p-6 shadow-xl border border-slate-200 my-auto">
            <h3 className="text-base font-bold mb-4">{studentModal.mode === 'add' ? 'Thêm Hồ Sơ Học Sinh' : 'Sửa Hồ Sơ Học Sinh'}</h3>
            <form onSubmit={async (e) => {
              e.preventDefault();
              const { mode, data } = studentModal;
              if (!data.name.trim() || !data.id.trim()) return;

              if (!validatePhone10Digits(data.phone)) {
                alert('Số điện thoại không hợp lệ! SĐT bắt buộc phải có đúng 10 chữ số và bắt đầu bằng số 0 (VD: 0932766613).');
                return;
              }

              const studentObj = {
                id: data.id.trim().toUpperCase(),
                name: data.name.trim(),
                className: data.className.trim(),
                phone: data.phone.trim(),
                parentName: (data.parentName || '').trim(),
                tuitionFee: Number(data.tuitionFee) || 0,
                status: data.status || 'active',
                belt: data.belt || 'Đai Trắng'
              };
              setStudents((prev) => mode === 'add' ? [...prev, studentObj] : prev.map((s) => s.id === studentObj.id ? studentObj : s));
              if (db) await setDoc(doc(db, 'students', studentObj.id), studentObj).catch(console.error);
              setStudentModal({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '', status: 'active' } });
              showToast('Đã lưu học sinh!');
            }} className="space-y-3.5">
              <div className="grid grid-cols-3 gap-3">
                <div className="col-span-1">
                  <label className="block text-xs font-bold text-slate-700 mb-1">Mã HS *</label>
                  <input type="text" required placeholder="HS001" value={studentModal.data.id} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, id: e.target.value } })} className="w-full px-3 py-2.5 border rounded-xl text-sm uppercase font-mono" />
                </div>
                <div className="col-span-2">
                  <label className="block text-xs font-bold text-slate-700 mb-1">Họ và Tên Học Sinh *</label>
                  <input type="text" required placeholder="Nhập họ tên" value={studentModal.data.name} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
                </div>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Khóa Học / Lớp *</label>
                <select value={studentModal.data.className} onChange={(e) => {
                  const found = courses.find((c) => c.name === e.target.value);
                  setStudentModal({ ...studentModal, data: { ...studentModal.data, className: e.target.value, tuitionFee: found ? found.fee : studentModal.data.tuitionFee } });
                }} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white">
                  {courses.length === 0 && (
                    <option value={studentModal.data.className || 'Lớp Cơ Bản'}>
                      {studentModal.data.className || 'Lớp Cơ Bản'}
                    </option>
                  )}
                  {courses.map((c) => <option key={c.id} value={c.name}>{c.name} — {formatVND(c.fee)}</option>)}
                  {courses.length > 0 && studentModal.data.className && !courses.some((c) => c.name === studentModal.data.className) && (
                    <option value={studentModal.data.className}>{studentModal.data.className}</option>
                  )}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Mức Học Phí Tháng (VNĐ) *</label>
                <input type="number" required placeholder="Học phí" value={studentModal.data.tuitionFee} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, tuitionFee: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Họ Tên Phụ Huynh</label>
                <input type="text" placeholder="Nhập tên phụ huynh" value={studentModal.data.parentName || ''} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, parentName: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">SĐT Phụ Huynh (Bắt buộc 10 số) *</label>
                <input type="tel" required placeholder="VD: 0932766613" value={studentModal.data.phone} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, phone: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
                <p className="text-[10px] text-slate-400 mt-1">Định dạng: Gồm 10 chữ số, bắt đầu bằng số 0</p>
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setStudentModal({ ...studentModal, open: false })} className="px-4 py-2 text-xs text-slate-600 bg-slate-100 rounded-xl">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs font-semibold rounded-xl">Lưu Hồ Sơ</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}