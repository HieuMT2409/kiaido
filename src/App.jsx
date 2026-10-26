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
  BarChart3,
  Filter,
  FileText,
  AlertCircle,
  BookPlus,
  UserPlus,
  UserPlus2
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

export default function App() {
  const [courses, setCourses] = useState([]);
  const [students, setStudents] = useState([]);
  const [payments, setPayments] = useState({});
  const [products, setProducts] = useState([]);
  const [shopOrders, setShopOrders] = useState([]);

  const [firebaseStatus, setFirebaseStatus] = useState(db ? 'connecting' : 'offline');
  const [selectedMonth, setSelectedMonth] = useState('2026-10');
  
  const [activeTab, setActiveTab] = useState('tuition');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  
  const [searchQuery, setSearchQuery] = useState('');
  const [classFilter, setClassFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');

  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);

  const [showNotifPanel, setShowNotifPanel] = useState(false);
  const [paymentModal, setPaymentModal] = useState({ open: false, student: null, paidDate: '', method: 'Chuyển khoản', note: '' });
  const [studentModal, setStudentModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '', status: 'active' } });
  const [courseModal, setCourseModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', fee: '', description: '' } });
  
  const [productModal, setProductModal] = useState({ open: false, mode: 'add', data: { id: '', name: '', price: '', stock: '', category: 'Đồng phục' } });
  const [sellModal, setSellModal] = useState({ open: false, productId: '', quantity: 1, buyerName: '', note: '' });

  const [excelChoiceModal, setExcelChoiceModal] = useState(false);
  const [excelPreviewModal, setExcelPreviewModal] = useState({ open: false, previewRows: [], fileName: '', missingCourses: [] });
  const fileInputRef = useRef(null);

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

    return () => {
      unsubCourses();
      unsubStudents();
      unsubPayments();
      unsubProducts();
      unsubShopOrders();
    };
  }, []);

  useEffect(() => {
    setCurrentPage(1);
  }, [activeTab, searchQuery, statusFilter, rowsPerPage]);

  const currentMonthPayments = payments[selectedMonth] || {};

  const activeStudentsList = useMemo(() => students.filter((s) => s.status !== 'inactive'), [students]);
  const inactiveStudentsList = useMemo(() => students.filter((s) => s.status === 'inactive'), [students]);

  const enrichedStudents = useMemo(() => {
    return activeStudentsList.map((student) => {
      const paymentInfo = currentMonthPayments[student.id] || { paid: false, paidDate: null, method: 'Chuyển khoản', note: '' };
      return {
        ...student,
        paid: Boolean(paymentInfo.paid),
        paidDate: paymentInfo.paidDate || null,
        method: paymentInfo.method || 'Chuyển khoản',
        note: paymentInfo.note || ''
      };
    });
  }, [activeStudentsList, currentMonthPayments]);

  const unpaidStudents = useMemo(() => enrichedStudents.filter((s) => !s.paid), [enrichedStudents]);
  const paidStudents = useMemo(() => enrichedStudents.filter((s) => s.paid), [enrichedStudents]);

  const stats = useMemo(() => {
    const expectedRevenue = enrichedStudents.reduce((sum, s) => sum + Number(s.tuitionFee || 0), 0);
    const collectedRevenue = paidStudents.reduce((sum, s) => sum + Number(s.tuitionFee || 0), 0);
    const pendingRevenue = expectedRevenue - collectedRevenue;
    const completionRate = enrichedStudents.length > 0 ? Math.round((paidStudents.length / enrichedStudents.length) * 100) : 0;
    return { expectedRevenue, collectedRevenue, pendingRevenue, completionRate };
  }, [enrichedStudents, paidStudents]);

  const shopStats = useMemo(() => {
    const totalShopRevenue = shopOrders.reduce((sum, o) => sum + Number(o.totalAmount || 0), 0);
    const totalItemsSold = shopOrders.reduce((sum, o) => sum + Number(o.quantity || 0), 0);
    return { totalShopRevenue, totalItemsSold };
  }, [shopOrders]);

  // BIỂU ĐỒ 7 NGÀY (ĐÃ AN TOÀN HÓA DỮ LIỆU ĐỂ TRÁNH LỖI TRẮNG TRANG)
  const last7DaysChartData = useMemo(() => {
    const days = [];
    const today = new Date();
    for (let i = 6; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(today.getDate() - i);
      const dateStr = d.toISOString().slice(0, 10);
      const label = `${d.getDate()}/${d.getMonth() + 1}`;

      const shopRev = (shopOrders || [])
        .filter(o => o && o.date && String(o.date).slice(0, 10) === dateStr)
        .reduce((sum, o) => sum + Number(o.totalAmount || 0), 0);

      let tuitionRev = 0;
      if (payments && typeof payments === 'object') {
        Object.keys(payments).forEach(mKey => {
          const mObj = payments[mKey];
          if (mObj && typeof mObj === 'object') {
            Object.keys(mObj).forEach(sId => {
              const payInfo = mObj[sId];
              if (payInfo && payInfo.paid && payInfo.paidDate && String(payInfo.paidDate).slice(0, 10) === dateStr) {
                tuitionRev += Number(payInfo.amount || 0);
              }
            });
          }
        });
      }

      days.push({
        label,
        tuition: tuitionRev,
        shop: shopRev,
        total: tuitionRev + shopRev
      });
    }
    const maxVal = Math.max(...days.map(d => d.total), 100000);
    return { days, maxVal };
  }, [shopOrders, payments]);

  const filteredStudents = useMemo(() => {
    return enrichedStudents.filter((s) => {
      const matchSearch = s.name.toLowerCase().includes(searchQuery.toLowerCase()) || s.id.toLowerCase().includes(searchQuery.toLowerCase()) || (s.phone && s.phone.includes(searchQuery));
      const matchClass = classFilter === 'ALL' || s.className === classFilter;
      const matchStatus = statusFilter === 'ALL' || (statusFilter === 'PAID' && s.paid) || (statusFilter === 'UNPAID' && !s.paid);
      return matchSearch && matchClass && matchStatus;
    });
  }, [enrichedStudents, searchQuery, classFilter, statusFilter]);

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
    const nowIso = new Date().toISOString().slice(0, 16);
    setPaymentModal({ open: true, student, paidDate: student.paidDate || nowIso, method: student.method || 'Chuyển khoản', note: student.note || '' });
  };

  const handleSavePaymentModal = async (e) => {
    e.preventDefault();
    const { student, paidDate, method, note } = paymentModal;
    if (!student) return;

    const updatedMonth = {
      ...(payments[selectedMonth] || {}),
      [student.id]: { paid: true, paidDate: paidDate || new Date().toISOString().slice(0, 16), method, note, amount: Number(student.tuitionFee) }
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
    if (!window.confirm(`Bạn có muốn cho học sinh ${student.name} nghỉ học không?`)) return;
    const updated = { ...student, status: 'inactive' };
    setStudents((prev) => prev.map((s) => s.id === student.id ? updated : s));
    if (db) await setDoc(doc(db, 'students', student.id), updated).catch(console.error);
    showToast(`Đã chuyển học sinh ${student.name} sang danh sách nghỉ học.`, 'warning');
  };

  const handleRestoreStudent = async (student) => {
    const updated = { ...student, status: 'active' };
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
      price: Number(data.price) || 0,
      stock: Number(data.stock) || 0,
      category: data.category || 'Đồng phục'
    };

    setProducts((prev) => mode === 'add' ? [...prev, prodObj] : prev.map((p) => p.id === prodId ? prodObj : p));
    if (db) {
      await setDoc(doc(db, 'products', prodId), prodObj).catch((err) => showToast(err.message, 'error'));
    }
    setProductModal({ open: false, mode: 'add', data: { id: '', name: '', price: '', stock: '', category: 'Đồng phục' } });
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
    const totalAmount = product.price * qty;
    const orderId = `ORDER_${Date.now()}`;
    const orderObj = {
      id: orderId,
      productId: product.id,
      productName: product.name,
      price: product.price,
      quantity: qty,
      totalAmount,
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
    showToast(`Bán thành công ${qty} ${product.name}! Thu về ${formatVND(totalAmount)}`);
  };

  const handleExportTuitionExcel = () => {
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
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(studentRows), `HocPhi_${selectedMonth}`);
    XLSX.writeFile(wb, `BaoCaoHocPhi_${selectedMonth}.xlsx`);
    showToast('Xuất báo cáo học phí thành công!');
  };

  const handleExportStudentListExcel = () => {
    const collator = new Intl.Collator('vi', { sensitivity: 'variant' });

    const sortedStudents = [...activeStudentsList].sort((a, b) => {
      const partsA = a.name.trim().split(/\s+/);
      const partsB = b.name.trim().split(/\s+/);

      // Tên = từ cuối cùng
      const givenA = partsA[partsA.length - 1];
      const givenB = partsB[partsB.length - 1];

      const byGiven = collator.compare(givenA, givenB);
      if (byGiven !== 0) return byGiven;

      // Nếu trùng tên thì so tiếp họ + tên đệm
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
      'Đơn Giá (VNĐ)': Number(o.price),
      'Tổng Tiền (VNĐ)': Number(o.totalAmount),
      'Người Mua': o.buyerName,
      'Ghi Chú': o.note || ''
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(shopRows), 'DoanhThuBanHang');
    XLSX.writeFile(wb, 'BaoCaoDoanhThuBanHang_Shop.xlsx');
    showToast('Xuất báo cáo doanh thu bán hàng thành công!');
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 flex flex-col md:flex-row relative">
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg text-white text-sm font-medium bg-slate-900">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast.message}</span>
        </div>
      )}

      <input type="file" ref={fileInputRef} onChange={handleFileChange} accept=".xlsx, .xls, .csv" className="hidden" />

      {/* ================= SIDEBAR ================= */}
      {isMobileMenuOpen && (
        <div 
          className="fixed inset-0 bg-black/50 z-40 md:hidden backdrop-blur-xs"
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}

      <aside className={`fixed md:static inset-y-0 left-0 z-50 w-64 bg-black text-slate-300 flex flex-col justify-between border-r border-slate-800 shrink-0 transform transition-transform duration-300 ease-in-out ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div>
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
              className="md:hidden text-slate-400 hover:text-white p-1"
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
              <DollarSign className="w-4 h-4" /> Thu Học Phí
            </button>
            <button
              onClick={() => { setActiveTab('revenue'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'revenue' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <TrendingUp className="w-4 h-4" /> Báo Cáo Doanh Thu
            </button>
            <button
              onClick={() => { setActiveTab('students'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'students' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <Users className="w-4 h-4" /> Danh sách Học Sinh ({activeStudentsList.length})
            </button>
            <button
              onClick={() => { setActiveTab('inactive_students'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'inactive_students' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <UserX className="w-4 h-4" /> Học Sinh Đã Nghỉ ({inactiveStudentsList.length})
            </button>
            <button
              onClick={() => { setActiveTab('courses'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'courses' ? 'bg-red-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <BookOpen className="w-4 h-4" /> Khóa Học & Lớp ({courses.length})
            </button>

            <p className="px-3 pt-4 pb-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-500">Kho Quần Áo & Dụng Cụ</p>
            <button
              onClick={() => { setActiveTab('inventory'); setIsMobileMenuOpen(false); }}
              className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-medium transition ${activeTab === 'inventory' ? 'bg-indigo-600 text-white shadow-md' : 'hover:bg-slate-900 text-slate-300'}`}
            >
              <Package className="w-4 h-4" /> Quản Lý Kho ({products.length})
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
      <div className="flex-1 flex flex-col min-w-0 overflow-y-auto">
        <header className="bg-white border-b border-slate-200 sticky top-0 z-20 px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3 shadow-xs">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setIsMobileMenuOpen(true)}
              className="md:hidden p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 transition"
              title="Mở menu"
            >
              <Menu className="w-5 h-5" />
            </button>

            <span className="text-sm font-bold text-slate-800 uppercase tracking-wide">
              {activeTab === 'tuition' && 'Quản Lý Thu Học Phí'}
              {activeTab === 'revenue' && 'Báo Cáo Doanh Thu'}
              {activeTab === 'students' && 'Danh Sách Học Sinh Đang Học'}
              {activeTab === 'inactive_students' && 'Danh Sách Học Sinh Đã Nghỉ Học'}
              {activeTab === 'courses' && 'Quản Lý Khóa Học & Học Phí'}
              {activeTab === 'inventory' && 'Quản Lý Kho Hàng & Sản Phẩm'}
            </span>
          </div>

          <div className="flex items-center gap-3 flex-wrap">
            <div className="flex items-center bg-slate-100 rounded-xl p-1 border border-slate-200">
              <button onClick={() => changeMonth(-1)} className="p-1 hover:bg-white rounded-lg transition"><ChevronLeft className="w-4 h-4" /></button>
              <div className="flex items-center gap-1.5 px-2 text-xs font-semibold text-slate-800">
                <Calendar className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
                <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} className="bg-transparent border-none focus:outline-none cursor-pointer text-xs" />
              </div>
              <button onClick={() => changeMonth(1)} className="p-1 hover:bg-white rounded-lg transition"><ChevronRight className="w-4 h-4" /></button>
            </div>

            <div className="relative">
              <button onClick={() => setShowNotifPanel(!showNotifPanel)} className={`relative p-2 rounded-xl border transition ${unpaidStudents.length > 0 ? 'bg-rose-50 border-rose-200 text-rose-600' : 'bg-slate-100 border-slate-200 text-slate-600'}`}>
                <Bell className="w-4 h-4" />
                {unpaidStudents.length > 0 && <span className="absolute -top-1 -right-1 bg-rose-600 text-white text-[10px] font-bold w-4 h-4 rounded-full flex items-center justify-center">{unpaidStudents.length}</span>}
              </button>
              {showNotifPanel && (
                <div className="absolute right-0 mt-2 w-72 sm:w-80 bg-white rounded-2xl shadow-xl border border-slate-200 z-50 overflow-hidden">
                  <div className="p-3 bg-rose-50 border-b border-rose-100 font-semibold text-sm text-rose-800">Chưa đóng học phí: {unpaidStudents.length} học sinh</div>
                  <div className="max-h-60 overflow-y-auto divide-y divide-slate-100">
                    {unpaidStudents.map((s) => (
                      <div key={s.id} className="p-2.5 flex items-center justify-between text-xs">
                        <span className="font-medium">{s.name}</span>
                        <button onClick={() => openPaymentModal(s)} className="px-2.5 py-1 bg-emerald-600 text-white rounded-lg">Thu ngay</button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </header>

        <main className="max-w-7xl mx-auto w-full px-4 sm:px-6 py-6 space-y-6">
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <div className="flex items-center gap-2 flex-wrap">
              {activeTab === 'tuition' || activeTab === 'students' ? (
                <>
                  <button onClick={handleExportStudentListExcel} className="px-4 py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold flex items-center gap-1.5 shadow-xs" title="Xuất danh sách học sinh theo alphabet">
                    <Download className="w-4 h-4 text-emerald-600" /> Xuất File Danh Sách Học Sinh
                  </button>
                  <button onClick={() => setExcelChoiceModal(true)} className="px-4 py-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-semibold flex items-center gap-1.5 shadow-xs">
                    <Upload className="w-4 h-4 text-indigo-600" /> Nhập danh sách học sinh
                  </button>
                </>
              ) : null}
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={() => setSellModal({ open: true, productId: products[0]?.id || '', quantity: 1, buyerName: '', note: '' })}
                className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs"
              >
                <ShoppingCart className="w-4 h-4" /> Bán hàng
              </button>
              <button onClick={() => setCourseModal({ open: true, mode: 'add', data: { id: `COURSE_${Date.now()}`, name: '', fee: '', description: '' } })} className="px-3.5 py-2 bg-amber-500 hover:bg-amber-600 text-white rounded-xl text-xs font-semibold flex items-center gap-1">
                <BookPlus className="w-3.5 h-3.5" /> Thêm Khóa Học
              </button>
              <button onClick={() => setStudentModal({ open: true, mode: 'add', data: { id: `HS${String(students.length + 1).padStart(3, '0')}`, name: '', className: courses[0]?.name || '', phone: '', parentName: '', tuitionFee: courses[0]?.fee || 1200000, status: 'active' } })} className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold flex items-center gap-1">
                <UserPlus2 className="w-3.5 h-3.5" /> Thêm Học Sinh
              </button>
            </div>
          </div>

          {/* ====================================================
              TAB: BÁO CÁO DOANH THU + BIỂU ĐỒ 7 NGÀY
          ==================================================== */}
          {activeTab === 'revenue' && (
            <div className="space-y-6">
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div>
                    <h3 className="font-bold text-base text-slate-900 flex items-center gap-2"><BarChart3 className="w-5 h-5 text-indigo-600" /> Biểu Đồ Doanh Thu 7 Ngày Gần Đây</h3>
                    <p className="text-xs text-slate-500">Tổng hợp dòng tiền thu học phí và bán hàng thực tế theo ngày</p>
                  </div>
                  <div className="flex items-center gap-4 text-xs font-semibold">
                    <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-full bg-emerald-500 inline-block"></span> Học phí</span>
                    <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-full bg-indigo-600 inline-block"></span> Bán đồ shop</span>
                  </div>
                </div>

                <div className="pt-6 pb-2 flex items-end justify-between gap-2 sm:gap-4 h-56 border-b border-slate-200 px-2">
                  {last7DaysChartData.days.map((item, idx) => {
                    const totalHeightPct = Math.round((item.total / last7DaysChartData.maxVal) * 100);
                    const tuitionHeightPct = item.total > 0 ? Math.round((item.tuition / item.total) * totalHeightPct) : 0;
                    const shopHeightPct = totalHeightPct - tuitionHeightPct;

                    return (
                      <div key={idx} className="flex-1 flex flex-col items-center h-full justify-end group relative">
                        <div className="absolute -top-12 opacity-0 group-hover:opacity-100 transition bg-slate-900 text-white text-[10px] py-1 px-2 rounded-lg pointer-events-none whitespace-nowrap z-10 shadow-md">
                          <p className="font-bold">{item.label}</p>
                          <p>Tổng: {formatVND(item.total)}</p>
                        </div>

                        <div className="w-full max-w-[36px] bg-slate-100 rounded-t-xl overflow-hidden flex flex-col justify-end h-40">
                          {item.total === 0 ? (
                            <div className="w-full h-0.5 bg-slate-200"></div>
                          ) : (
                            <>
                              <div style={{ height: `${shopHeightPct}%` }} className="w-full bg-indigo-600 transition-all duration-500" title={`Shop: ${formatVND(item.shop)}`}></div>
                              <div style={{ height: `${tuitionHeightPct}%` }} className="w-full bg-emerald-500 transition-all duration-500" title={`Học phí: ${formatVND(item.tuition)}`}></div>
                            </>
                          )}
                        </div>
                        <span className="text-[11px] font-semibold text-slate-600 mt-2">{item.label}</span>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                <div className="p-4 sm:p-5 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-base text-slate-900 flex items-center gap-2"><DollarSign className="w-5 h-5 text-emerald-600" /> Doanh Thu Học Phí (Tháng {selectedMonth})</h3>
                    <p className="text-xs text-slate-500">Quản lý quỹ học phí học viên đóng hàng tháng</p>
                  </div>
                  <button onClick={handleExportTuitionExcel} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs">
                    <FileSpreadsheet className="w-4 h-4" /> Xuất File Doanh Thu
                  </button>
                </div>
                <div className="p-5 grid grid-cols-1 sm:grid-cols-3 gap-4 bg-emerald-50/30">
                  <div className="p-4 bg-white rounded-xl border border-emerald-100">
                    <p className="text-xs text-slate-500">Dự kiến thu</p>
                    <p className="text-lg font-bold text-slate-800 mt-1">{formatVND(stats.expectedRevenue)}</p>
                  </div>
                  <div className="p-4 bg-white rounded-xl border border-emerald-100">
                    <p className="text-xs text-emerald-700 font-medium">Thực thu (Đã đóng)</p>
                    <p className="text-lg font-bold text-emerald-600 mt-1">{formatVND(stats.collectedRevenue)}</p>
                  </div>
                  <div className="p-4 bg-white rounded-xl border border-emerald-100">
                    <p className="text-xs text-rose-700 font-medium">Còn thiếu (Chưa đóng)</p>
                    <p className="text-lg font-bold text-rose-600 mt-1">{formatVND(stats.pendingRevenue)}</p>
                  </div>
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                <div className="p-4 sm:p-5 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <h3 className="font-bold text-base text-slate-900 flex items-center gap-2"><Store className="w-5 h-5 text-indigo-600" /> Doanh Thu Bán Quần Áo & Dụng Cụ (Shop)</h3>
                    <p className="text-xs text-slate-500">Quỹ doanh thu bán đồ tập tách biệt hoàn toàn khỏi học phí</p>
                  </div>
                  <button onClick={handleExportShopExcel} className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs">
                    <FileSpreadsheet className="w-4 h-4" /> Xuất File Doanh Thu Shop Riêng
                  </button>
                </div>

                <div className="p-5 bg-indigo-50/30 flex items-center justify-between border-b border-slate-100 flex-wrap gap-3">
                  <div>
                    <p className="text-xs text-slate-500">Tổng doanh thu bán hàng tích lũy</p>
                    <p className="text-2xl font-extrabold text-indigo-700 mt-1">{formatVND(shopStats.totalShopRevenue)}</p>
                    <p className="text-xs text-slate-500 mt-0.5">Số lượng đã bán: <strong className="text-slate-800">{shopStats.totalItemsSold} sản phẩm</strong></p>
                  </div>
                  <button onClick={() => setSellModal({ open: true, productId: products[0]?.id || '', quantity: 1, buyerName: '', note: '' })} className="px-4 py-2.5 bg-indigo-600 text-white rounded-xl text-xs font-bold shadow-sm">
                    + Bán đồ ngay
                  </button>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                        <th className="p-3.5">Thời Gian</th>
                        <th className="p-3.5">Sản Phẩm</th>
                        <th className="p-3.5">Số Lượng</th>
                        <th className="p-3.5">Đơn Giá</th>
                        <th className="p-3.5">Thành Tiền</th>
                        <th className="p-3.5">Người Mua / Ghi chú</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-sm">
                      {shopOrders.length === 0 ? (
                        <tr><td colSpan="6" className="p-8 text-center text-slate-400">Chưa có giao dịch bán hàng nào.</td></tr>
                      ) : (
                        paginatedShopOrders.map((o) => (
                          <tr key={o.id} className="hover:bg-slate-50">
                            <td className="p-3.5 text-xs text-slate-500 font-medium">{formatDateTime(o.date)}</td>
                            <td className="p-3.5 font-semibold text-slate-900">{o.productName}</td>
                            <td className="p-3.5 font-bold text-slate-700">{o.quantity}</td>
                            <td className="p-3.5">{formatVND(o.price)}</td>
                            <td className="p-3.5 font-bold text-indigo-600">{formatVND(o.totalAmount)}</td>
                            <td className="p-3.5 text-xs text-slate-600">
                              <p className="font-semibold text-slate-800">{o.buyerName}</p>
                              {o.note && <p className="text-[11px] text-slate-400">{o.note}</p>}
                            </td>
                          </tr>
                        ))
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
              TAB: HỌC SINH ĐANG HỌC & THU HỌC PHÍ (CÓ PHÂN TRANG)
          ==================================================== */}
          {(activeTab === 'tuition' || activeTab === 'students') && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-3.5 sm:p-4 border-b border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center gap-3 bg-slate-50">
                <div className="flex items-center gap-2.5 bg-white px-3.5 py-2 rounded-xl border border-slate-200 flex-1">
                  <Search className="w-4 h-4 text-slate-400 shrink-0" />
                  <input type="text" placeholder="Tìm kiếm học sinh theo tên, mã HS, số điện thoại..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full bg-transparent border-none text-sm focus:outline-none" />
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  <Filter className="w-4 h-4 text-slate-500" />
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

                      <div className="flex items-center justify-between gap-2 pt-2 border-t border-slate-100 flex-wrap">
                        {!s.paid ? (
                          <button onClick={() => openPaymentModal(s)} className="flex-1 py-2 bg-emerald-600 active:bg-emerald-700 text-white rounded-xl text-xs font-semibold text-center shadow-xs">
                            Thu học phí
                          </button>
                        ) : (
                          <button onClick={() => handleMarkUnpaid(s)} className="flex-1 py-2 bg-rose-50 active:bg-rose-100 text-rose-600 rounded-xl text-xs font-semibold text-center">
                            Hủy thu
                          </button>
                        )}
                        <button onClick={() => setStudentModal({ open: true, mode: 'edit', data: { ...s } })} className="px-3 py-2 bg-slate-100 text-slate-700 rounded-xl text-xs font-semibold">Sửa</button>
                        <button onClick={() => handleDeactivateStudent(s)} className="px-3 py-2 bg-amber-50 text-amber-700 rounded-xl text-xs font-semibold" title="Nghỉ">Nghỉ</button>
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
                              <button onClick={() => openPaymentModal(s)} className="px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-semibold">Thu tiền</button>
                            ) : (
                              <button onClick={() => handleMarkUnpaid(s)} className="px-3 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Hủy thu</button>
                            )}
                            <button onClick={() => setStudentModal({ open: true, mode: 'edit', data: { ...s } })} className="px-2.5 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-semibold">Sửa</button>
                            <button onClick={() => handleDeactivateStudent(s)} className="px-2.5 py-1.5 bg-amber-50 text-amber-700 hover:bg-amber-100 rounded-lg text-xs font-semibold">Nghỉ</button>
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
              TAB: HỌC SINH ĐÃ NGHỈ
          ==================================================== */}
          {activeTab === 'inactive_students' && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-4 border-b border-slate-200 bg-amber-50/50">
                <h3 className="font-bold text-sm sm:text-base text-slate-900">Danh Sách Học Sinh Đã Nghỉ / Bảo Lưu</h3>
                <p className="text-xs text-slate-500">Dữ liệu bảo lưu an toàn. Có thể cho đi học lại bất cứ lúc nào.</p>
              </div>
              <div className="overflow-x-auto">
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
              TAB: QUẢN LÝ KHO (CÓ PHÂN TRANG)
          ==================================================== */}
          {activeTab === 'inventory' && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between flex-wrap gap-3">
                <div>
                  <h3 className="font-bold text-sm sm:text-base text-slate-900">Quản Lý Kho Quần Áo & Dụng Cụ Võ Thuật</h3>
                  <p className="text-xs text-slate-500">Quản lý số lượng tồn kho áo võ, đai, bảo hộ</p>
                </div>
                <div className="flex items-center gap-2">
                  <button onClick={() => setSellModal({ open: true, productId: products[0]?.id || '', quantity: 1, buyerName: '', note: '' })} className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-xs">
                    <ShoppingCart className="w-4 h-4" /> Bán hàng ngay
                  </button>
                  <button onClick={() => setProductModal({ open: true, mode: 'add', data: { id: `PROD_${Date.now()}`, name: '', price: '', stock: '', category: 'Đồng phục' } })} className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5">
                    <Plus className="w-4 h-4" /> Thêm sản phẩm
                  </button>
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-slate-200 text-xs font-semibold text-slate-500 bg-slate-50">
                      <th className="p-3.5">Mã SP</th>
                      <th className="p-3.5">Tên Sản Phẩm</th>
                      <th className="p-3.5">Danh Mục</th>
                      <th className="p-3.5">Giá Bán</th>
                      <th className="p-3.5">Tồn Kho</th>
                      <th className="p-3.5 text-right">Thao Tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-sm">
                    {products.length === 0 ? (
                      <tr><td colSpan="6" className="p-8 text-center text-slate-400">Chưa có sản phẩm nào trong kho.</td></tr>
                    ) : (
                      paginatedProducts.map((p) => (
                        <tr key={p.id} className="hover:bg-slate-50">
                          <td className="p-3.5 font-mono text-xs font-bold text-slate-500">{p.id}</td>
                          <td className="p-3.5 font-semibold text-slate-900">{p.name}</td>
                          <td className="p-3.5"><span className="px-2 py-0.5 bg-slate-100 rounded text-xs font-medium text-slate-700">{p.category}</span></td>
                          <td className="p-3.5 font-bold text-indigo-600">{formatVND(p.price)}</td>
                          <td className="p-3.5">
                            {Number(p.stock) <= 5 ? (
                              <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200">Còn {p.stock} (Sắp hết)</span>
                            ) : (
                              <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">Còn {p.stock} sản phẩm</span>
                            )}
                          </td>
                          <td className="p-3.5 text-right space-x-1.5">
                            <button onClick={() => setProductModal({ open: true, mode: 'edit', data: { ...p } })} className="px-2.5 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-semibold">Sửa</button>
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
                        <button onClick={() => setCourseModal({ open: true, mode: 'edit', data: { ...c } })} className="px-2.5 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-semibold">Sửa</button>
                        <button onClick={async () => {
                          if (!window.confirm(`Xóa khóa học "${c.name}"?`)) return;
                          setCourses((prev) => prev.filter((item) => item.id !== c.id));
                          if (db) await deleteDoc(doc(db, 'courses', c.id)).catch(console.error);
                          showToast('Đã xóa khóa học', 'warning');
                        }} className="px-2.5 py-1.5 bg-rose-50 text-rose-600 rounded-lg text-xs font-semibold">Xóa</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </main>
      </div>

      {/* ====================================================
          MODALS
      ==================================================== */}
      {excelChoiceModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-xl w-full p-6 shadow-xl border border-slate-200 space-y-5">
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
          <div className="bg-white rounded-2xl max-w-xl w-full p-6 shadow-xl border border-slate-200 space-y-4">
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
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
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
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold mb-4">{productModal.mode === 'add' ? 'Thêm Sản Phẩm Mới' : 'Sửa Sản Phẩm'}</h3>
            <form onSubmit={handleSaveProduct} className="space-y-3.5">
              <input type="text" required placeholder="Tên sản phẩm" value={productModal.data.name} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              <select value={productModal.data.category} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, category: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white">
                <option value="Đồng phục">Đồng phục</option>
                <option value="Đai võ">Đai võ</option>
                <option value="Bảo hộ">Bảo hộ</option>
                <option value="Khác">Khác</option>
              </select>
              <input type="number" required step="10000" placeholder="Giá bán" value={productModal.data.price} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, price: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              <input type="number" required placeholder="Tồn kho" value={productModal.data.stock} onChange={(e) => setProductModal({ ...productModal, data: { ...productModal.data, stock: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold" />
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
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
            <h3 className="text-base font-bold text-slate-900">Xác Nhận Thu Học Phí — {paymentModal.student.name}</h3>
            <form onSubmit={handleSavePaymentModal} className="space-y-4 mt-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Phương thức thanh toán *</label>
                <select value={paymentModal.method} onChange={(e) => setPaymentModal({ ...paymentModal, method: e.target.value })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white font-semibold text-indigo-600">
                  <option value="Chuyển khoản">Chuyển khoản ngân hàng</option>
                  <option value="Tiền mặt">Tiền mặt tại trung tâm</option>
                  <option value="Ví điện tử">Ví MoMo / ZaloPay</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Ngày & Giờ đóng tiền</label>
                <input type="datetime-local" required value={paymentModal.paidDate} onChange={(e) => setPaymentModal({ ...paymentModal, paidDate: e.target.value })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              </div>
              <div className="flex justify-end gap-2 pt-3 border-t">
                <button type="button" onClick={() => setPaymentModal({ ...paymentModal, open: false })} className="px-4 py-2.5 rounded-xl text-xs font-semibold bg-slate-100">Hủy</button>
                <button type="submit" className="px-5 py-2.5 rounded-xl text-xs font-semibold bg-emerald-600 text-white">Xác Nhận Thu</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {courseModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200">
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
            }} className="space-y-3">
              <input type="text" required placeholder="Tên lớp" value={courseModal.data.name} onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              <input type="number" required placeholder="Học phí" value={courseModal.data.fee} onChange={(e) => setCourseModal({ ...courseModal, data: { ...courseModal.data, fee: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setCourseModal({ ...courseModal, open: false })} className="px-4 py-2 text-xs text-slate-600">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs font-semibold rounded-xl">Lưu</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {studentModal.open && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl border border-slate-200 my-auto">
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
                status: 'active'
              };
              setStudents((prev) => mode === 'add' ? [...prev, studentObj] : prev.map((s) => s.id === studentObj.id ? studentObj : s));
              if (db) await setDoc(doc(db, 'students', studentObj.id), studentObj).catch(console.error);
              setStudentModal({ open: false, mode: 'add', data: { id: '', name: '', className: '', phone: '', parentName: '', tuitionFee: '', status: 'active' } });
              showToast('Đã lưu học sinh!');
            }} className="space-y-3.5">
              <input type="text" required placeholder="Mã HS" value={studentModal.data.id} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, id: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm uppercase font-mono" />
              <input type="text" required placeholder="Họ tên học sinh" value={studentModal.data.name} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, name: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              <select value={studentModal.data.className} onChange={(e) => {
                const found = courses.find((c) => c.name === e.target.value);
                setStudentModal({ ...studentModal, data: { ...studentModal.data, className: e.target.value, tuitionFee: found ? found.fee : studentModal.data.tuitionFee } });
              }} className="w-full px-3.5 py-2.5 border rounded-xl text-sm bg-white">
                {courses.map((c) => <option key={c.id} value={c.name}>{c.name} — {formatVND(c.fee)}</option>)}
              </select>
              <input type="number" required placeholder="Học phí" value={studentModal.data.tuitionFee} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, tuitionFee: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm font-bold text-indigo-600" />
              <input type="text" placeholder="Họ tên phụ huynh" value={studentModal.data.parentName || ''} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, parentName: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
              <div>
                <input type="tel" required placeholder="SĐT phụ huynh (Bắt buộc 10 số) *" value={studentModal.data.phone} onChange={(e) => setStudentModal({ ...studentModal, data: { ...studentModal.data, phone: e.target.value } })} className="w-full px-3.5 py-2.5 border rounded-xl text-sm" />
                <p className="text-[10px] text-slate-400 mt-1">Định dạng: Gồm 10 chữ số, bắt đầu bằng số 0</p>
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setStudentModal({ ...studentModal, open: false })} className="px-4 py-2 text-xs text-slate-600">Hủy</button>
                <button type="submit" className="px-4 py-2 bg-indigo-600 text-white text-xs font-semibold rounded-xl">Lưu Hồ Sơ</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}