import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import html2pdf from 'html2pdf.js';
import ReactCrop, { type Crop } from 'react-image-crop';
import 'react-image-crop/dist/ReactCrop.css';
import { Upload, Download, FolderOpen, FileText, FileSpreadsheet, MessageCircle, Menu, X, Clock, Settings, Save, Plus, ArrowRight, Edit, Check, Trash2, ChevronDown, ChevronLeft, Lock, Unlock, Maximize2, Minimize2, ShieldAlert, KeyRound } from 'lucide-react';
import { PWAInstallButton } from './components/PWAInstallButton';
import { OfflineIndicator } from './components/OfflineIndicator';
import { useOnlineStatus } from './hooks/useOnlineStatus';

interface Employee {
  serial: number;
  selected: boolean;
  code: string;
  name: string;
  days: number;
  advance: number;
  basic: number;
  trans: number;
  bonus: number;
  overtime: number;
  insur: number;
  tax: number;
  penalty: number;
  takaful: number;
  net: number;
}


interface Contact {
  id: string;
  name: string;
  phone: string;
}

const predefinedNames = [
  "حسن محمد حسن احمد",
  "محمد عمر محمد كامل",
  "رندا سعيد سيد عبدالخالق",
  "حبيبه شوقى شاكر السيد",
  "بسمه محمد زكريا سالم",
  "محمد احمد عباس محمد",
  "اسماء صالح على محمد",
  "فارس عمرو السيد علي",
  "نجوى محمد صديق فهمي",
  "امنيه اشرف سيد محمد",
  "ملك هيثم حلمي عبد المعطي",
  "تقى محمد عبدالحميد مصطفى",
  "احمد عوض الله جمال عوض الله",
  "سعيد محمد عادل حلمي",
  "منه احمد محمد حسن عبود"
];

const columnAliases: Record<string, string[]> = {
  code: ['كود الموظف', 'الكود الجديد', 'الكود الوظيفي', 'Code', 'ID'],
  name: ['اسم الموظف', 'إسم الموظف', 'الاسم', 'Name'],
  days: ['ايام الحضور', 'أيام الحضور', 'Days'],
  advance: ['سلفة', 'سلفه', 'سلف', 'Advance', 'Loans'],
  basic: ['اجمال الراتب الاساسي', 'اجمالى الراتب الاساسى', 'المرتب الأساسي ', 'Basic'],
  trans: ['بدل الانتقال', 'بدل انتقالات', 'بدل سفر', 'Trans'],
  bonus: ['حوافز', 'حوافز ', 'Incentives'],
  overtime: ['اضافي قيمه', 'اضافى قيمه', 'Overtime'],
  insur: ['تأمينات اجتماعيه', 'تامينات اجتماعية', 'Insurance'],
  tax: ['الضريبه', 'الضريبة ', 'Tax'],
  penalty: ['قيمة الجزاء', 'جزاءات', 'جزاء', 'Penalty', 'Penalties'],
  takaful: ['صندوق التكافل', 'صندوق التكافل الاجتماعى', 'Takaful'],
  net: ['الصافي النهائي', 'الصافى النهائى', 'صافى المرتب', 'Net']
};

function getVal(row: any, aliasKey: string): any {
  for (const alias of columnAliases[aliasKey]) {
    if (row[alias] !== undefined && row[alias] !== null && String(row[alias]).trim() !== '') {
      return row[alias];
    }
  }
  return aliasKey === 'name' || aliasKey === 'code' ? '---' : 0;
}

function num(v: any): number {
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function extractCycleYear(cycle: string): string {
  const normalized = cycle.replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
  const match = normalized.match(/(\d{4})\s*$/);
  return match ? match[1] : new Date().getFullYear().toString();
}

function getTwoPartName(fullName: string): string {
  const parts = String(fullName || '').trim().split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).join(' ') || '---';
}

function generateCycleOptions() {
  const options = [];
  const today = new Date();
  for (let i = -6; i <= 6; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() + i, 1);
    const prevM = new Date(today.getFullYear(), today.getMonth() + i - 1, 1);
    const val = `من 21 ${prevM.toLocaleDateString('ar-EG', { month: 'long' })} إلى 20 ${d.toLocaleDateString('ar-EG', { month: 'long', year: 'numeric' })}`;
    options.push({ value: val, label: val });
  }
  return options;
}

const cycleOptions = generateCycleOptions();

export default function App() {
  const [employeesData, setEmployeesData] = useState<Employee[]>([]);
  const [allSelected, setAllSelected] = useState(false);
  const [selectedCycle, setSelectedCycle] = useState(cycleOptions[6].value);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [expandedYears, setExpandedYears] = useState<Record<string, boolean>>({ [new Date().getFullYear().toString()]: true });
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editFormData, setEditFormData] = useState<Employee | null>(null);
  const [logoImage, setLogoImage] = useState<string | null>(null);
  const [pendingLogo, setPendingLogo] = useState<string | null>(null);
  const [cropImageSrc, setCropImageSrc] = useState<string | null>(null);
  const [crop, setCrop] = useState<Crop>();
  const [completedCrop, setCompletedCrop] = useState<Crop | null>(null);
  const imgRef = useRef<HTMLImageElement>(null);
  const [savedMonths, setSavedMonths] = useState<Record<string, Employee[]>>({});
  const pdfRenderContainerRef = useRef<HTMLDivElement>(null);
  const isOnline = useOnlineStatus();

  const [appPassword, setAppPassword] = useState<string>(() => {
    return localStorage.getItem('morattaba_app_password') || '0000';
  });
  const [autoLockOnOpen, setAutoLockOnOpen] = useState<boolean>(() => {
    return localStorage.getItem('morattaba_auto_lock_on_open') === 'true';
  });
  const [isAppLocked, setIsAppLocked] = useState<boolean>(() => {
    return (
      localStorage.getItem('morattaba_app_locked') === 'true' ||
      localStorage.getItem('morattaba_auto_lock_on_open') === 'true'
    );
  });
  const [unlockPasswordInput, setUnlockPasswordInput] = useState('');
  const [unlockError, setUnlockError] = useState('');
  const [showSettingsPasswordModal, setShowSettingsPasswordModal] = useState(false);
  const [settingsPasswordInput, setSettingsPasswordInput] = useState('');
  const [settingsPasswordError, setSettingsPasswordError] = useState('');
  const [newPasswordInput, setNewPasswordInput] = useState('');
  const [passwordSavedMsg, setPasswordSavedMsg] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isLandscapeForced, setIsLandscapeForced] = useState(true);
  const [isMobilePortrait, setIsMobilePortrait] = useState<boolean>(() => {
    if (typeof window === 'undefined') return false;
    const isMobile =
      /android|iphone|ipad|ipod|mobile/i.test(window.navigator.userAgent) ||
      window.matchMedia('(pointer: coarse)').matches;
    return isMobile && window.innerWidth < window.innerHeight && window.innerWidth < 768;
  });
  const [saveDialog, setSaveDialog] = useState<{
    type: 'pdf' | 'excel';
    employees: Employee[];
    fileName: string;
  } | null>(null);
  const [isSavingFile, setIsSavingFile] = useState(false);
  const [saveDialogError, setSaveDialogError] = useState('');
  const [saveToast, setSaveToast] = useState<string | null>(null);

  React.useEffect(() => {
    const saved = JSON.parse(localStorage.getItem('morattaba_saved_months') || '{}');
    setSavedMonths(saved);
    
    const savedLogo = localStorage.getItem('morattaba_logo');
    if (savedLogo) {
      setLogoImage(savedLogo);
    }

    
    try {
      const encryptedContacts = localStorage.getItem('morattaba_contacts');
      if (encryptedContacts) {
        setContacts(JSON.parse(decodeURIComponent(atob(encryptedContacts))));
      } else {
        setContacts(predefinedNames.map((name, i) => ({ id: String(i), name, phone: '' })));
      }
    } catch (e) {
      setContacts(predefinedNames.map((name, i) => ({ id: String(i), name, phone: '' })));
    }
  }, []);

  React.useEffect(() => {
    const isStandalone =
      window.matchMedia('(display-mode: standalone)').matches ||
      window.matchMedia('(display-mode: fullscreen)').matches ||
      (window.navigator as unknown as { standalone?: boolean }).standalone === true;
    const isMobileDevice =
      /android|iphone|ipad|ipod|mobile/i.test(window.navigator.userAgent) ||
      window.matchMedia('(pointer: coarse)').matches;

    const updateMobilePortrait = () => {
      setIsMobilePortrait(
        isMobileDevice && window.innerWidth < window.innerHeight && window.innerWidth < 768
      );
    };
    updateMobilePortrait();

    const tryLockLandscape = async (allowFullscreen = false) => {
      if (!isLandscapeForced) return;
      if (!isStandalone && !isMobileDevice) return;
      try {
        if (allowFullscreen && !document.fullscreenElement && document.documentElement.requestFullscreen) {
          await document.documentElement.requestFullscreen().catch(() => {});
        }
        const orientation = screen.orientation as unknown as { lock?: (mode: string) => Promise<void> };
        await orientation?.lock?.('landscape');
      } catch {
        // Fallback CSS rotation handles portrait mobile browsers seamlessly
      }
    };

    if (isLandscapeForced) {
      tryLockLandscape(false);
    }

    const handleFirstMobileTouch = () => {
      if (isLandscapeForced && isMobileDevice && window.innerWidth < window.innerHeight) {
        tryLockLandscape(true);
      }
    };

    if (isMobileDevice) {
      window.addEventListener('touchstart', handleFirstMobileTouch, { once: true, passive: true });
      window.addEventListener('pointerdown', handleFirstMobileTouch, { once: true, passive: true });
    }

    const handleFullscreenChange = () => {
      setIsFullscreen(!!document.fullscreenElement);
    };
    document.addEventListener('fullscreenchange', handleFullscreenChange);
    window.addEventListener('resize', updateMobilePortrait);
    window.addEventListener('orientationchange', updateMobilePortrait);
    return () => {
      document.removeEventListener('fullscreenchange', handleFullscreenChange);
      window.removeEventListener('resize', updateMobilePortrait);
      window.removeEventListener('orientationchange', updateMobilePortrait);
      window.removeEventListener('touchstart', handleFirstMobileTouch);
      window.removeEventListener('pointerdown', handleFirstMobileTouch);
    };
  }, [isLandscapeForced]);

  const saveCurrentData = (data: Employee[], cycle: string) => {
    if (data.length === 0) return;
    const saved = JSON.parse(localStorage.getItem('morattaba_saved_months') || '{}');
    saved[cycle] = data;
    localStorage.setItem('morattaba_saved_months', JSON.stringify(saved));
    setSavedMonths(saved);
  };

  const handleSaveAndArchiveToSidebar = () => {
    if (employeesData.length === 0) {
      setSaveToast('يرجى استيراد بيانات الرواتب أولاً ليتم حفظها وترحيلها');
      setTimeout(() => setSaveToast(null), 3500);
      return;
    }
    saveCurrentData(employeesData, selectedCycle);
    const year = extractCycleYear(selectedCycle);
    setExpandedYears(prev => ({ ...prev, [year]: true }));
    setIsSettingsOpen(false);
    setIsSidebarOpen(true);
    showSuccessToast(`تم حفظ وترحيل الراتب أسفل عام ${year} في القائمة الجانبية`);
  };

  const deleteCycle = (cycle: string) => {
    if (window.confirm(`هل أنت متأكد من حذف بيانات "${cycle}"؟`)) {
      const saved = { ...savedMonths };
      delete saved[cycle];
      localStorage.setItem('morattaba_saved_months', JSON.stringify(saved));
      setSavedMonths(saved);
      if (selectedCycle === cycle) {
        setEmployeesData([]);
      }
    }
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const base64 = evt.target?.result as string;
      setCropImageSrc(base64);
    };
    reader.readAsDataURL(file);
  };

  const handleConfirmCrop = () => {
    if (!completedCrop || !imgRef.current) return;
    
    const canvas = document.createElement('canvas');
    const scaleX = imgRef.current.naturalWidth / imgRef.current.width;
    const scaleY = imgRef.current.naturalHeight / imgRef.current.height;
    canvas.width = completedCrop.width;
    canvas.height = completedCrop.height;
    const ctx = canvas.getContext('2d');
    
    if (!ctx) return;
    
    ctx.drawImage(
      imgRef.current,
      completedCrop.x * scaleX,
      completedCrop.y * scaleY,
      completedCrop.width * scaleX,
      completedCrop.height * scaleY,
      0,
      0,
      completedCrop.width,
      completedCrop.height
    );
    
    const base64Image = canvas.toDataURL('image/png');
    setPendingLogo(base64Image);
    setCropImageSrc(null);
    setCrop(undefined);
    setCompletedCrop(null);
  };

  const handleSaveLogo = () => {
    if (pendingLogo) {
      setLogoImage(pendingLogo);
      localStorage.setItem('morattaba_logo', pendingLogo);
      setPendingLogo(null);
    }
  };

  const loadMonth = (cycle: string) => {
    const data = savedMonths[cycle];
    if (data) {
      setEmployeesData(data);
      setSelectedCycle(cycle);
      setAllSelected(data.every(e => e.selected) && data.length > 0);
      setIsSidebarOpen(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rawJson = XLSX.utils.sheet_to_json(worksheet);

      const parsedData: Employee[] = rawJson.map((row: any, index) => ({
        serial: index + 1,
        selected: false,
        code: getVal(row, 'code'),
        name: getVal(row, 'name'),
        days: num(getVal(row, 'days')),
        advance: num(getVal(row, 'advance')),
        basic: num(getVal(row, 'basic')),
        trans: num(getVal(row, 'trans')),
        bonus: num(getVal(row, 'bonus')),
        overtime: num(getVal(row, 'overtime')),
        insur: num(getVal(row, 'insur')),
        tax: num(getVal(row, 'tax')),
        penalty: num(getVal(row, 'penalty')),
        takaful: num(getVal(row, 'takaful')),
        net: num(getVal(row, 'net'))
      }));

      setEmployeesData(parsedData);
      setAllSelected(false);
      saveCurrentData(parsedData, selectedCycle);
    };
    reader.readAsArrayBuffer(file);
    // Reset file input so same file can be uploaded again if needed
    e.target.value = '';
  };

  const toggleSelect = (index: number) => {
    const newData = [...employeesData];
    newData[index].selected = !newData[index].selected;
    setEmployeesData(newData);
    setAllSelected(newData.every(e => e.selected) && newData.length > 0);
  };

  const toggleSelectAll = () => {
    const newSelected = !allSelected;
    setAllSelected(newSelected);
    setEmployeesData(employeesData.map(emp => ({ ...emp, selected: newSelected })));
  };

  const selectedCount = employeesData.filter(e => e.selected).length;
  const totalCount = employeesData.length;
  const hasSelected = selectedCount > 0;

  const buildExcelBlob = (targetEmployees: Employee[]): Blob => {
    const wb = XLSX.utils.book_new();
    const verticalRows: any[] = [];

    targetEmployees.forEach(emp => {
      verticalRows.push({ 'البيان': '=== قسيمة راتب موظف ===', 'القيمة': `${emp.name} (كود: ${emp.code})` });
      verticalRows.push({ 'البيان': 'الكود الوظيفي', 'القيمة': emp.code });
      verticalRows.push({ 'البيان': 'اسم الموظف', 'القيمة': emp.name });
      verticalRows.push({ 'البيان': 'ايام الحضور', 'القيمة': emp.days });
      verticalRows.push({ 'البيان': 'سلفه', 'القيمة': emp.advance });
      verticalRows.push({ 'البيان': 'اجمال الراتب الاساسي', 'القيمة': emp.basic });
      verticalRows.push({ 'البيان': 'بدل الانتقال', 'القيمة': emp.trans });
      verticalRows.push({ 'البيان': 'حوافز', 'القيمة': emp.bonus });
      verticalRows.push({ 'البيان': 'اضافي قيمه', 'القيمة': emp.overtime });
      verticalRows.push({ 'البيان': 'تأمينات اجتماعيه', 'القيمة': emp.insur });
      verticalRows.push({ 'البيان': 'الضريبه', 'القيمة': emp.tax });
      verticalRows.push({ 'البيان': 'قيمة الجزاء', 'القيمة': emp.penalty });
      verticalRows.push({ 'البيان': 'صندوق التكافل', 'القيمة': emp.takaful });
      verticalRows.push({ 'البيان': 'الصافي النهائي', 'القيمة': emp.net });
      verticalRows.push({ 'البيان': '', 'القيمة': '' });
    });

    const ws = XLSX.utils.json_to_sheet(verticalRows);
    ws['!cols'] = [{ wch: 25 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, ws, "Employee_Salary_Slips");
    const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Blob([wbOut], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  };

  const buildPdfBlob = async (targetEmployees: Employee[], fileName: string): Promise<Blob | null> => {
    if (!targetEmployees.length || !pdfRenderContainerRef.current) return null;

    const container = pdfRenderContainerRef.current;
    container.innerHTML = '';
    container.classList.remove('hidden');

    const wrapper = document.createElement('div');
    wrapper.style.fontFamily = "'Tajawal', sans-serif";
    wrapper.style.direction = "rtl";
    wrapper.style.color = "#1e293b";

    targetEmployees.forEach((emp, idx) => {
      const slip = document.createElement('div');
      slip.style.padding = "30px";
      slip.style.backgroundColor = "#ffffff";
      slip.style.border = "2px solid #1f497d";
      slip.style.borderRadius = "12px";
      slip.style.marginBottom = "30px";
      if (idx < targetEmployees.length - 1) {
        slip.style.pageBreakAfter = "always";
      }

      slip.innerHTML = `
        <div style="display: flex; justify-content: center; align-items: center; gap: 20px; border-bottom: 2px solid #1f497d; padding-bottom: 15px; margin-bottom: 25px;">
            <div style="text-align: right;">
                <h1 style="color: #1f497d; font-size: 22px; font-weight: 800; margin: 0 0 5px 0;">تفاصيل الراتب</h1>
                <p style="color: #64748b; font-size: 13px; margin: 0;">${selectedCycle} - كود الموظف: ${emp.code}</p>
            </div>
            ${logoImage ? `<img src="${logoImage}" style="max-height: 60px; max-width: 120px; object-fit: contain;" />` : ''}
        </div>
        <table style="width: 100%; border-collapse: collapse; font-size: 15px;">
            <tbody>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; width: 45%; color: #475569;">اسم الموظف</th><td style="padding: 10px; font-weight: bold; color: #0f172a;">${emp.name}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">الكود الوظيفي</th><td style="padding: 10px;">${emp.code}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">أيام الحضور</th><td style="padding: 10px;">${emp.days} يوم</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">سلفة</th><td style="padding: 10px; color: #e11d48;">${emp.advance.toLocaleString('en-US')}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">إجمال الراتب الأساسي</th><td style="padding: 10px;">${emp.basic.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">بدل الانتقال</th><td style="padding: 10px;">${emp.trans.toLocaleString('en-US')}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">حوافز</th><td style="padding: 10px; color: #059669;">${emp.bonus.toLocaleString('en-US')}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">إضافي قيمة</th><td style="padding: 10px;">${emp.overtime.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">تأمينات اجتماعية</th><td style="padding: 10px;">${emp.insur.toLocaleString('en-US')}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">الضريبة</th><td style="padding: 10px; color: #e11d48;">${emp.tax.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">قيمة الجزاء</th><td style="padding: 10px; color: #e11d48;">${emp.penalty.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
                <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">صندوق التكافل</th><td style="padding: 10px;">${emp.takaful.toLocaleString('en-US')}</td></tr>
                <tr style="background: #eff6ff; border-top: 2px solid #1f497d; border-bottom: 2px solid #1f497d;">
                    <th style="padding: 14px 10px; color: #1e3a8a; font-size: 17px;">الصافي النهائي</th>
                    <td style="padding: 14px 10px; font-weight: 900; color: #1e3a8a; font-size: 18px;">${emp.net.toLocaleString('en-US', {minimumFractionDigits: 2})} جنيه</td>
                </tr>
            </tbody>
        </table>
        <div style="margin-top: 30px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px dashed #cbd5e1; padding-top: 10px;">
            تم الإصدار بواسطة نظام "مرتبا" لإدارة الرواتب AmirLamay
        </div>
      `;
      wrapper.appendChild(slip);
    });

    container.appendChild(wrapper);

    const opt = {
      margin: 15,
      filename: fileName,
      image: { type: 'jpeg' as 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' as 'portrait' }
    };

    try {
      const pdfBlob = await html2pdf().set(opt).from(wrapper).output('blob');
      container.classList.add('hidden');
      container.innerHTML = '';
      return pdfBlob;
    } catch (error) {
      console.error(error);
      container.classList.add('hidden');
      container.innerHTML = '';
      return null;
    }
  };

  const openSaveDialog = (type: 'pdf' | 'excel', customEmployees?: Employee[]) => {
    const targetEmployees = customEmployees || employeesData.filter(e => e.selected);
    if (!targetEmployees.length) return;

    const baseName =
      targetEmployees.length === 1
        ? targetEmployees[0].name.trim()
        : selectedCycle;

    setSaveDialogError('');
    setSaveDialog({
      type,
      employees: targetEmployees,
      fileName: baseName,
    });
  };

  const triggerBrowserDownload = (blob: Blob, finalFileName: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = finalFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const showSuccessToast = (message: string) => {
    setSaveToast(message);
    setTimeout(() => setSaveToast(null), 3500);
  };

  const handleConfirmSaveFile = async (mode: 'picker' | 'download' | 'share') => {
    if (!saveDialog) return;
    const rawName = saveDialog.fileName.trim();
    if (!rawName) {
      setSaveDialogError('يرجى إدخال اسم للملف أولاً');
      return;
    }

    const ext = saveDialog.type === 'pdf' ? '.pdf' : '.xlsx';
    const cleanBase = rawName.replace(/\.(pdf|xlsx|xls)$/i, '').replace(/[\\/:*?"<>|]/g, '_');
    const finalFileName = `${cleanBase}${ext}`;
    const mimeType =
      saveDialog.type === 'pdf'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    setSaveDialogError('');
    setIsSavingFile(true);

    try {
      let fileHandle: any = null;

      if (mode === 'picker' && typeof (window as any).showSaveFilePicker === 'function') {
        try {
          fileHandle = await (window as any).showSaveFilePicker({
            suggestedName: finalFileName,
            types: [
              saveDialog.type === 'pdf'
                ? { description: 'ملف PDF', accept: { 'application/pdf': ['.pdf'] } }
                : {
                    description: 'ملف Excel',
                    accept: {
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                    },
                  },
            ],
          });
        } catch (pickerErr: any) {
          if (pickerErr?.name === 'AbortError') {
            setIsSavingFile(false);
            return;
          }
          // Fallback to download when File System Access API is blocked (e.g., cross-origin iframe)
          fileHandle = null;
        }
      }

      const blob =
        saveDialog.type === 'pdf'
          ? await buildPdfBlob(saveDialog.employees, finalFileName)
          : buildExcelBlob(saveDialog.employees);

      if (!blob) {
        setSaveDialogError('تعذر تجهيز الملف، يرجى المحاولة مرة أخرى');
        setIsSavingFile(false);
        return;
      }

      if (fileHandle) {
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        setSaveDialog(null);
        showSuccessToast(`تم حفظ الملف "${fileHandle.name || finalFileName}" بنجاح`);
      } else if (mode === 'share') {
        const file = new File([blob], finalFileName, { type: mimeType });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          try {
            await navigator.share({
              files: [file],
              title: finalFileName,
            });
            setSaveDialog(null);
            showSuccessToast(`تم فتح خيارات الحفظ والمشاركة للملف "${finalFileName}"`);
          } catch (shareErr: any) {
            if (shareErr?.name === 'AbortError') {
              setIsSavingFile(false);
              return;
            }
            triggerBrowserDownload(blob, finalFileName);
            setSaveDialog(null);
            showSuccessToast(`تم تنزيل الملف "${finalFileName}" بنجاح`);
          }
        } else {
          triggerBrowserDownload(blob, finalFileName);
          setSaveDialog(null);
          showSuccessToast(`تم تنزيل الملف "${finalFileName}" بنجاح`);
        }
      } else {
        triggerBrowserDownload(blob, finalFileName);
        setSaveDialog(null);
        showSuccessToast(`تم حفظ الملف "${finalFileName}" في مجلد التنزيلات`);
      }
    } catch (err: any) {
      if (err?.name !== 'AbortError') {
        console.error(err);
        setSaveDialogError('حدث خطأ أثناء حفظ الملف، يرجى المحاولة مرة أخرى');
      }
    } finally {
      setIsSavingFile(false);
    }
  };

  const exportExcel = () => {
    openSaveDialog('excel');
  };

  const exportPdf = () => {
    openSaveDialog('pdf');
  };

  const generatePdfBlobForEmployee = async (emp: Employee): Promise<Blob | null> => {
    if (!pdfRenderContainerRef.current) return null;

    const container = pdfRenderContainerRef.current;
    container.innerHTML = '';
    container.classList.remove('hidden');

    const wrapper = document.createElement('div');
    wrapper.style.fontFamily = "'Tajawal', sans-serif";
    wrapper.style.direction = "rtl";
    wrapper.style.color = "#1e293b";

    const slip = document.createElement('div');
    slip.style.padding = "30px";
    slip.style.backgroundColor = "#ffffff";
    slip.style.border = "2px solid #1f497d";
    slip.style.borderRadius = "12px";
    
    slip.innerHTML = `
      <div style="display: flex; justify-content: center; align-items: center; gap: 20px; border-bottom: 2px solid #1f497d; padding-bottom: 15px; margin-bottom: 25px;">
          <div style="text-align: right;">
              <h1 style="color: #1f497d; font-size: 22px; font-weight: 800; margin: 0 0 5px 0;">تفاصيل الراتب</h1>
              <p style="color: #64748b; font-size: 13px; margin: 0;">${selectedCycle} - كود الموظف: ${emp.code}</p>
          </div>
          ${logoImage ? `<img src="${logoImage}" style="max-height: 60px; max-width: 120px; object-fit: contain;" />` : ''}
      </div>
      <table style="width: 100%; border-collapse: collapse; font-size: 15px;">
          <tbody>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; width: 45%; color: #475569;">اسم الموظف</th><td style="padding: 10px; font-weight: bold; color: #0f172a;">${emp.name}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">الكود الوظيفي</th><td style="padding: 10px;">${emp.code}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">أيام الحضور</th><td style="padding: 10px;">${emp.days} يوم</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">سلفة</th><td style="padding: 10px; color: #e11d48;">${emp.advance.toLocaleString('en-US')}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">إجمال الراتب الأساسي</th><td style="padding: 10px;">${emp.basic.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">بدل الانتقال</th><td style="padding: 10px;">${emp.trans.toLocaleString('en-US')}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">حوافز</th><td style="padding: 10px; color: #059669;">${emp.bonus.toLocaleString('en-US')}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">إضافي قيمة</th><td style="padding: 10px;">${emp.overtime.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">تأمينات اجتماعية</th><td style="padding: 10px;">${emp.insur.toLocaleString('en-US')}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">الضريبة</th><td style="padding: 10px; color: #e11d48;">${emp.tax.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">قيمة الجزاء</th><td style="padding: 10px; color: #e11d48;">${emp.penalty.toLocaleString('en-US', {minimumFractionDigits: 2})}</td></tr>
              <tr style="border-bottom: 1px solid #f1f5f9;"><th style="padding: 10px; background: #f8fafc; text-align: right; color: #475569;">صندوق التكافل</th><td style="padding: 10px;">${emp.takaful.toLocaleString('en-US')}</td></tr>
              <tr style="background: #eff6ff; border-top: 2px solid #1f497d; border-bottom: 2px solid #1f497d;">
                  <th style="padding: 14px 10px; color: #1e3a8a; font-size: 17px;">الصافي النهائي</th>
                  <td style="padding: 14px 10px; font-weight: 900; color: #1e3a8a; font-size: 18px;">${emp.net.toLocaleString('en-US', {minimumFractionDigits: 2})} جنيه</td>
              </tr>
          </tbody>
      </table>
      <div style="margin-top: 30px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px dashed #cbd5e1; padding-top: 10px;">
          تم الإصدار بواسطة نظام "مرتبا" لإدارة الرواتب AmirLamay
      </div>
    `;
    wrapper.appendChild(slip);
    container.appendChild(wrapper);

    const opt = {
      margin: 15,
      filename: `Salary_${emp.name}.pdf`,
      image: { type: 'jpeg' as 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' as 'portrait' }
    };

    try {
      const pdfBlob = await html2pdf().set(opt).from(wrapper).output('blob');
      container.classList.add('hidden');
      container.innerHTML = '';
      return pdfBlob;
    } catch (error) {
      console.error(error);
      container.classList.add('hidden');
      container.innerHTML = '';
      return null;
    }
  };


  const saveContacts = (newContacts: Contact[]) => {
    setContacts(newContacts);
    const encrypted = btoa(encodeURIComponent(JSON.stringify(newContacts)));
    localStorage.setItem('morattaba_contacts', encrypted);
  };

  const handleSettingsClick = () => {
    setSettingsPasswordInput('');
    setSettingsPasswordError('');
    setShowSettingsPasswordModal(true);
  };

  const handleSettingsPasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (settingsPasswordInput === appPassword) {
      setShowSettingsPasswordModal(false);
      setIsSettingsOpen(true);
      setIsSidebarOpen(true);
    } else {
      setSettingsPasswordError('كلمة المرور خاطئة');
    }
  };

  const lockAppNow = () => {
    localStorage.setItem('morattaba_app_locked', 'true');
    setIsAppLocked(true);
    setIsSettingsOpen(false);
    setIsSidebarOpen(false);
    setUnlockPasswordInput('');
    setUnlockError('');
  };

  const handleUnlockApp = (e: React.FormEvent) => {
    e.preventDefault();
    if (unlockPasswordInput === appPassword) {
      localStorage.setItem('morattaba_app_locked', 'false');
      setIsAppLocked(false);
      setUnlockPasswordInput('');
      setUnlockError('');
    } else {
      setUnlockError('كلمة المرور غير صحيحة');
    }
  };

  const toggleAutoLockOnOpen = () => {
    const next = !autoLockOnOpen;
    setAutoLockOnOpen(next);
    localStorage.setItem('morattaba_auto_lock_on_open', String(next));
  };

  const handleSaveNewPassword = () => {
    const trimmed = newPasswordInput.trim();
    if (!trimmed) return;
    setAppPassword(trimmed);
    localStorage.setItem('morattaba_app_password', trimmed);
    setNewPasswordInput('');
    setPasswordSavedMsg(true);
    setTimeout(() => setPasswordSavedMsg(false), 2500);
  };

  const toggleLandscapeFullscreen = async () => {
    const currentlyActive = isFullscreen || (isMobilePortrait && isLandscapeForced);
    if (currentlyActive) {
      setIsLandscapeForced(false);
      try {
        const orientation = screen.orientation as unknown as { unlock?: () => void };
        orientation?.unlock?.();
        if (document.fullscreenElement) {
          await document.exitFullscreen();
        }
      } catch {
        // Ignore if unlock/exitFullscreen not active
      }
    } else {
      setIsLandscapeForced(true);
      try {
        if (!document.fullscreenElement && document.documentElement.requestFullscreen) {
          await document.documentElement.requestFullscreen();
        }
        const orientation = screen.orientation as unknown as { lock?: (mode: string) => Promise<void> };
        await orientation?.lock?.('landscape');
      } catch {
        // Fallback CSS rotation handles portrait mobile browsers seamlessly
      }
    }
  };

  const handleShareSelected = async () => {
    const selected = employeesData.filter(e => e.selected);
    for (const emp of selected) {
      await shareEmployeePdf(emp);
    }
  };

  const handleEditSelected = () => {
    const idx = employeesData.findIndex(e => e.selected);
    if (idx !== -1) {
      handleEditClick(idx, employeesData[idx]);
    }
  };

  const handleDeleteSelected = () => {
    if (window.confirm('هل تريد بالتأكيد حذف السجلات المحددة؟ (نعم للحذف، لا للإلغاء)')) {
      const remaining = employeesData.filter(emp => !emp.selected).map((emp, i) => ({ ...emp, serial: i + 1 }));
      setEmployeesData(remaining);
      saveCurrentData(remaining, selectedCycle);
      setAllSelected(false);
    }
  };

  const handleEditClick = (idx: number, emp: Employee) => {
    setEditingIndex(idx);
    setEditFormData({ ...emp });
  };

  const handleEditChange = (field: keyof Employee, value: string) => {
    if (!editFormData) return;
    const numVal = parseFloat(value) || 0;
    const updated = { ...editFormData, [field]: field === 'name' || field === 'code' ? value : numVal };
    
    // Recalculate net if numeric field
    if (['basic', 'trans', 'bonus', 'overtime', 'advance', 'insur', 'tax', 'penalty', 'takaful'].includes(field)) {
      updated.net = updated.basic + updated.trans + updated.bonus + updated.overtime - updated.advance - updated.insur - updated.tax - updated.penalty - updated.takaful;
    }
    
    setEditFormData(updated);
  };

  const handleEditSave = (idx: number) => {
    if (!editFormData) return;
    const newData = [...employeesData];
    newData[idx] = editFormData;
    setEmployeesData(newData);
    saveCurrentData(newData, selectedCycle);
    setEditingIndex(null);
    setEditFormData(null);
  };

  const handleEditCancel = () => {
    setEditingIndex(null);
    setEditFormData(null);
  };

  const shareEmployeePdf = async (emp: Employee) => {
    const contact = contacts.find(c => c.name.trim() === emp.name.trim());
    if (contact && contact.phone) {
      const pdfBlob = await generatePdfBlobForEmployee(emp);
      if (!pdfBlob) return;
      
      const file = new File([pdfBlob], `Salary_${emp.name}.pdf`, { type: 'application/pdf' });
      
      // Auto-download file for WhatsApp attachment later
      const url = URL.createObjectURL(pdfBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Salary_${emp.name}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
      
      const text = encodeURIComponent(`مرحباً ${emp.name}، مرفق قسيمة تفاصيل الراتب الخاصة بك.`);
      const waLink = `https://wa.me/${contact.phone}?text=${text}`;
      window.open(waLink, '_blank');
      return;
    }

    const pdfBlob = await generatePdfBlobForEmployee(emp);
    if (!pdfBlob) return;

    const file = new File([pdfBlob], `Salary_${emp.name}.pdf`, { type: 'application/pdf' });
    
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: `تفاصيل الراتب - ${emp.name}`,
        });
      } catch (err) {
        console.error('Error sharing:', err);
      }
    } else {
      // Fallback: Open Save File dialog so user can choose name & save location
      openSaveDialog('pdf', [emp]);
    }
  };

  const handleCycleChange = (newCycle: string) => {
    setSelectedCycle(newCycle);
    if (employeesData.length > 0) {
      saveCurrentData(employeesData, newCycle);
    }
  };

  const totalNet = employeesData.reduce((sum, emp) => sum + emp.net, 0);

  if (isAppLocked) {
    return (
      <div className="min-h-screen w-full bg-indigo-950 text-white flex items-center justify-center p-4 font-sans" dir="rtl">
        <div className="w-full max-w-md bg-white text-slate-800 rounded-2xl shadow-2xl border border-indigo-200 p-6 sm:p-8 text-center space-y-5">
          <div className="mx-auto w-16 h-16 bg-rose-100 text-rose-600 rounded-2xl flex items-center justify-center shadow-inner">
            <ShieldAlert className="w-9 h-9" />
          </div>
          <div className="space-y-2">
            <div className="inline-flex items-center gap-2 bg-indigo-50 text-indigo-800 px-3 py-1 rounded-full text-xs font-bold">
              <span>مُرَتَّبَا | Morattaba</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black text-slate-900">التطبيق متوقف ومقفل حالياً</h1>
            <p className="text-sm text-slate-600 leading-relaxed">
              هذا التطبيق لا يعمل نهائياً في الوقت الحالي ولا يمكن الوصول إلى أي بيانات إلا إذا أراد المستخدم تشغيله.
            </p>
          </div>

          <form onSubmit={handleUnlockApp} className="space-y-3 text-right">
            <label className="block text-xs font-bold text-slate-700">
              أدخل كلمة المرور لتشغيل التطبيق (الافتراضية: 0000)
            </label>
            <div className="relative">
              <KeyRound className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="password"
                value={unlockPasswordInput}
                onChange={(e) => {
                  setUnlockPasswordInput(e.target.value);
                  setUnlockError('');
                }}
                placeholder="أدخل كلمة المرور..."
                className="w-full pr-9 pl-3 py-2.5 rounded-lg border border-slate-300 text-sm outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                autoFocus
              />
            </div>
            {unlockError && (
              <p className="text-xs font-bold text-rose-600">{unlockError}</p>
            )}
            <button
              type="submit"
              className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 rounded-lg font-bold text-sm transition-colors flex items-center justify-center gap-2 shadow-sm cursor-pointer"
            >
              <Unlock className="w-4 h-4" />
              <span>تشغيل التطبيق الآن</span>
            </button>
          </form>

          <div className="pt-2 border-t border-slate-100 flex justify-center">
            <PWAInstallButton />
          </div>
        </div>
        <OfflineIndicator />
      </div>
    );
  }

  const isLandscapeActive = isFullscreen || (isMobilePortrait && isLandscapeForced);
  const rotatedMobileClasses =
    isMobilePortrait && isLandscapeForced
      ? 'fixed top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 rotate-90 w-[100dvh] h-[100vw]'
      : 'w-full h-dvh';

  return (
    <div
      className={`flex flex-col ${rotatedMobileClasses} bg-slate-50 text-slate-800 font-sans overflow-y-auto overflow-x-hidden [scrollbar-width:none] [&::-webkit-scrollbar]:hidden md:overflow-hidden`}
      dir="rtl"
    >
      {/* Header */}
      <header className="bg-indigo-950 text-white px-2 sm:px-4 py-1.5 sm:py-2.5 flex justify-between items-center shrink-0 border-b-2 sm:border-b-4 border-indigo-500 gap-1.5">
        <div className="flex items-center gap-1.5 sm:gap-3">
          <button onClick={() => setIsSidebarOpen(true)} className="p-1.5 hover:bg-indigo-800 rounded-lg transition-colors shrink-0">
            <Menu className="w-5 h-5" />
          </button>
          <div className="w-7 h-7 sm:w-8 sm:h-8 bg-indigo-500 rounded-lg flex items-center justify-center text-base sm:text-lg font-black shrink-0">M</div>
          <div className="hidden sm:block">
            <h1 className="text-sm sm:text-base font-bold leading-tight">مُرَتَّبَا | <span className="text-indigo-300 font-medium">Morattaba</span></h1>
            <p className="text-[10px] text-indigo-200 opacity-80">نظام إدارة وتصدير قسائم رواتب الموظفين v2.4</p>
          </div>
        </div>
        <div className="flex gap-1 sm:gap-1.5 items-center">
          <div className="hidden lg:flex items-center gap-1.5 bg-slate-800/50 px-2.5 py-1 rounded-md border border-slate-700">
            <span className={`w-2 h-2 rounded-full ${isOnline ? 'bg-emerald-500' : 'bg-amber-500'}`}></span>
            <span className="text-[11px] font-medium">{isOnline ? 'المتصفح: متصل' : 'وضع عدم الاتصال'}</span>
          </div>
          <PWAInstallButton />
          <button
            onClick={toggleLandscapeFullscreen}
            className="flex items-center gap-1 bg-indigo-800/80 hover:bg-indigo-700 border border-indigo-400/40 px-2 py-1.5 rounded text-[11px] font-bold text-indigo-100 transition-colors cursor-pointer"
            title={isLandscapeActive ? 'إلغاء عرض الشاشة' : 'فتح بعرض الشاشة (الوضع الأفقي)'}
          >
            {isLandscapeActive ? <Minimize2 className="w-3.5 h-3.5 shrink-0" /> : <Maximize2 className="w-3.5 h-3.5 shrink-0" />}
            <span className="hidden md:inline">{isLandscapeActive ? 'إنهاء العرض' : 'عرض الشاشة'}</span>
          </button>
          <button
            onClick={lockAppNow}
            className="flex items-center gap-1 bg-rose-600/90 hover:bg-rose-700 px-2 py-1.5 rounded text-[11px] font-bold text-white transition-colors cursor-pointer"
            title="إيقاف وقفل التطبيق"
          >
            <Lock className="w-3.5 h-3.5 shrink-0" />
            <span className="hidden md:inline">قفل التطبيق</span>
          </button>
          <button
            onClick={handleSaveAndArchiveToSidebar}
            className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 px-2 sm:px-3 py-1.5 rounded text-[11px] sm:text-xs font-bold text-white shadow-sm transition-colors cursor-pointer"
            title="حفظ وترحيل الراتب في القائمة الجانبية أسفل العام"
          >
            <Save className="w-3.5 h-3.5 shrink-0" />
            <span>حفظ</span>
          </button>
          <label className="cursor-pointer bg-indigo-600 hover:bg-indigo-700 px-2 sm:px-3 py-1.5 rounded text-[11px] sm:text-xs font-bold shadow-sm transition-colors flex items-center gap-1.5">
             <Upload className="w-3.5 h-3.5" />
             <span className="hidden sm:inline">استيراد ملف Excel</span>
             <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleFileUpload} />
          </label>
        </div>
      </header>

      {/* Sidebar Overlay */}
      {isSidebarOpen && (
        <div 
          className="fixed inset-0 bg-slate-900/50 z-50 transition-opacity"
          onClick={() => setIsSidebarOpen(false)}
        ></div>
      )}

      
      {/* Sidebar */}
      <div className={`fixed top-0 right-0 h-full w-80 sm:w-96 bg-white shadow-2xl z-50 transform transition-transform duration-300 ease-in-out ${isSidebarOpen ? 'translate-x-0' : 'translate-x-full'}`}>
        {isSettingsOpen ? (
          <>
            <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <div className="flex items-center gap-3">
                <button onClick={() => setIsSettingsOpen(false)} className="p-1.5 hover:bg-slate-200 rounded-full transition-colors">
                  <ArrowRight className="w-5 h-5 text-slate-600" />
                </button>
                <h2 className="font-bold text-lg text-slate-800">الإعدادات</h2>
              </div>
              <button onClick={() => { setIsSettingsOpen(false); setIsSidebarOpen(false); }} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>
            <div className="p-4 overflow-y-auto h-[calc(100%-60px)] bg-slate-50">
              <div className="space-y-4">
                <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-sm space-y-3">
                  <h3 className="font-bold text-slate-700 text-sm flex items-center gap-2">
                    <Lock className="w-4 h-4 text-indigo-600" />
                    حماية وتشغيل التطبيق
                  </h3>
                  <label className="flex items-start gap-2.5 cursor-pointer bg-slate-50 p-2.5 rounded-lg border border-slate-200">
                    <input
                      type="checkbox"
                      checked={autoLockOnOpen}
                      onChange={toggleAutoLockOnOpen}
                      className="mt-0.5 accent-indigo-600 w-4 h-4 cursor-pointer shrink-0"
                    />
                    <div>
                      <div className="text-xs font-bold text-slate-800">قفل التطبيق تلقائياً عند كل فتح</div>
                      <div className="text-[11px] text-slate-500 mt-0.5">لا يعمل التطبيق نهائياً عند فتحه إلا بعد إدخال كلمة المرور برغبة المستخدم.</div>
                    </div>
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="password"
                      value={newPasswordInput}
                      onChange={(e) => setNewPasswordInput(e.target.value)}
                      placeholder="تغيير كلمة المرور..."
                      className="flex-1 text-xs bg-slate-50 px-2.5 py-2 rounded border border-slate-200 outline-none focus:border-indigo-400"
                    />
                    <button
                      onClick={handleSaveNewPassword}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded text-xs font-bold transition-colors shrink-0"
                    >
                      حفظ
                    </button>
                  </div>
                  {passwordSavedMsg && (
                    <p className="text-xs font-bold text-emerald-600">تم تحديث كلمة المرور بنجاح</p>
                  )}
                  <button
                    onClick={lockAppNow}
                    className="w-full bg-rose-600 hover:bg-rose-700 text-white py-2 rounded-lg text-xs font-bold transition-colors flex items-center justify-center gap-2"
                  >
                    <Lock className="w-4 h-4" />
                    إيقاف وقفل التطبيق الآن
                  </button>
                </div>

                <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-sm flex flex-col items-center justify-center gap-3">
                  <h3 className="font-bold text-slate-700 text-sm w-full text-right">شعار الشركة (اللوجو)</h3>
                  <label className="cursor-pointer group relative flex flex-col items-center justify-center w-32 h-32 rounded-lg border-2 border-dashed border-slate-300 hover:border-indigo-400 bg-slate-50 overflow-hidden transition-colors">
                    <input type="file" accept="image/*" className="hidden" onChange={handleLogoUpload} />
                    {pendingLogo || logoImage ? (
                      <>
                        <img src={(pendingLogo || logoImage) as string} alt="Logo" className="w-full h-full object-contain" />
                        <div className="absolute inset-0 bg-black/40 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity">
                          <span className="text-white text-xs font-bold">تغيير الصورة</span>
                        </div>
                      </>
                    ) : (
                      <>
                        <Upload className="w-8 h-8 text-slate-400 mb-2 group-hover:text-indigo-400 transition-colors" />
                        <span className="text-xs text-slate-500 font-medium">اختر صورة</span>
                      </>
                    )}
                  </label>
                  {pendingLogo && (
                    <button
                      onClick={handleSaveLogo}
                      className="mt-2 w-full bg-emerald-600 hover:bg-emerald-700 text-white py-2 rounded-lg text-sm font-bold transition-colors flex items-center justify-center gap-2"
                    >
                      <Save className="w-4 h-4" /> حفظ الصورة
                    </button>
                  )}
                </div>
                
                <div className="pt-2">
                  <h3 className="font-bold text-slate-700 text-sm mb-3 text-right">أرقام الموظفين</h3>
                </div>
                {contacts.map((contact, idx) => (
                  <div key={contact.id} className="bg-white p-3 rounded-lg border border-slate-200 shadow-sm flex flex-col gap-2">
                    <input 
                      type="text" 
                      value={contact.name}
                      onChange={(e) => {
                        const newContacts = [...contacts];
                        newContacts[idx].name = e.target.value;
                        saveContacts(newContacts);
                      }}
                      placeholder="اسم الموظف"
                      className="text-sm font-bold bg-transparent border-none outline-none w-full text-slate-800"
                    />
                    <input 
                      type="text" 
                      value={contact.phone}
                      onChange={(e) => {
                        const newContacts = [...contacts];
                        newContacts[idx].phone = e.target.value;
                        saveContacts(newContacts);
                      }}
                      placeholder="رقم الواتساب (مثال: 201xxxxxxxxx)"
                      className="text-xs bg-slate-100 p-2 rounded border border-slate-200 outline-none focus:border-indigo-400 w-full text-slate-700"
                      dir="ltr"
                    />
                  </div>
                ))}
                
                <button
                  onClick={() => {
                    const newContacts = [...contacts, { id: Date.now().toString(), name: 'موظف جديد', phone: '' }];
                    saveContacts(newContacts);
                  }}
                  className="w-full py-3 border-2 border-dashed border-indigo-300 text-indigo-600 rounded-lg font-bold text-sm hover:bg-indigo-50 transition-colors flex items-center justify-center gap-2"
                >
                  <Plus className="w-4 h-4" /> إضافة اسم جديد
                </button>
              </div>
            </div>
          </>
        ) : (
          <>
            <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
              <h2 className="font-bold text-lg text-slate-800">مرتبات الأشهر</h2>
              <div className="flex items-center gap-1">
                <button onClick={handleSettingsClick} className="p-2 hover:bg-slate-200 rounded-full transition-colors" title="الإعدادات">
                  <Settings className="w-5 h-5 text-slate-500" />
                </button>
                <button onClick={() => setIsSidebarOpen(false)} className="p-2 hover:bg-slate-200 rounded-full transition-colors" title="إغلاق">
                  <X className="w-5 h-5 text-slate-500" />
                </button>
              </div>
            </div>
            <div className="p-4 overflow-y-auto h-[calc(100%-60px)]">
              {Object.keys(savedMonths).length === 0 ? (
                <p className="text-slate-500 text-sm text-center mt-10">لا توجد سجلات محفوظة حالياً</p>
              ) : (
                <div className="space-y-4">
                  {Object.entries(
                    Object.keys(savedMonths).reduce((acc: Record<string, string[]>, cycle) => {
                      const year = extractCycleYear(cycle);
                      if (!acc[year]) acc[year] = [];
                      acc[year].push(cycle);
                      return acc;
                    }, {})
                  )
                  .sort(([yearA], [yearB]) => parseInt(yearB) - parseInt(yearA)) // Descending year
                  .map(([year, cycles]) => (
                    <div key={year} className="bg-white rounded-lg border border-slate-200 overflow-hidden shadow-sm">
                      <button
                        onClick={() => setExpandedYears(prev => ({ ...prev, [year]: !prev[year] }))}
                        className="w-full flex items-center justify-between p-3 bg-slate-100 hover:bg-slate-200 transition-colors"
                      >
                        <h3 className="font-bold text-slate-800">عام {year}</h3>
                        {expandedYears[year] ? <ChevronDown className="w-5 h-5 text-slate-500" /> : <ChevronLeft className="w-5 h-5 text-slate-500" />}
                      </button>
                      {expandedYears[year] && (
                        <div className="p-2 space-y-2 bg-slate-50">
                          {cycles.map((cycle) => (
                            <div key={cycle} className="relative group w-full flex rounded-lg border border-slate-200 hover:border-indigo-300 bg-white hover:bg-indigo-50 transition-colors overflow-hidden">
                              <button
                                onClick={() => loadMonth(cycle)}
                                className="flex-1 text-right p-3 flex items-center gap-3"
                              >
                                <Clock className="w-5 h-5 text-indigo-500 shrink-0" />
                                <div>
                                  <div className="font-bold text-slate-700 text-sm">{cycle}</div>
                                  <div className="text-xs text-slate-500 mt-1">{savedMonths[cycle].length} موظف</div>
                                </div>
                              </button>
                              <button 
                                onClick={(e) => { e.stopPropagation(); deleteCycle(cycle); }}
                                className="p-3 text-rose-500 hover:bg-rose-100 transition-colors flex items-center justify-center shrink-0"
                                title="حذف السجل"
                              >
                                <Trash2 className="w-5 h-5" />
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </>
        )}
      </div>


      {/* Main Content Area */}
      <div className={`${employeesData.length === 0 ? 'flex-1' : 'md:flex-1 md:min-h-0 md:overflow-hidden'} mx-0 sm:mx-3 mb-1 sm:mb-2 bg-white rounded-none sm:rounded-xl border-y sm:border border-slate-200 shadow-sm flex flex-col shrink-0`}>
        <div className="flex flex-wrap items-center justify-between p-1.5 sm:p-2 border-b bg-slate-50 gap-1.5 shrink-0">
          <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
             <h2 className="font-bold text-xs sm:text-sm text-slate-700">سجل البيانات الحالية</h2>
             {employeesData.length > 0 && (
                <button
                   onClick={toggleSelectAll}
                   className="bg-slate-200 hover:bg-slate-300 text-slate-700 px-2 py-1 rounded text-[10px] sm:text-xs font-bold transition-colors"
                 >
                   {allSelected ? 'إلغاء تحديد الكل' : 'تحديد جميع اللاينات'}
                 </button>
             )}
             {hasSelected && (
                <>
                  <button
                     onClick={handleShareSelected}
                     className="bg-emerald-100 hover:bg-emerald-200 text-emerald-700 px-2 py-1 rounded text-[10px] sm:text-xs font-bold transition-colors flex items-center gap-1"
                     title="مشاركة عبر واتساب"
                   >
                     <MessageCircle className="w-3.5 h-3.5" /> مشاركة
                   </button>
                  <button
                     onClick={handleEditSelected}
                     disabled={selectedCount !== 1}
                     className="bg-indigo-100 hover:bg-indigo-200 disabled:bg-slate-100 disabled:text-slate-400 text-indigo-700 px-2 py-1 rounded text-[10px] sm:text-xs font-bold transition-colors disabled:cursor-not-allowed flex items-center gap-1"
                     title={selectedCount === 1 ? 'تعديل الموظف المحدد' : 'حدد موظفاً واحداً للتعديل'}
                   >
                     <Edit className="w-3.5 h-3.5" /> تعديل
                   </button>
                  <button
                     onClick={handleDeleteSelected}
                     className="bg-rose-100 hover:bg-rose-200 text-rose-700 px-2 py-1 rounded text-[10px] sm:text-xs font-bold transition-colors flex items-center gap-1"
                   >
                     <Trash2 className="w-3.5 h-3.5" /> حذف المحددة
                   </button>
                </>
             )}
          </div>
          <div className="flex flex-wrap items-center gap-1.5 w-full sm:w-auto">
            <select
              value={selectedCycle}
              onChange={(e) => handleCycleChange(e.target.value)}
              className="bg-white border border-slate-300 text-slate-700 px-2 py-1 rounded text-[10px] sm:text-xs font-bold outline-none focus:ring-2 focus:ring-indigo-500 flex-1 sm:flex-none"
            >
              {cycleOptions.map((opt, i) => (
                <option key={i} value={opt.value}>{opt.label}</option>
              ))}
            </select>
            <div className="flex items-center gap-1.5 w-full sm:w-auto">
              <button
                onClick={exportExcel}
                disabled={!hasSelected}
                className="flex-1 sm:flex-none justify-center bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-200 disabled:text-slate-400 text-white px-2.5 py-1 rounded text-[10px] sm:text-xs font-bold transition-colors disabled:cursor-not-allowed flex items-center gap-1"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" /> تصدير Excel
              </button>
              <button
                onClick={exportPdf}
                disabled={!hasSelected}
                className="flex-1 sm:flex-none justify-center bg-rose-600 hover:bg-rose-700 disabled:bg-slate-200 disabled:text-slate-400 text-white px-2.5 py-1 rounded text-[10px] sm:text-xs font-bold transition-colors disabled:cursor-not-allowed flex items-center gap-1"
              >
                <FileText className="w-3.5 h-3.5 shrink-0" /> تصدير PDF
              </button>
            </div>
          </div>
        </div>
        <div className="md:flex-1 md:min-h-0 md:overflow-y-auto overflow-x-hidden [scrollbar-width:none] [&::-webkit-scrollbar]:hidden relative">
           {employeesData.length > 0 ? (
              <table className="w-full table-auto text-[8px] sm:text-[10px] lg:text-[11px] text-right border-collapse">
                  <thead className="bg-slate-900 text-white sticky top-0 z-10 whitespace-nowrap">
                    <tr>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 w-5 sm:w-7 text-center">
                        <input
                          type="checkbox"
                          checked={allSelected}
                          onChange={toggleSelectAll}
                          className="accent-indigo-500 w-2.5 h-2.5 sm:w-3 sm:h-3 cursor-pointer"
                        />
                      </th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center w-5 sm:w-7">م</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">الكود</th>
                      <th className="py-1.5 px-1 sm:px-1.5 border-l border-slate-700">اسم الموظف</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">الحضور</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">سلفة</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">الأساسي</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">انتقالات</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">حوافز</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">إضافي</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">تأمينات</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">ضرائب</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">جزاءات</th>
                      <th className="py-1.5 px-0.5 sm:px-1 border-l border-slate-700 text-center">تكافل</th>
                      <th className="py-1.5 px-0.5 sm:px-1.5 bg-indigo-600 text-center font-bold">الصافي</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {employeesData.map((emp, idx) => (
                      editingIndex === idx && editFormData ? (
                        <tr key={idx} className="bg-indigo-50 transition-colors">
                          <td className="p-0.5 sm:p-1 text-center">
                             <div className="flex flex-col items-center gap-0.5">
                               <button onClick={() => handleEditSave(idx)} className="text-emerald-600 bg-emerald-100 p-0.5 sm:p-1 rounded-full hover:bg-emerald-200">
                                 <Check className="w-3 h-3 sm:w-3.5 sm:h-3.5" />
                               </button>
                               <button onClick={handleEditCancel} className="text-rose-600 bg-rose-100 p-0.5 sm:p-1 rounded-full hover:bg-rose-200">
                                 <X className="w-3 h-3 sm:w-3.5 sm:h-3.5" />
                               </button>
                             </div>
                          </td>
                          <td className="p-0.5 sm:p-1 text-center text-slate-500">{emp.serial}</td>
                          <td className="p-0.5 sm:p-1 text-center font-mono font-bold text-indigo-700">{emp.code}</td>
                          <td className="p-0.5 sm:p-1 font-bold whitespace-nowrap" title={emp.name}>{getTwoPartName(emp.name)}</td>
                          <td className="p-0.5"><input type="number" value={editFormData.days} onChange={e => handleEditChange('days', e.target.value)} className="w-8 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px]" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.advance} onChange={e => handleEditChange('advance', e.target.value)} className="w-10 sm:w-14 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px] text-rose-600" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.basic} onChange={e => handleEditChange('basic', e.target.value)} className="w-11 sm:w-16 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px]" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.trans} onChange={e => handleEditChange('trans', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px]" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.bonus} onChange={e => handleEditChange('bonus', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px] text-emerald-600" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.overtime} onChange={e => handleEditChange('overtime', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px]" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.insur} onChange={e => handleEditChange('insur', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px]" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.tax} onChange={e => handleEditChange('tax', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px] text-rose-600" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.penalty} onChange={e => handleEditChange('penalty', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px] text-rose-600" /></td>
                          <td className="p-0.5"><input type="number" value={editFormData.takaful} onChange={e => handleEditChange('takaful', e.target.value)} className="w-9 sm:w-12 border rounded px-0.5 py-0.5 text-center text-[8px] sm:text-[10px]" /></td>
                          <td className="p-0.5 sm:p-1 text-center font-black bg-indigo-100 text-indigo-900 tabular-nums">{editFormData.net.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                        </tr>
                      ) : (
                      <tr
                        key={idx}
                        className={`${idx % 2 === 0 ? 'bg-indigo-50/50' : 'bg-white'} hover:bg-indigo-100/50 transition-colors cursor-pointer`}
                        onClick={() => toggleSelect(idx)}
                      >
                        <td className="py-1 px-0.5 sm:px-1 text-center" onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            checked={emp.selected}
                            onChange={() => toggleSelect(idx)}
                            className="accent-indigo-500 w-2.5 h-2.5 sm:w-3 sm:h-3 cursor-pointer"
                          />
                        </td>
                        <td className="py-1 px-0.5 sm:px-1 text-center text-slate-500">{emp.serial}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center font-mono font-bold text-indigo-700">{emp.code}</td>
                        <td className="py-1 px-1 sm:px-1.5">
                          <span className="font-bold whitespace-nowrap truncate block" title={emp.name}>{getTwoPartName(emp.name)}</span>
                        </td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums">{emp.days}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums text-rose-600">{emp.advance ? emp.advance.toLocaleString('en-US') : '0'}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums">{emp.basic.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums">{emp.trans.toLocaleString('en-US')}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums text-emerald-600">{emp.bonus.toLocaleString('en-US')}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums">{emp.overtime.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums">{emp.insur.toLocaleString('en-US')}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums text-rose-600">{emp.tax.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums text-rose-600">{emp.penalty.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                        <td className="py-1 px-0.5 sm:px-1 text-center tabular-nums">{emp.takaful.toLocaleString('en-US')}</td>
                        <td className="py-1 px-0.5 sm:px-1.5 text-center tabular-nums font-black bg-indigo-50 text-indigo-900">{emp.net.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                      </tr>
                      )
                    ))}
                  </tbody>
              </table>
           ) : (
              <div className="flex flex-col items-center justify-center h-full text-slate-400 space-y-2">
                 <FileSpreadsheet className="w-10 h-10 text-slate-300" />
                 <p className="text-xs sm:text-sm">قم برفع ملف Excel لعرض البيانات هنا</p>
              </div>
           )}
        </div>
      </div>
      
      {/* KPI Grid */}
      <div className="grid grid-cols-3 gap-1.5 sm:gap-3 px-2 sm:px-4 py-1.5 sm:py-2 shrink-0">
        <div className="bg-white p-1.5 sm:p-2 rounded-lg border border-slate-200 shadow-sm">
          <p className="text-slate-500 text-[9px] sm:text-xs font-bold uppercase tracking-wider">إجمالي الموظفين</p>
          <p className="text-sm sm:text-xl font-black text-slate-900">{totalCount}</p>
        </div>
        <div className="bg-white p-1.5 sm:p-2 rounded-lg border border-indigo-200 shadow-sm ring-1 ring-indigo-500/10">
          <p className="text-indigo-600 text-[9px] sm:text-xs font-bold uppercase tracking-wider">تم تحديدهم للتصدير</p>
          <p className="text-sm sm:text-xl font-black text-indigo-700">{selectedCount}</p>
        </div>
        <div className="bg-white p-1.5 sm:p-2 rounded-lg border border-slate-200 shadow-sm">
          <p className="text-slate-500 text-[9px] sm:text-xs font-bold uppercase tracking-wider">إجمالي الصافي</p>
          <p className="text-sm sm:text-xl font-black text-slate-900 text-left tracking-tighter tabular-nums">
            {totalNet.toLocaleString('en-US', { minimumFractionDigits: 2 })} <span className="text-[9px] sm:text-xs font-normal">ج.م</span>
          </p>
        </div>
      </div>

      {/* Footer */}
      <footer className="bg-slate-900 text-slate-400 px-3 sm:px-6 py-1 flex justify-between items-center shrink-0 text-[9px] sm:text-[10px]">
         <div>&copy; {new Date().getFullYear()} مُرَتَّبَا - جميع الحقوق محفوظة</div>
         <div className="flex gap-2 sm:gap-4">
            <span className="hidden sm:inline">تواصل مع الدعم الفني</span>
            <span className="hidden sm:inline">دليل المستخدم</span>
            <span className="text-indigo-400 font-bold uppercase tracking-widest">SECURE-v2</span>
         </div>
      </footer>

      {/* Hidden PDF Container */}
      <div ref={pdfRenderContainerRef} className="hidden"></div>

      {cropImageSrc && (
        <div className="fixed inset-0 z-[100] bg-black/70 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl p-4 max-w-lg w-full flex flex-col gap-4">
            <h3 className="font-bold text-slate-800 text-lg text-right">قص الشعار</h3>
            <div className="overflow-auto max-h-[60vh] flex justify-center bg-slate-50 border border-slate-200 rounded">
              <ReactCrop
                crop={crop}
                onChange={(_, percentCrop) => setCrop(percentCrop)}
                onComplete={(c) => setCompletedCrop(c)}
              >
                <img
                  ref={imgRef}
                  src={cropImageSrc}
                  alt="Crop me"
                  className="max-w-full"
                  onLoad={(e) => {
                     setCrop({ unit: '%', width: 90, height: 90, x: 5, y: 5 });
                  }}
                />
              </ReactCrop>
            </div>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => { setCropImageSrc(null); setCrop(undefined); setCompletedCrop(null); }}
                className="px-4 py-2 text-slate-600 bg-slate-100 hover:bg-slate-200 rounded font-bold"
              >
                إلغاء
              </button>
              <button
                onClick={handleConfirmCrop}
                className="px-4 py-2 text-white bg-indigo-600 hover:bg-indigo-700 rounded font-bold flex items-center gap-2"
              >
                <Check className="w-4 h-4" /> تأكيد القص
              </button>
            </div>
          </div>
        </div>
      )}

      {showSettingsPasswordModal && (
        <div
          className="fixed inset-0 z-[110] bg-slate-950/70 flex items-center justify-center p-4"
          onClick={() => setShowSettingsPasswordModal(false)}
        >
          <div
            className="bg-white rounded-xl p-5 max-w-sm w-full shadow-2xl border border-slate-200 space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-slate-100 pb-2.5">
              <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                <KeyRound className="w-5 h-5 text-indigo-600" />
                إدخال كلمة مرور الإعدادات
              </h3>
              <button
                onClick={() => setShowSettingsPasswordModal(false)}
                className="p-1 hover:bg-slate-100 rounded-full text-slate-500"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleSettingsPasswordSubmit} className="space-y-3">
              <input
                type="password"
                value={settingsPasswordInput}
                onChange={(e) => {
                  setSettingsPasswordInput(e.target.value);
                  setSettingsPasswordError('');
                }}
                placeholder="أدخل كلمة المرور (الافتراضية: 0000)"
                className="w-full px-3 py-2.5 rounded-lg border border-slate-300 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                autoFocus
              />
              {settingsPasswordError && (
                <p className="text-xs font-bold text-rose-600">{settingsPasswordError}</p>
              )}
              <div className="flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowSettingsPasswordModal(false)}
                  className="px-4 py-2 text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg text-xs font-bold"
                >
                  إلغاء
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg text-xs font-bold"
                >
                  دخول
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {saveDialog && (
        <div
          className="fixed inset-0 z-[115] bg-slate-950/70 flex items-center justify-center p-4"
          onClick={() => !isSavingFile && setSaveDialog(null)}
        >
          <div
            className="bg-white rounded-2xl p-5 max-w-md w-full shadow-2xl border border-slate-200 space-y-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${saveDialog.type === 'pdf' ? 'bg-rose-100 text-rose-600' : 'bg-emerald-100 text-emerald-600'}`}>
                  {saveDialog.type === 'pdf' ? <FileText className="w-5 h-5" /> : <FileSpreadsheet className="w-5 h-5" />}
                </div>
                <div>
                  <h3 className="font-bold text-slate-900 text-base">
                    {saveDialog.type === 'pdf' ? 'حفظ ملف PDF' : 'حفظ ملف Excel'}
                  </h3>
                  <p className="text-xs text-slate-500">
                    {saveDialog.employees.length === 1
                      ? `الموظف: ${saveDialog.employees[0].name}`
                      : `عدد الموظفين المحددين: ${saveDialog.employees.length}`}
                  </p>
                </div>
              </div>
              <button
                onClick={() => !isSavingFile && setSaveDialog(null)}
                className="p-1.5 hover:bg-slate-100 rounded-full text-slate-500"
                disabled={isSavingFile}
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-2">
              <label className="block text-xs font-bold text-slate-700">
                اسم الملف (يمكنك تعديله قبل الحفظ)
              </label>
              <div className="flex items-center border border-slate-300 rounded-lg overflow-hidden focus-within:ring-2 focus-within:ring-indigo-500 focus-within:border-indigo-500 bg-white">
                <input
                  type="text"
                  value={saveDialog.fileName}
                  onChange={(e) => {
                    setSaveDialog({ ...saveDialog, fileName: e.target.value });
                    setSaveDialogError('');
                  }}
                  placeholder="أدخل اسم الملف..."
                  className="flex-1 px-3 py-2.5 text-sm outline-none text-slate-800"
                  autoFocus
                  disabled={isSavingFile}
                />
                <span className="bg-slate-100 text-slate-600 px-3 py-2.5 text-xs font-mono font-bold border-r border-slate-200" dir="ltr">
                  {saveDialog.type === 'pdf' ? '.pdf' : '.xlsx'}
                </span>
              </div>
              {saveDialogError && (
                <p className="text-xs font-bold text-rose-600">{saveDialogError}</p>
              )}
            </div>

            <div className="space-y-2 pt-1">
              <button
                type="button"
                disabled={isSavingFile}
                onClick={() => handleConfirmSaveFile('picker')}
                className="w-full bg-indigo-600 hover:bg-indigo-700 disabled:opacity-60 text-white py-2.5 px-4 rounded-xl text-sm font-bold transition-colors flex items-center justify-center gap-2 shadow-sm cursor-pointer"
              >
                <FolderOpen className="w-4 h-4 shrink-0" />
                <span>{isSavingFile ? 'جاري تجهيز الملف...' : 'اختيار مكان الحفظ وحفظ الملف'}</span>
              </button>

              {typeof navigator !== 'undefined' && typeof navigator.share === 'function' && (
                <button
                  type="button"
                  disabled={isSavingFile}
                  onClick={() => handleConfirmSaveFile('share')}
                  className="w-full bg-slate-100 hover:bg-slate-200 disabled:opacity-60 text-slate-800 py-2.5 px-4 rounded-xl text-xs font-bold transition-colors flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Save className="w-4 h-4 shrink-0" />
                  <span>حفظ في ملفات الهاتف أو المشاركة</span>
                </button>
              )}

              <button
                type="button"
                disabled={isSavingFile}
                onClick={() => handleConfirmSaveFile('download')}
                className="w-full bg-white hover:bg-slate-50 border border-slate-200 disabled:opacity-60 text-slate-700 py-2 px-4 rounded-xl text-xs font-bold transition-colors flex items-center justify-center gap-2 cursor-pointer"
              >
                <Download className="w-4 h-4 shrink-0" />
                <span>تنزيل مباشر بالاسم المحدد</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {saveToast && (
        <div className="fixed bottom-4 right-4 z-[120] bg-emerald-600 text-white px-4 py-2.5 rounded-xl shadow-lg text-xs font-bold flex items-center gap-2">
          <Check className="w-4 h-4 shrink-0" />
          <span>{saveToast}</span>
        </div>
      )}

      <OfflineIndicator />
    </div>
  );
}

