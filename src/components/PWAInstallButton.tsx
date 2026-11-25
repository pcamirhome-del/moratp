import React, { useState } from 'react';
import { Download, Smartphone, X } from 'lucide-react';
import { usePWAInstall } from '../hooks/usePWAInstall';

export const PWAInstallButton: React.FC = () => {
  const { isInstallable, isInstalled, isIOS, install } = usePWAInstall();
  const [showGuide, setShowGuide] = useState(false);

  if (isInstalled) {
    return null;
  }

  if (isInstallable) {
    return (
      <button
        onClick={install}
        className="flex items-center gap-1.5 rounded bg-emerald-600 hover:bg-emerald-700 px-3 py-2 text-xs sm:text-sm font-bold text-white shadow-sm transition-colors cursor-pointer"
        title="تثبيت التطبيق على الهاتف أو الجهاز"
      >
        <Download className="w-4 h-4 shrink-0" />
        <span>تثبيت التطبيق</span>
      </button>
    );
  }

  return (
    <>
      <button
        onClick={() => setShowGuide(true)}
        className="flex items-center gap-1.5 rounded bg-indigo-800/80 hover:bg-indigo-700 border border-indigo-400/40 px-2.5 py-2 text-xs font-bold text-indigo-100 transition-colors cursor-pointer"
        title="تثبيت التطبيق على الهاتف"
      >
        <Smartphone className="w-4 h-4 shrink-0" />
        <span className="hidden sm:inline">تثبيت التطبيق</span>
      </button>

      {showGuide && (
        <div
          className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-950/70 p-4"
          onClick={() => setShowGuide(false)}
          dir="rtl"
        >
          <div
            className="w-full max-w-sm rounded-xl bg-white p-5 shadow-2xl border border-slate-200 text-slate-800"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 mb-3">
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Smartphone className="w-5 h-5 text-indigo-600" />
                تثبيت تطبيق مُرَتَّبَا على الهاتف
              </h3>
              <button
                onClick={() => setShowGuide(false)}
                className="p-1 hover:bg-slate-100 rounded-full text-slate-500"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {isIOS ? (
              <div className="space-y-2.5 text-sm text-slate-600 leading-relaxed">
                <p>لتثبيت التطبيق على أجهزة iPhone / iPad وتشغيله كتطبيق مستقل:</p>
                <ol className="list-decimal list-inside space-y-1.5 font-medium text-slate-700">
                  <li>افتح التطبيق في متصفح <strong>Safari</strong>.</li>
                  <li>اضغط على زر <strong>المشاركة (Share)</strong> في شريط المتصفح.</li>
                  <li>اختر <strong>إضافة إلى الشاشة الرئيسية (Add to Home Screen)</strong>.</li>
                </ol>
              </div>
            ) : (
              <div className="space-y-2.5 text-sm text-slate-600 leading-relaxed">
                <p>لتثبيت التطبيق على الهاتف وتشغيله مباشرة بعرض الشاشة:</p>
                <ol className="list-decimal list-inside space-y-1.5 font-medium text-slate-700">
                  <li>افتح قائمة المتصفح (النقاط الثلاث <strong>⋮</strong> في الأعلى).</li>
                  <li>اختر <strong>تثبيت التطبيق (Install App)</strong> أو <strong>الإضافة إلى الشاشة الرئيسية</strong>.</li>
                  <li>سيظهر التطبيق كأيقونة مستقلة على هاتفك ويعمل بعرض الشاشة.</li>
                </ol>
              </div>
            )}

            <button
              onClick={() => setShowGuide(false)}
              className="mt-4 w-full rounded-lg bg-indigo-600 hover:bg-indigo-700 py-2 text-sm font-bold text-white transition-colors"
            >
              حسناً
            </button>
          </div>
        </div>
      )}
    </>
  );
};
