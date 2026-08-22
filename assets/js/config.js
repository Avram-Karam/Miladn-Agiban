/**
 * ════════════════════════════════════════════════════════════
 *  config.js — إعدادات مؤتمر الشباب 2026
 *  ✅ آمن للرفع على GitHub
 *  🔒 يعمل على GitHub Pages و Netlify تلقائياً
 * ════════════════════════════════════════════════════════════
 *
 *  Netlify  → /_api/gas (Proxy مخفي) ← GAS_URL في Environment Variables
 *  GitHub Pages → DIRECT_GAS_URL مباشرة (مع no-cors للكتابة)
 *
 *  🔒 الأمان: GAS_TOKEN لا يُخزَّن هنا بعد الآن.
 *  على Netlify: يُحقَّن التوكن من السيرفر تلقائياً عبر gas-proxy.js
 *  على GitHub Pages: يُرسَّل الطلب بدون توكن (GAS يتحقق من Origin)
 */

window.YC_CONFIG = {
    // ── مسار الـ Proxy الآمن (Netlify فقط) ──
    GAS_URL: '/_api/gas',

    // ── الرابط المباشر لـ Google Apps Script (GitHub Pages fallback) ──
    // هذا الرابط يُستخدم فقط عند عدم توفر Netlify Proxy
    DIRECT_GAS_URL: 'https://script.google.com/macros/s/AKfycbzTq2wFN35ZvwNdMpx7_DlZLQhZIQH0HRZ2e1p_HhcjLItG7c_SfXM_OaNPilG4UcwDFg/exec',

    // ── إصدار الإعدادات ──
    CONFIG_VERSION: '2026.08.22-profile-fix2'
};
