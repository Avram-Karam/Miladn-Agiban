/**
 * api/gas.js — Vercel Serverless Function
 * ─────────────────────────────────────────────────────────────────────────────
 * يعمل كـ Proxy آمن بين الموقع و Google Apps Script على Vercel.
 * الـ GAS_URL الحقيقي مخزن في Vercel Environment Variables فقط.
 *
 * Supports:
 *   GET  /api/gas            → جلب كل المشتركين من Sheets
 *   POST /api/gas {action}   → أي عملية إدارية (add, update, bulkImport ...)
 * ─────────────────────────────────────────────────────────────────────────────
 */

export default async function handler(req, res) {
    // ── CORS ──
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Cache-Control', 'no-store, no-cache');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    const GAS_URL = process.env.GAS_URL || 'https://script.google.com/macros/s/AKfycbzTq2wFN35ZvwNdMpx7_DlZLQhZIQH0HRZ2e1p_HhcjLItG7c_SfXM_OaNPilG4UcwDFg/exec';

    if (!GAS_URL) {
        return res.status(503).json({
            status : 'error',
            message: 'GAS_URL environment variable is not set in Vercel dashboard.'
        });
    }

    try {
        const method = req.method || 'GET';

        // بناء الـ URL المستهدف مع Query Params
        let targetUrl = GAS_URL;
        const params = new URLSearchParams(req.query).toString();
        if (params) {
            targetUrl += (targetUrl.includes('?') ? '&' : '?') + params;
        }

        const fetchOptions = {
            method : method,
            headers: { 'Content-Type': 'application/json' },
            redirect: 'follow'
        };

        // في حالة POST: أضف Token تلقائياً
        if (method === 'POST' && req.body) {
            let body = req.body;
            if (typeof body === 'string') {
                try { body = JSON.parse(body); } catch(e) { body = {}; }
            }
            const envToken = process.env.GAS_TOKEN || '';
            if (envToken && !body.token) {
                body.token = envToken;
            }
            fetchOptions.body = JSON.stringify(body);
        }

        const gasResponse = await fetch(targetUrl, fetchOptions);
        const responseText = await gasResponse.text();

        let responseData;
        try {
            responseData = JSON.parse(responseText);
        } catch(e) {
            return res.status(502).json({
                status : 'error',
                message: 'Invalid JSON response from Google Apps Script',
                raw    : responseText.substring(0, 500)
            });
        }

        return res.status(200).json(responseData);

    } catch (err) {
        console.error('[gas-proxy] Error:', err.message);
        return res.status(502).json({
            status : 'error',
            message: err.message
        });
    }
}
