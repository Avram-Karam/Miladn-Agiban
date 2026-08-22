/**
 * gas-proxy.js — Netlify Serverless Function
 * ─────────────────────────────────────────────────────────────────────────────
 * يعمل كـ Proxy آمن بين الموقع و Google Apps Script.
 * الـ GAS_URL الحقيقي مخزن في Netlify Environment Variables فقط.
 * الـ Security Token مخزن في GAS_TOKEN Environment Variable.
 *
 * Supports:
 *   GET  /_api/gas            → جلب كل المشتركين من Sheets
 *   POST /_api/gas {action}   → أي عملية إدارية (add, update, bulkImport ...)
 * ─────────────────────────────────────────────────────────────────────────────
 */

exports.handler = async (event) => {
    const GAS_URL = process.env.GAS_URL || 'https://script.google.com/macros/s/AKfycbzTq2wFN35ZvwNdMpx7_DlZLQhZIQH0HRZ2e1p_HhcjLItG7c_SfXM_OaNPilG4UcwDFg/exec';

    // ── إذا لم يُضبط المتغير في Netlify ──
    if (!GAS_URL) {
        return respond(503, {
            status : 'error',
            message: 'GAS_URL environment variable is not set in Netlify dashboard.'
        });
    }

    // ── CORS Preflight ──
    if (event.httpMethod === 'OPTIONS') {
        return {
            statusCode: 200,
            headers   : corsHeaders(),
            body      : ''
        };
    }

    try {
        const method = event.httpMethod || 'GET';

        // ── بناء الـ URL المستهدف مع Query Params ──
        let targetUrl = GAS_URL;
        if (event.queryStringParameters && Object.keys(event.queryStringParameters).length > 0) {
            const params = new URLSearchParams(event.queryStringParameters).toString();
            targetUrl += (targetUrl.includes('?') ? '&' : '?') + params;
        }

        // ── إعداد الطلب لـ GAS ──
        const fetchOptions = {
            method : method,
            headers: { 'Content-Type': 'application/json' },
            redirect: 'follow'
        };

        // ── في حالة POST: أضف Token تلقائياً إذا لم يكن موجوداً ──
        if (method === 'POST' && event.body) {
            let body = {};
            try { body = JSON.parse(event.body); } catch(e) { body = {}; }

            // أضف الـ GAS_TOKEN من البيئة إذا كان موجوداً ولم يُرسل من الـ Client
            const envToken = process.env.GAS_TOKEN || '';
            if (envToken && !body.token) {
                body.token = envToken;
            }
            fetchOptions.body = JSON.stringify(body);
        }

        // ── إرسال الطلب لـ GAS ──
        const gasResponse = await fetch(targetUrl, fetchOptions);
        const responseText = await gasResponse.text();

        // ── محاولة parse الـ JSON ──
        let responseData;
        try {
            responseData = JSON.parse(responseText);
        } catch(e) {
            // إذا فشل الـ parse أعد النص كما هو مع خطأ
            return respond(502, {
                status : 'error',
                message: 'Invalid JSON response from Google Apps Script',
                raw    : responseText.substring(0, 500)
            });
        }

        return respond(200, responseData);

    } catch (err) {
        console.error('[gas-proxy] Error:', err.message);
        return respond(502, {
            status : 'error',
            message: err.message
        });
    }
};

// ─── Response Helper ──────────────────────────────────────────────────────────
function respond(statusCode, data) {
    return {
        statusCode,
        headers: corsHeaders(),
        body   : JSON.stringify(data)
    };
}

// ─── CORS Headers ─────────────────────────────────────────────────────────────
function corsHeaders() {
    return {
        'Content-Type'                : 'application/json',
        'Access-Control-Allow-Origin' : '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Cache-Control'               : 'no-store, no-cache'
    };
}
