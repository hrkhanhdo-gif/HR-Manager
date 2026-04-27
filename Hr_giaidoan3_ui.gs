// =============================================================================
// HRIS GIAI ĐOẠN 3 — TÍCH HỢP AI (CLAUDE API)
// Dán file này vào Apps Script Editor, BÊN DƯỚI file Giai Đoạn 1 & 2
// =============================================================================
// MODULE 14 — Cấu hình AI & API Key management
// MODULE 15 — Phân tích CV tự động
// MODULE 16 — Chatbot HR (hỏi đáp dữ liệu nhân sự bằng tiếng Việt)
// MODULE 17 — Tóm tắt kết quả phỏng vấn
// MODULE 18 — Gợi ý hành động từ dữ liệu tuyển dụng
// MODULE 19 — UI AI Panel (tích hợp vào Web App Giai đoạn 2)
// =============================================================================


// =============================================================================
// MODULE 14 — CẤU HÌNH AI
// =============================================================================

const AI_CONFIG = {
  MODEL:      'claude-opus-4-6',
  MAX_TOKENS: 1500,
  API_URL:    'https://api.anthropic.com/v1/messages',

  // System prompt chung cho HR context
  HR_SYSTEM_PROMPT: `Bạn là trợ lý HR thông minh cho một công ty Việt Nam. 
Bạn có quyền truy cập dữ liệu nhân sự và tuyển dụng.
Trả lời bằng tiếng Việt, ngắn gọn, chuyên nghiệp.
Khi phân tích dữ liệu, đưa ra nhận xét cụ thể và hành động khuyến nghị.
Không bịa đặt thông tin không có trong dữ liệu được cung cấp.`,
};

/**
 * Lưu API Key vào PropertiesService (an toàn, không lộ trong code)
 * Chạy hàm này 1 lần để set API key
 * @param {string} apiKey - Claude API key
 */
function setApiKey(apiKey) {
  PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY', apiKey);
  Logger.log('✅ API key đã được lưu an toàn');
}

/**
 * Lấy API key từ PropertiesService
 * @returns {string}
 */
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!key) throw new Error('Chưa cài đặt Claude API key. Chạy hàm setApiKey("sk-ant-...") trước.');
  return key;
}

/**
 * Gọi Claude API
 * @param {string} userMessage - nội dung cần xử lý
 * @param {string} systemPrompt - context/persona (optional, dùng HR_SYSTEM_PROMPT mặc định)
 * @param {Array}  history      - lịch sử hội thoại [{role, content}] (optional)
 * @returns {string} phản hồi từ Claude
 */
function callClaude(userMessage, systemPrompt, history) {
  const apiKey = getApiKey();

  const messages = [];
  // Thêm lịch sử nếu có (cho chatbot multi-turn)
  if (history && Array.isArray(history)) {
    history.forEach(h => messages.push({ role: h.role, content: h.content }));
  }
  messages.push({ role: 'user', content: userMessage });

  const payload = {
    model:      AI_CONFIG.MODEL,
    max_tokens: AI_CONFIG.MAX_TOKENS,
    system:     systemPrompt || AI_CONFIG.HR_SYSTEM_PROMPT,
    messages,
  };

  const options = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response  = UrlFetchApp.fetch(AI_CONFIG.API_URL, options);
  const result    = JSON.parse(response.getContentText());

  if (result.error) throw new Error(result.error.message);
  return result.content[0].text;
}


// =============================================================================
// MODULE 15 — PHÂN TÍCH CV TỰ ĐỘNG
// =============================================================================

/**
 * Phân tích CV ứng viên và so sánh với yêu cầu tuyển dụng
 * @param {string} uvId      - ID ứng viên
 * @param {string} jdText    - Mô tả công việc (Job Description)
 * @param {string} cvText    - Nội dung CV (text, sau khi extract từ PDF/Drive)
 * @returns {Object} { score, summary, strengths, gaps, recommendation }
 */
function analyzeCv(uvId, jdText, cvText) {
  const prompt = `Bạn là chuyên gia tuyển dụng HR. Hãy phân tích CV ứng viên so với yêu cầu công việc.

=== MÔ TẢ CÔNG VIỆC (JD) ===
${jdText}

=== NỘI DUNG CV ===
${cvText}

Hãy phân tích và trả về JSON (chỉ JSON, không có text nào khác):
{
  "score": <số từ 0-100, độ phù hợp tổng thể>,
  "summary": "<tóm tắt ngắn 2-3 câu về ứng viên>",
  "strengths": ["<điểm mạnh 1>", "<điểm mạnh 2>", "<điểm mạnh 3>"],
  "gaps": ["<điểm thiếu 1>", "<điểm thiếu 2>"],
  "recommendation": "<Pass/Consider/Reject>",
  "suggested_questions": ["<câu hỏi PV gợi ý 1>", "<câu hỏi PV gợi ý 2>", "<câu hỏi PV gợi ý 3>"],
  "salary_range": "<đề xuất mức lương phù hợp nếu có thông tin>"
}`;

  const systemPrompt = `Bạn là chuyên gia tuyển dụng HR senior với 10 năm kinh nghiệm. 
Phân tích khách quan, dựa trên dữ liệu thực tế trong CV.
Chỉ trả về JSON thuần túy, không markdown, không giải thích thêm.`;

  try {
    const raw     = callClaude(prompt, systemPrompt);
    const cleaned = raw.replace(/```json|```/g, '').trim();
    const result  = JSON.parse(cleaned);

    // Lưu kết quả vào sheet UNG_VIEN
    saveAiAnalysis(uvId, 'CV_ANALYSIS', result);
    writeAuditLog('UPDATE', CONFIG.SHEETS.UNG_VIEN, uvId, 'AI_Analysis', '', result.score, 'AI phân tích CV');

    return { success: true, data: result };
  } catch (e) {
    Logger.log('analyzeCv error: ' + e.message);
    return { success: false, error: 'Lỗi phân tích CV: ' + e.message };
  }
}

/**
 * Lưu kết quả phân tích AI vào sheet AI_RESULTS
 */
function saveAiAnalysis(entityId, analysisType, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AI_RESULTS');

  if (!sheet) {
    sheet = ss.insertSheet('AI_RESULTS');
    const headers = ['ID', 'Entity_ID', 'Analysis_Type', 'Score', 'Result_JSON', 'Created_At'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2C2C2A').setFontColor('#FFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    generateId('AI'),
    entityId,
    analysisType,
    result.score || '',
    JSON.stringify(result),
    new Date(),
  ]);
}

/**
 * Batch: Phân tích tất cả ứng viên "Đang xét" chưa được phân tích
 * Chạy thủ công hoặc đặt trigger hàng ngày
 */
function runBatchCvAnalysis() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const uvSheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);
  if (!uvSheet) return;

  const data = uvSheet.getDataRange().getValues().slice(1);
  const toAnalyze = data.filter(r => r[5] === 'Đang xét' && r[8]); // có link CV

  Logger.log(`Batch CV analysis: ${toAnalyze.length} ứng viên cần phân tích`);

  toAnalyze.forEach(row => {
    const uvId  = row[0];
    const vitri = row[2];

    // JD mặc định (thực tế bạn lấy từ sheet TUYEN_DUNG)
    const jd = `Vị trí: ${vitri}. Yêu cầu: kinh nghiệm liên quan, kỹ năng giao tiếp tốt, teamwork.`;

    // Trong thực tế: fetch CV từ Google Drive bằng link row[8]
    // Tạm thời dùng placeholder
    const cvText = `CV ứng viên cho vị trí ${vitri} — link: ${row[8]}`;

    try {
      analyzeCv(uvId, jd, cvText);
      Utilities.sleep(2000); // Tránh rate limit
    } catch (e) {
      Logger.log(`Lỗi phân tích UV ${uvId}: ${e.message}`);
    }
  });
}


// =============================================================================
// MODULE 16 — CHATBOT HR (Multi-turn)
// =============================================================================

/**
 * Xử lý tin nhắn chatbot từ frontend
 * @param {string} message   - câu hỏi của user
 * @param {Array}  history   - lịch sử hội thoại [{role, content}]
 * @returns {Object} { reply, suggestedQuestions }
 */
function chatWithHr(message, history) {
  // Lấy context dữ liệu hiện tại từ sheets
  const context = buildHrContext();

  const systemPrompt = `${AI_CONFIG.HR_SYSTEM_PROMPT}

=== DỮ LIỆU HIỆN TẠI ===
${context}

Khi trả lời:
- Nếu hỏi về số liệu cụ thể, trích dẫn từ dữ liệu trên
- Nếu phát hiện vấn đề (turnover cao, vị trí trống lâu), chủ động đề xuất giải pháp
- Câu trả lời tối đa 200 từ, trừ khi được yêu cầu phân tích chi tiết
- Cuối mỗi câu trả lời, đề xuất 2-3 câu hỏi tiếp theo dưới dạng JSON:
  {"reply": "...", "suggestions": ["câu gợi ý 1", "câu gợi ý 2"]}
- Chỉ trả về JSON thuần túy`;

  try {
    const raw     = callClaude(message, systemPrompt, history || []);
    const cleaned = raw.replace(/```json|```/g, '').trim();

    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch (_) {
      // Nếu AI không trả về JSON đúng format
      parsed = { reply: raw, suggestions: ['Xem báo cáo tổng quan', 'Phân tích turnover', 'Danh sách ứng viên'] };
    }

    writeAuditLog('CREATE', 'AI_CHAT', '', 'ChatMessage', message.slice(0, 50), '', 'HR Chatbot');
    return { success: true, data: parsed };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Xây dựng context tóm tắt từ dữ liệu thực tế trong sheets
 * @returns {string}
 */
function buildHrContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = [];

  try {
    const nsSheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
    if (nsSheet) {
      const nsData = nsSheet.getDataRange().getValues().slice(1).filter(r => r[0]);
      const active = nsData.filter(r => r[8] !== 'Đã nghỉ');
      const byDept = {};
      active.forEach(r => { const d = r[2]||'Chưa phân'; byDept[d]=(byDept[d]||0)+1; });
      lines.push(`NHÂN SỰ: Tổng ${active.length} người đang làm việc.`);
      lines.push(`Phân bổ bộ phận: ${Object.entries(byDept).map(([k,v])=>`${k}(${v})`).join(', ')}`);
    }
  } catch (e) {}

  try {
    const uvSheet = ss.getSheetByName(CONFIG.SHEETS.UNG_VIEN);
    if (uvSheet) {
      const uvData = uvSheet.getDataRange().getValues().slice(1).filter(r => r[0]);
      const byStatus = {};
      uvData.forEach(r => { const s = r[5]||'Mới nộp'; byStatus[s]=(byStatus[s]||0)+1; });
      lines.push(`TUYỂN DỤNG: ${uvData.length} ứng viên tổng cộng.`);
      lines.push(`Pipeline: ${Object.entries(byStatus).map(([k,v])=>`${k}(${v})`).join(', ')}`);
    }
  } catch (e) {}

  try {
    const bdays = getSinhNhatTuanNay();
    if (bdays.length > 0) {
      lines.push(`SINH NHẬT TUẦN NÀY: ${bdays.map(b=>b.name+(b.days_until===0?' (hôm nay)':' ('+b.days_until+'ngày)')).join(', ')}`);
    }
  } catch (e) {}

  return lines.join('\n') || 'Chưa có dữ liệu trong hệ thống.';
}


// =============================================================================
// MODULE 17 — TÓM TẮT KẾT QUẢ PHỎNG VẤN
// =============================================================================

/**
 * Tạo tóm tắt & đề xuất quyết định sau phỏng vấn
 * @param {string} pvId      - ID phỏng vấn
 * @param {string} rawNotes  - ghi chú thô từ người phỏng vấn
 * @param {string} uvName    - tên ứng viên
 * @param {string} vitri     - vị trí ứng tuyển
 * @returns {Object}
 */
function summarizeInterview(pvId, rawNotes, uvName, vitri) {
  const prompt = `Bạn là HR chuyên nghiệp. Dưới đây là ghi chú thô sau buổi phỏng vấn:

Ứng viên: ${uvName}
Vị trí: ${vitri}
Ghi chú của người phỏng vấn:
"${rawNotes}"

Hãy xử lý và trả về JSON:
{
  "summary": "<tóm tắt chuyên nghiệp 3-5 câu>",
  "technical_rating": <1-5>,
  "soft_skill_rating": <1-5>,
  "culture_fit_rating": <1-5>,
  "overall_rating": <1-5>,
  "key_points": ["<điểm chính 1>", "<điểm chính 2>"],
  "concerns": ["<lo ngại nếu có>"],
  "decision": "<Offer/Second Round/Reject/On Hold>",
  "decision_reason": "<lý do ngắn gọn>",
  "next_steps": "<bước tiếp theo đề xuất>"
}`;

  try {
    const raw     = callClaude(prompt, AI_CONFIG.HR_SYSTEM_PROMPT);
    const cleaned = raw.replace(/```json|```/g, '').trim();
    const result  = JSON.parse(cleaned);

    // Lưu tóm tắt vào sheet PHONG_VAN
    updateInterviewSummary(pvId, result);
    writeAuditLog('UPDATE', CONFIG.SHEETS.PHONG_VAN, pvId, 'AI_Summary', '', result.decision, 'AI tóm tắt PV');

    return { success: true, data: result };
  } catch (e) {
    return { success: false, error: 'Lỗi tóm tắt phỏng vấn: ' + e.message };
  }
}

function updateInterviewSummary(pvId, summary) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PHONG_VAN);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === pvId) {
      // Cập nhật cột Kết quả và Ghi chú với AI summary
      sheet.getRange(i + 1, 8).setValue(summary.decision);
      sheet.getRange(i + 1, 9).setValue(
        `[AI] ${summary.summary} | Rating: ${summary.overall_rating}/5 | ${summary.decision_reason}`
      );
      break;
    }
  }
}


// =============================================================================
// MODULE 18 — GỢI Ý HÀNH ĐỘNG & PHÂN TÍCH TURNOVER
// =============================================================================

/**
 * Phân tích rủi ro turnover và đề xuất hành động
 * Chạy hàng tuần để phát hiện sớm
 * @returns {Object} { risks, actions, report }
 */
function analyzeTurnoverRisk() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const nsSheet = ss.getSheetByName(CONFIG.SHEETS.NHAN_SU);
  if (!nsSheet) return { success: false, error: 'Không tìm thấy sheet nhân sự' };

  const data    = nsSheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1).filter(r => r[0] && r[8] !== 'Đã nghỉ');

  // Tính toán các chỉ số rủi ro
  const today    = new Date();
  const riskData = rows.map(row => {
    const ngayVao = row[4] ? new Date(row[4]) : null;
    const months  = ngayVao ? Math.floor((today - ngayVao) / (30 * 24 * 3600 * 1000)) : 0;

    return {
      id:      row[0],
      name:    row[1],
      dept:    row[2],
      role:    row[3],
      months,
      status:  row[8],
    };
  });

  // Tìm rủi ro cao: 18-36 tháng mà vẫn cùng vai trò
  const highRisk = riskData.filter(r => r.months >= 18 && r.months <= 48);
  const newJoins = riskData.filter(r => r.months <= 3);

  const prompt = `Bạn là HR Analyst. Phân tích dữ liệu nhân sự và đánh giá rủi ro turnover:

TỔNG NHÂN SỰ: ${rows.length} người
CÓ NGUY CƠ CAO (18-48 tháng không thăng tiến): ${highRisk.length} người
  ${highRisk.slice(0, 5).map(r => `- ${r.name} (${r.dept}, ${r.months} tháng)`).join('\n  ')}

MỚI VÀO (<3 tháng): ${newJoins.length} người

Phân tích và trả về JSON:
{
  "turnover_risk_level": "<Low/Medium/High/Critical>",
  "risk_summary": "<nhận xét tổng quan 2-3 câu>",
  "immediate_actions": [
    {"action": "<hành động 1>", "priority": "High", "target": "<đối tượng>"},
    {"action": "<hành động 2>", "priority": "Medium", "target": "<đối tượng>"}
  ],
  "retention_strategies": ["<chiến lược 1>", "<chiến lược 2>", "<chiến lược 3>"],
  "hiring_suggestions": "<gợi ý tuyển dụng nếu cần>",
  "onboarding_alert": "<cảnh báo về nhân viên mới nếu cần>"
}`;

  try {
    const raw     = callClaude(prompt, AI_CONFIG.HR_SYSTEM_PROMPT);
    const cleaned = raw.replace(/```json|```/g, '').trim();
    const result  = JSON.parse(cleaned);

    writeAuditLog('CREATE', 'AI_ANALYSIS', '', 'TurnoverRisk', '', result.turnover_risk_level, 'Phân tích turnover AI');
    saveAiAnalysis('SYSTEM', 'TURNOVER_RISK', result);

    return { success: true, data: result, raw_stats: { total: rows.length, high_risk: highRisk.length, new_joins: newJoins.length } };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Tạo bản mô tả công việc (JD) tự động từ thông tin cơ bản
 * @param {Object} params - { vitri, bophan, yeu_cau_co_ban, muc_luong }
 */
function generateJobDescription(params) {
  const prompt = `Viết mô tả công việc (Job Description) chuyên nghiệp bằng tiếng Việt cho:

Vị trí: ${params.vitri}
Bộ phận: ${params.bophan || 'Chưa xác định'}
Yêu cầu cơ bản: ${params.yeu_cau_co_ban || 'Không có'}
Mức lương: ${params.muc_luong || 'Thỏa thuận'}

Trả về JSON:
{
  "title": "<tên chức danh chuẩn>",
  "overview": "<mô tả tổng quan vị trí 2-3 câu>",
  "responsibilities": ["<trách nhiệm 1>", "<trách nhiệm 2>", "<trách nhiệm 3>", "<trách nhiệm 4>", "<trách nhiệm 5>"],
  "requirements": ["<yêu cầu 1>", "<yêu cầu 2>", "<yêu cầu 3>"],
  "nice_to_have": ["<ưu tiên 1>", "<ưu tiên 2>"],
  "benefits": ["<quyền lợi 1>", "<quyền lợi 2>", "<quyền lợi 3>"],
  "salary": "${params.muc_luong || 'Thỏa thuận theo năng lực'}"
}`;

  try {
    const raw     = callClaude(prompt, AI_CONFIG.HR_SYSTEM_PROMPT);
    const cleaned = raw.replace(/```json|```/g, '').trim();
    return { success: true, data: JSON.parse(cleaned) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// MODULE 19 — API HANDLERS (thêm vào handleRequest switch)
// Copy các case này vào switch trong handleRequest() của Giai đoạn 1+2
// =============================================================================

/*
  Thêm các case sau vào switch của handleRequest():

  case 'AI_ANALYZE_CV':
    if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
    return analyzeCv(payload.uvId, payload.jdText, payload.cvText);

  case 'AI_CHAT':
    if (!hasPermission(user.role, 'READ')) throw new Error('Không có quyền');
    return chatWithHr(payload.message, payload.history);

  case 'AI_SUMMARIZE_INTERVIEW':
    if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
    return summarizeInterview(payload.pvId, payload.notes, payload.uvName, payload.vitri);

  case 'AI_TURNOVER_RISK':
    if (!hasPermission(user.role, 'READ')) throw new Error('Không có quyền');
    return analyzeTurnoverRisk();

  case 'AI_GENERATE_JD':
    if (!hasPermission(user.role, 'WRITE')) throw new Error('Không có quyền');
    return generateJobDescription(payload);
*/


// =============================================================================
// MODULE 19 — HTML AI PANEL
// Thêm vào cuối <body> của buildFullApp() trong Giai đoạn 2
// =============================================================================

function buildAiPanel() {
  return `

<!-- ═══════════════════════════════
     AI ASSISTANT PANEL
════════════════════════════════ -->
<style>
  /* ── AI FAB Button ── */
  .ai-fab {
    position: fixed; bottom: 28px; right: 28px; z-index: 300;
    width: 52px; height: 52px; border-radius: 50%;
    background: linear-gradient(135deg, #3D35A8, #1D9E75);
    border: none; cursor: pointer; box-shadow: 0 4px 20px rgba(61,53,168,0.4);
    display: flex; align-items: center; justify-content: center;
    font-size: 20px; color: #fff; transition: transform 0.2s, box-shadow 0.2s;
    font-family: inherit;
  }
  .ai-fab:hover { transform: scale(1.08); box-shadow: 0 6px 28px rgba(61,53,168,0.5); }

  /* ── AI Drawer ── */
  .ai-drawer {
    position: fixed; right: 0; top: 0; bottom: 0; width: 400px; max-width: 95vw;
    background: #fff; border-left: 1px solid rgba(0,0,0,0.08);
    box-shadow: -10px 0 40px rgba(0,0,0,0.12);
    z-index: 400; display: flex; flex-direction: column;
    transform: translateX(100%); transition: transform 0.3s ease;
  }
  .ai-drawer.open { transform: translateX(0); }

  .ai-drawer-header {
    padding: 18px 20px; border-bottom: 1px solid rgba(0,0,0,0.08);
    display: flex; align-items: center; justify-content: space-between;
    background: linear-gradient(135deg, #3D35A8, #2d279a);
    color: #fff;
  }
  .ai-drawer-title { font-size: 15px; font-weight: 600; }
  .ai-drawer-sub   { font-size: 11px; opacity: 0.7; margin-top: 2px; }
  .ai-close { cursor: pointer; font-size: 20px; opacity: 0.7; background:none; border:none; color:#fff; }

  /* ── Tabs ── */
  .ai-tabs { display: flex; border-bottom: 1px solid rgba(0,0,0,0.08); background: #fafafa; }
  .ai-tab {
    flex: 1; padding: 10px 4px; text-align: center; font-size: 11px; font-weight: 500;
    color: #888; cursor: pointer; border-bottom: 2px solid transparent;
    transition: all 0.15s; font-family: inherit; background: none; border-top: none; border-left: none; border-right: none;
  }
  .ai-tab.active { color: #3D35A8; border-bottom-color: #3D35A8; background: #fff; }

  /* ── Chat ── */
  .chat-messages {
    flex: 1; overflow-y: auto; padding: 16px; display: flex; flex-direction: column; gap: 12px;
  }
  .chat-msg { display: flex; gap: 8px; align-items: flex-start; }
  .chat-msg.user { flex-direction: row-reverse; }
  .msg-bubble {
    max-width: 80%; padding: 10px 13px; border-radius: 12px; font-size: 13px; line-height: 1.5;
  }
  .chat-msg.user  .msg-bubble { background: #3D35A8; color: #fff; border-radius: 12px 12px 4px 12px; }
  .chat-msg.ai    .msg-bubble { background: #f4f4f4; color: #1a1a1a; border-radius: 12px 12px 12px 4px; }
  .msg-avatar {
    width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center;
    justify-content: center; font-size: 12px; flex-shrink: 0;
  }
  .msg-avatar.ai-av { background: linear-gradient(135deg, #3D35A8, #1D9E75); color: #fff; }
  .msg-avatar.user-av { background: #e0e0e0; color: #555; }

  .chat-suggestions { display: flex; flex-wrap: wrap; gap: 6px; padding: 0 16px 10px; }
  .chat-suggestion {
    font-size: 11px; padding: 5px 10px; border-radius: 20px;
    border: 1px solid #e0e0e0; background: #fff; cursor: pointer; color: #3D35A8;
    font-family: inherit; transition: all 0.15s;
  }
  .chat-suggestion:hover { background: #EEEDFE; border-color: #3D35A8; }

  .chat-input-row {
    padding: 12px 16px; border-top: 1px solid rgba(0,0,0,0.08);
    display: flex; gap: 8px; align-items: flex-end;
  }
  .chat-input {
    flex: 1; padding: 9px 12px; border: 1px solid #e0e0e0; border-radius: 20px;
    font-family: inherit; font-size: 13px; resize: none; outline: none; max-height: 100px;
  }
  .chat-input:focus { border-color: #3D35A8; }
  .chat-send {
    width: 36px; height: 36px; border-radius: 50%; background: #3D35A8; border: none;
    color: #fff; cursor: pointer; font-size: 16px; display: flex; align-items: center;
    justify-content: center; flex-shrink: 0; transition: background 0.15s;
  }
  .chat-send:hover { background: #2d279a; }

  /* ── AI Tools Panel ── */
  .ai-tools { padding: 16px; overflow-y: auto; flex: 1; }
  .ai-tool-card {
    border: 1px solid rgba(0,0,0,0.08); border-radius: 10px; padding: 14px 16px;
    margin-bottom: 12px; cursor: pointer; transition: all 0.15s;
  }
  .ai-tool-card:hover { border-color: #3D35A8; background: #EEEDFE; }
  .ai-tool-icon { font-size: 20px; margin-bottom: 6px; }
  .ai-tool-title { font-size: 13px; font-weight: 600; margin-bottom: 4px; }
  .ai-tool-desc  { font-size: 12px; color: #888; line-height: 1.5; }

  .ai-result {
    background: #f9f9f9; border-radius: 10px; padding: 14px; margin-top: 12px;
    font-size: 13px; line-height: 1.6; white-space: pre-wrap; display: none;
  }
  .ai-result.show { display: block; }
  .ai-score { font-size: 28px; font-weight: 300; color: #3D35A8; }

  .ai-typing { display: flex; gap: 4px; align-items: center; padding: 10px 13px; }
  .ai-typing span { width: 6px; height: 6px; background: #aaa; border-radius: 50%; animation: bounce 1.2s infinite; }
  .ai-typing span:nth-child(2) { animation-delay: 0.2s; }
  .ai-typing span:nth-child(3) { animation-delay: 0.4s; }
  @keyframes bounce { 0%,60%,100%{transform:translateY(0)} 30%{transform:translateY(-6px)} }

  .score-bar { height: 6px; border-radius: 3px; background: #eee; margin: 4px 0 10px; overflow: hidden; }
  .score-fill { height: 100%; border-radius: 3px; background: linear-gradient(90deg, #1D9E75, #3D35A8); transition: width 0.8s ease; }
  .rating-row { display: flex; justify-content: space-between; font-size: 12px; margin: 4px 0; }
  .stars { color: #BA7517; }
</style>

<!-- FAB -->
<button class="ai-fab" onclick="toggleAiDrawer()" title="HR AI Assistant">✦</button>

<!-- Drawer -->
<div class="ai-drawer" id="ai-drawer">
  <div class="ai-drawer-header">
    <div>
      <div class="ai-drawer-title">✦ HR AI Assistant</div>
      <div class="ai-drawer-sub">Powered by Claude</div>
    </div>
    <button class="ai-close" onclick="toggleAiDrawer()">×</button>
  </div>

  <!-- Tabs -->
  <div class="ai-tabs">
    <button class="ai-tab active" onclick="switchAiTab('chat')">💬 Chat</button>
    <button class="ai-tab" onclick="switchAiTab('cv')">📄 Phân tích CV</button>
    <button class="ai-tab" onclick="switchAiTab('risk')">⚠️ Rủi ro</button>
    <button class="ai-tab" onclick="switchAiTab('jd')">📝 Tạo JD</button>
  </div>

  <!-- ── TAB CHAT ── -->
  <div id="ai-tab-chat" style="display:flex;flex-direction:column;flex:1;overflow:hidden;">
    <div class="chat-messages" id="chat-messages">
      <div class="chat-msg ai">
        <div class="msg-avatar ai-av">✦</div>
        <div class="msg-bubble">Xin chào! Tôi là trợ lý HR của bạn. Hỏi tôi bất cứ điều gì về nhân sự, tuyển dụng, hay dữ liệu hệ thống nhé!</div>
      </div>
    </div>
    <div class="chat-suggestions" id="chat-suggestions">
      <button class="chat-suggestion" onclick="sendSuggestion(this)">Tóm tắt tình hình nhân sự</button>
      <button class="chat-suggestion" onclick="sendSuggestion(this)">Ứng viên nào đang thử việc?</button>
      <button class="chat-suggestion" onclick="sendSuggestion(this)">Sinh nhật tuần này?</button>
    </div>
    <div class="chat-input-row">
      <textarea class="chat-input" id="chat-input" placeholder="Hỏi về nhân sự, ứng viên..." rows="1"
        onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();sendChat();}"></textarea>
      <button class="chat-send" onclick="sendChat()">➤</button>
    </div>
  </div>

  <!-- ── TAB CV ── -->
  <div id="ai-tab-cv" class="ai-tools" style="display:none;">
    <div style="font-size:13px;color:#555;margin-bottom:14px;">Nhập thông tin để AI phân tích độ phù hợp của CV với JD.</div>
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">ID ứng viên</label>
      <input class="form-control" id="ai-cv-uvid" placeholder="UV-20260427-xxxx">
    </div>
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">Mô tả công việc (JD)</label>
      <textarea class="form-control" id="ai-cv-jd" rows="4" placeholder="Yêu cầu vị trí, kỹ năng cần thiết..."></textarea>
    </div>
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">Nội dung CV (paste text)</label>
      <textarea class="form-control" id="ai-cv-text" rows="4" placeholder="Paste nội dung CV vào đây..."></textarea>
    </div>
    <button class="btn btn-primary" style="width:100%" onclick="runCvAnalysis()">✦ Phân tích CV</button>
    <div class="ai-result" id="cv-result"></div>
  </div>

  <!-- ── TAB RISK ── -->
  <div id="ai-tab-risk" class="ai-tools" style="display:none;">
    <div class="ai-tool-card" onclick="runTurnoverAnalysis()">
      <div class="ai-tool-icon">⚠️</div>
      <div class="ai-tool-title">Phân tích rủi ro Turnover</div>
      <div class="ai-tool-desc">AI quét toàn bộ nhân sự, phát hiện nhân viên có nguy cơ nghỉ việc cao và đề xuất hành động giữ chân.</div>
    </div>
    <div class="ai-result" id="risk-result"></div>
  </div>

  <!-- ── TAB JD ── -->
  <div id="ai-tab-jd" class="ai-tools" style="display:none;">
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">Vị trí cần tuyển *</label>
      <input class="form-control" id="ai-jd-vitri" placeholder="Sales Executive, Backend Developer...">
    </div>
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">Bộ phận</label>
      <input class="form-control" id="ai-jd-bophan" placeholder="Sales, Tech, Marketing...">
    </div>
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">Yêu cầu cơ bản</label>
      <textarea class="form-control" id="ai-jd-yc" rows="3" placeholder="2 năm kinh nghiệm, biết Excel, tiếng Anh B2..."></textarea>
    </div>
    <div class="form-group" style="margin-bottom:10px;">
      <label class="form-label">Mức lương</label>
      <input class="form-control" id="ai-jd-luong" placeholder="15-20 triệu, Thỏa thuận...">
    </div>
    <button class="btn btn-primary" style="width:100%" onclick="runGenerateJd()">✦ Tạo JD</button>
    <div class="ai-result" id="jd-result"></div>
  </div>
</div>

<script>
// ── AI DRAWER ──
let aiDrawerOpen = false;
function toggleAiDrawer() {
  aiDrawerOpen = !aiDrawerOpen;
  document.getElementById('ai-drawer').classList.toggle('open', aiDrawerOpen);
}

function switchAiTab(tab) {
  ['chat','cv','risk','jd'].forEach(t => {
    document.getElementById('ai-tab-'+t).style.display = t===tab ? 'flex' : 'none';
    if(t==='chat' && tab!=='chat') document.getElementById('ai-tab-'+t).style.display='none';
  });
  document.querySelectorAll('.ai-tab').forEach((btn,i) => {
    btn.classList.toggle('active', ['chat','cv','risk','jd'][i]===tab);
  });
  if(tab==='chat') document.getElementById('ai-tab-chat').style.display='flex';
}

// ── CHAT ──
let chatHistory = [];

function sendSuggestion(btn) { 
  document.getElementById('chat-input').value = btn.textContent; 
  sendChat(); 
}

function sendChat() {
  const input = document.getElementById('chat-input');
  const msg   = input.value.trim();
  if(!msg) return;
  input.value = '';

  appendMessage('user', msg);
  appendTyping();
  document.getElementById('chat-suggestions').innerHTML = '';

  chatHistory.push({ role:'user', content: msg });

  google.script.run.withSuccessHandler(res => {
    removeTyping();
    if(!res.success) { appendMessage('ai','❌ ' + res.error); return; }
    const d = res.data;
    appendMessage('ai', d.reply || d);
    chatHistory.push({ role:'assistant', content: d.reply || String(d) });
    if(chatHistory.length > 12) chatHistory = chatHistory.slice(-12);

    // Câu gợi ý tiếp theo
    if(d.suggestions && d.suggestions.length) {
      const el = document.getElementById('chat-suggestions');
      el.innerHTML = d.suggestions.map(s =>
        \`<button class="chat-suggestion" onclick="sendSuggestion(this)">\${s}</button>\`
      ).join('');
    }
  }).handleRequest({ action:'AI_CHAT', payload:{ message:msg, history:chatHistory.slice(0,-1) } });
}

function appendMessage(role, text) {
  const el   = document.getElementById('chat-messages');
  const div  = document.createElement('div');
  div.className = 'chat-msg ' + role;
  div.innerHTML = \`
    <div class="msg-avatar \${role==='ai'?'ai-av':'user-av'}">\${role==='ai'?'✦':'U'}</div>
    <div class="msg-bubble">\${text.replace(/\\n/g,'<br>')}</div>\`;
  el.appendChild(div);
  el.scrollTop = el.scrollHeight;
}

function appendTyping() {
  const el  = document.getElementById('chat-messages');
  const div = document.createElement('div');
  div.className = 'chat-msg ai'; div.id = 'typing-indicator';
  div.innerHTML = \`<div class="msg-avatar ai-av">✦</div>
    <div class="msg-bubble"><div class="ai-typing"><span></span><span></span><span></span></div></div>\`;
  el.appendChild(div);
  el.scrollTop = el.scrollHeight;
}

function removeTyping() {
  const el = document.getElementById('typing-indicator');
  if(el) el.remove();
}

// ── CV ANALYSIS ──
function runCvAnalysis() {
  const uvId   = document.getElementById('ai-cv-uvid').value.trim();
  const jdText = document.getElementById('ai-cv-jd').value.trim();
  const cvText = document.getElementById('ai-cv-text').value.trim();
  if(!uvId || !jdText || !cvText) { alert('Vui lòng điền đủ thông tin'); return; }

  const btn = event.target;
  btn.textContent = '⏳ Đang phân tích...'; btn.disabled = true;

  google.script.run.withSuccessHandler(res => {
    btn.textContent = '✦ Phân tích CV'; btn.disabled = false;
    const el = document.getElementById('cv-result');
    el.classList.add('show');
    if(!res.success) { el.textContent = '❌ ' + res.error; return; }
    const d = res.data;
    const scoreColor = d.score>=75?'#1D9E75':d.score>=50?'#BA7517':'#C0392B';
    el.innerHTML = \`
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">
        <div class="ai-score" style="color:\${scoreColor}">\${d.score}</div>
        <div><div style="font-size:11px;color:#888">Điểm phù hợp</div>
          <div style="font-size:13px;font-weight:600;color:\${scoreColor}">\${d.recommendation}</div></div>
      </div>
      <div class="score-bar"><div class="score-fill" style="width:\${d.score}%"></div></div>
      <div style="font-size:12px;color:#555;margin-bottom:10px;">\${d.summary}</div>
      <div style="font-size:11px;font-weight:600;color:#1D9E75;margin-bottom:4px;">✅ ĐIỂM MẠNH</div>
      \${(d.strengths||[]).map(s=>\`<div style="font-size:12px;padding:2px 0">• \${s}</div>\`).join('')}
      \${d.gaps&&d.gaps.length?\`<div style="font-size:11px;font-weight:600;color:#C0392B;margin:8px 0 4px;">⚠️ THIẾU SÓT</div>
      \${d.gaps.map(g=>\`<div style="font-size:12px;padding:2px 0">• \${g}</div>\`).join('')}\`:''}
      \${d.suggested_questions&&d.suggested_questions.length?\`<div style="font-size:11px;font-weight:600;color:#3D35A8;margin:8px 0 4px;">💬 CÂU HỎI PV GỢI Ý</div>
      \${d.suggested_questions.map((q,i)=>\`<div style="font-size:12px;padding:2px 0">\${i+1}. \${q}</div>\`).join('')}\`:''}
    \`;
  }).handleRequest({ action:'AI_ANALYZE_CV', payload:{ uvId, jdText, cvText } });
}

// ── TURNOVER RISK ──
function runTurnoverAnalysis() {
  const el  = document.getElementById('risk-result');
  el.classList.add('show');
  el.innerHTML = '⏳ AI đang phân tích toàn bộ nhân sự...';

  google.script.run.withSuccessHandler(res => {
    if(!res.success) { el.textContent='❌ '+res.error; return; }
    const d = res.data;
    const s = res.raw_stats || {};
    const riskColors = { Low:'#1D9E75', Medium:'#BA7517', High:'#C0392B', Critical:'#8B0000' };
    const rc = riskColors[d.turnover_risk_level] || '#555';
    el.innerHTML = \`
      <div style="font-size:18px;font-weight:300;color:\${rc};margin-bottom:6px;">
        Mức rủi ro: <strong>\${d.turnover_risk_level}</strong>
      </div>
      <div style="font-size:12px;color:#555;margin-bottom:12px;">\${d.risk_summary}</div>
      <div style="font-size:11px;font-weight:600;color:#333;margin-bottom:6px;">🚨 HÀNH ĐỘNG NGAY</div>
      \${(d.immediate_actions||[]).map(a=>\`
        <div style="background:#fff;border:1px solid #eee;border-radius:6px;padding:8px 10px;margin-bottom:6px;">
          <div style="font-size:12px;font-weight:500">\${a.action}</div>
          <div style="font-size:11px;color:#888">Priority: \${a.priority} | \${a.target}</div>
        </div>\`).join('')}
      <div style="font-size:11px;font-weight:600;color:#333;margin:10px 0 6px;">💡 CHIẾN LƯỢC GIỮ CHÂN</div>
      \${(d.retention_strategies||[]).map(s=>\`<div style="font-size:12px;padding:2px 0">• \${s}</div>\`).join('')}
    \`;
  }).handleRequest({ action:'AI_TURNOVER_RISK', payload:{} });
}

// ── GENERATE JD ──
function runGenerateJd() {
  const vitri  = document.getElementById('ai-jd-vitri').value.trim();
  if(!vitri) { alert('Vui lòng nhập vị trí cần tuyển'); return; }

  const btn = event.target;
  btn.textContent = '⏳ Đang tạo JD...'; btn.disabled = true;

  const payload = {
    vitri,
    bophan:          document.getElementById('ai-jd-bophan').value.trim(),
    yeu_cau_co_ban:  document.getElementById('ai-jd-yc').value.trim(),
    muc_luong:       document.getElementById('ai-jd-luong').value.trim(),
  };

  google.script.run.withSuccessHandler(res => {
    btn.textContent = '✦ Tạo JD'; btn.disabled = false;
    const el = document.getElementById('jd-result');
    el.classList.add('show');
    if(!res.success) { el.textContent='❌ '+res.error; return; }
    const d = res.data;
    el.innerHTML = \`
      <div style="font-size:14px;font-weight:600;margin-bottom:8px;">\${d.title}</div>
      <div style="font-size:12px;color:#555;margin-bottom:10px;">\${d.overview}</div>
      <div style="font-size:11px;font-weight:600;margin-bottom:4px;">TRÁCH NHIỆM</div>
      \${(d.responsibilities||[]).map(r=>\`<div style="font-size:12px;padding:2px 0">• \${r}</div>\`).join('')}
      <div style="font-size:11px;font-weight:600;margin:8px 0 4px;">YÊU CẦU</div>
      \${(d.requirements||[]).map(r=>\`<div style="font-size:12px;padding:2px 0">• \${r}</div>\`).join('')}
      \${d.nice_to_have&&d.nice_to_have.length?\`<div style="font-size:11px;font-weight:600;margin:8px 0 4px;color:#3D35A8;">ƯU TIÊN</div>
      \${d.nice_to_have.map(r=>\`<div style="font-size:12px;padding:2px 0">• \${r}</div>\`).join('')}\`:''}
      <div style="font-size:11px;font-weight:600;margin:8px 0 4px;">QUYỀN LỢI</div>
      \${(d.benefits||[]).map(r=>\`<div style="font-size:12px;padding:2px 0">• \${r}</div>\`).join('')}
      <div style="margin-top:10px;padding:8px 10px;background:#EEEDFE;border-radius:6px;font-size:12px;">
        💰 Mức lương: <strong>\${d.salary}</strong>
      </div>
      <button class="btn btn-ghost btn-sm" style="margin-top:10px;width:100%" onclick="copyJd()">📋 Copy JD</button>
    \`;
    window._lastJd = d;
  }).handleRequest({ action:'AI_GENERATE_JD', payload });
}

function copyJd() {
  if(!window._lastJd) return;
  const d = window._lastJd;
  const text = \`\${d.title}\\n\\n\${d.overview}\\n\\nTrách nhiệm:\\n\${d.responsibilities.map(r=>'• '+r).join('\\n')}\\n\\nYêu cầu:\\n\${d.requirements.map(r=>'• '+r).join('\\n')}\\n\\nMức lương: \${d.salary}\`;
  navigator.clipboard.writeText(text).then(()=>{ toast('Đã copy JD!','success'); }).catch(()=>{});
}
</script>`;
}

// =============================================================================
// MODULE 20 — HELPER: Thêm AI Panel vào buildFullApp()
// Thêm dòng sau vào cuối buildFullApp(), trước </body>:
//   ${buildAiPanel()}
// =============================================================================

/**
 * Phiên bản buildFullApp() đã tích hợp AI panel
 * Gọi hàm này THAY cho buildFullApp() trong doGet()
 */
function buildFullAppWithAi(user, activePage) {
  const base = buildFullApp(user, activePage);
  const aiPanel = buildAiPanel();
  // Chèn AI panel trước </body>
  return base.replace('</body>', aiPanel + '\n</body>');
}


// =============================================================================
// MODULE 21 — SETUP AI (chạy 1 lần)
// =============================================================================

/**
 * Chạy hàm này để cài đặt AI:
 * 1. Set API key
 * 2. Tạo sheet AI_RESULTS
 * 3. Đặt trigger phân tích turnover hàng tuần
 */
function setupAi() {
  // ⚠️ THAY "sk-ant-..." BẰNG API KEY THẬT CỦA BẠN
  // Lấy tại: https://console.anthropic.com/
  // setApiKey('sk-ant-api03-...');

  // Tạo sheet AI_RESULTS
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('AI_RESULTS')) {
    const sheet = ss.insertSheet('AI_RESULTS');
    const headers = ['ID', 'Entity_ID', 'Analysis_Type', 'Score', 'Result_JSON', 'Created_At'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2C2C2A').setFontColor('#FFF').setFontWeight('bold');
    Logger.log('✅ Đã tạo sheet AI_RESULTS');
  }

  // Trigger phân tích turnover mỗi thứ 2
  ScriptApp.newTrigger('analyzeTurnoverRisk')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('✅ AI setup hoàn tất! Nhớ chạy setApiKey("sk-ant-...") để set API key.');
  SpreadsheetApp.getUi().alert(
    '✅ Giai đoạn 3 AI đã được cài đặt!\n\n' +
    'Bước tiếp theo:\n' +
    '1. Lấy API key tại console.anthropic.com\n' +
    '2. Chạy: setApiKey("sk-ant-api03-...")\n' +
    '3. Đổi doGet() để dùng buildFullAppWithAi() thay vì buildFullApp()\n\n' +
    'Tính năng AI:\n' +
    '• Chatbot HR hỏi đáp tiếng Việt\n' +
    '• Phân tích CV & chấm điểm tự động\n' +
    '• Phân tích rủi ro turnover hàng tuần\n' +
    '• Tạo JD tự động\n' +
    '• Tóm tắt kết quả phỏng vấn'
  );
}
