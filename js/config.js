// ============================================================
// ZMA ARTCC TRAINING PORTAL — CONFIG
// ============================================================

const CONFIG = {
  SUPABASE_URL: 'https://yprhefpaozkuopjcuwas.supabase.co',
  SUPABASE_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlwcmhlZnBhb3prdW9wamN1d2FzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcwNzc1NDEsImV4cCI6MjA5MjY1MzU0MX0.8TGY_D_qE5XdogEBOKndA7LiomXt_MRGejDfJxyu5Os',
  VATUSA_API: 'https://api.vatusa.net/v2',
  VATUSA_API_KEY: 'UAobqIiRBtq2cOnu',
  FACILITY: 'ZMA',
  EMAILJS_PUBLIC_KEY: 'vkV3HRSeyW9Tz34t7',
  EMAILJS_SERVICE_ID: 'service_i98tzkp',
  EMAILJS_TEMPLATE_CONFIRM: 'template_rbuk3tw',  // Submission confirmation — keep as-is
  EMAILJS_TEMPLATE_MASTER:  'template_eg92hs8',  // Master wrapper — all other emails
  SUPPORT_EMAIL: 'ta@zmaartcc.net',
};

// Training types that require an exam password in the approval email
const EXAM_REQUIRED_TYPES = [
  'S1 Training (Tier 2)',
  'S2 Training (Tier 2)',
  'S3 Training (Tier 2)',
  'C1 Training',
  'Tier 1 MIA GND/DEL',
  'Tier 1 MIA LCL',
  'Tier 1 MIA TRACON',
  'Visitor/Transfer — Miami Center Domestic',
  'Visitor/Transfer — Miami CAB',
];

const TRAINING_TYPES = {
  'S1 Training': ['S1 Training (Tier 2)', 'S1 Training (Unrestricted)'],
  'S2 Training': ['S2 Training (Tier 2)', 'S2 Training (Unrestricted)'],
  'S3 Training': ['S3 Training (Tier 2)', 'S3 Training (Unrestricted)'],
  'C1 Training': ['C1 Training'],
  'Center & Oceanic': ['Miami Center Oceanic'],
  'Tier 1': ['Tier 1 MIA GND/DEL', 'Tier 1 MIA LCL', 'Tier 1 MIA TRACON'],
  'Visitor / Transfer': ['Visitor/Transfer — Miami Center Domestic', 'Visitor/Transfer — Miami CAB'],
};

const CERTIFICATIONS = ['S1', 'S2', 'S3', 'C1', 'T1S1', 'T1S2', 'T1S3', 'ZMO'];

const EXAM_CATEGORIES = [
  { id: 'tier2', label: 'Tier 2 Exams' },
  { id: 'tier1', label: 'Tier 1 Exams' },
  { id: 'training', label: 'Training Exams' },
];

// ============================================================
// Supabase REST client
// ============================================================
const SB = {
  headers: {
    'apikey': CONFIG.SUPABASE_KEY,
    'Authorization': `Bearer ${CONFIG.SUPABASE_KEY}`,
    'Content-Type': 'application/json',
    'Prefer': 'return=representation',
  },

  async query(table, params = '') {
    const res = await fetch(`${CONFIG.SUPABASE_URL}/rest/v1/${table}${params}`, { headers: this.headers });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  },

  async insert(table, data) {
    const res = await fetch(`${CONFIG.SUPABASE_URL}/rest/v1/${table}`, {
      method: 'POST',
      headers: this.headers,
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  },

  async update(table, filter, data) {
    const res = await fetch(`${CONFIG.SUPABASE_URL}/rest/v1/${table}?${filter}`, {
      method: 'PATCH',
      headers: this.headers,
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error(await res.text());
    return res.json();
  },

  async delete(table, filter) {
    const res = await fetch(`${CONFIG.SUPABASE_URL}/rest/v1/${table}?${filter}`, {
      method: 'DELETE',
      headers: this.headers,
    });
    if (!res.ok) throw new Error(await res.text());
    return res.status === 204 ? true : res.json();
  },

  subscribe(table, callback) {
    return setInterval(callback, 5000);
  },
};

// ============================================================
// VATUSA API helpers
// ============================================================
const VATUSA = {
  async getController(cid) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/user/${cid}?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return null;
      const data = await res.json();
      return data.data || null;
    } catch { return null; }
  },

  async getFacilityRoster(facility) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/facility/${facility}/roster?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return [];
      const data = await res.json();
      return data.data || [];
    } catch { return []; }
  },

  async getFacilityStaff(facility) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/facility/${facility}/staff?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return [];
      const data = await res.json();
      return data.data || [];
    } catch { return []; }
  },

  async isMemberOfFacility(cid, facility) {
    const ctrl = await this.getController(cid);
    if (!ctrl) return false;
    if (ctrl.facility === facility) return { type: 'home', controller: ctrl };
    if (ctrl.visiting_facilities && ctrl.visiting_facilities.some(f => f.facility === facility)) {
      return { type: 'visitor', controller: ctrl };
    }
    return false;
  },

  async getExamGrades(cid) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/academy/transcript/${cid}?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return [];
      const data = await res.json();
      // Transcript returns array of course objects with name, passed, score etc.
      return data.data || data || [];
    } catch { return []; }
  },

  ratingName(rating) {
    const map = { 1:'OBS', 2:'S1', 3:'S2', 4:'S3', 5:'C1', 6:'C2', 7:'C3', 8:'I1', 9:'I3', 10:'SUP', 11:'ADM' };
    return map[rating] || rating;
  },
};

// ============================================================
// Email helpers
// ============================================================
const EM = {

  // Shared professional shell — injected into the master template
  _wrap(headerText, bodyHtml) {
    return `
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;padding:40px 20px;">
        <tr><td align="center">
          <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">

            <tr>
              <td style="background:#0a1628;border-radius:12px 12px 0 0;padding:32px 40px;text-align:center;">
                <span style="display:inline-block;background:#1e4d8c;color:#fff;font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;padding:5px 16px;border-radius:100px;margin-bottom:16px;">ZMA Miami ARTCC</span>
                <div style="color:#fff;font-size:22px;font-weight:700;letter-spacing:-0.5px;">${headerText}</div>
              </td>
            </tr>

            <tr>
              <td style="background:#fff;padding:40px;">
                ${bodyHtml}
              </td>
            </tr>

            <tr>
              <td style="background:#f8fafc;border-top:1px solid #e2e8f0;border-radius:0 0 12px 12px;padding:24px 40px;text-align:center;">
                <p style="margin:0;font-size:12px;color:#94a3b8;">This is an automated message from the ZMA ARTCC Training Portal.<br>Please do not reply directly to this email.</p>
              </td>
            </tr>

          </table>
        </td></tr>
      </table>`;
  },

  // Low-level send via master template
  async _send(toEmail, subject, headerText, bodyHtml) {
    try {
      await emailjs.send(
        CONFIG.EMAILJS_SERVICE_ID,
        CONFIG.EMAILJS_TEMPLATE_MASTER,
        {
          to_email: toEmail,
          subject: subject,
          message_body: this._wrap(headerText, bodyHtml),
        },
        CONFIG.EMAILJS_PUBLIC_KEY
      );
      return true;
    } catch(e) { console.error('Email error:', e); return false; }
  },

  // ── 1. Submission confirmation (uses its own template) ───
  async sendTrainingConfirmation(toEmail, toName, trainingType) {
    try {
      await emailjs.send(CONFIG.EMAILJS_SERVICE_ID, CONFIG.EMAILJS_TEMPLATE_CONFIRM, {
        to_email: toEmail,
        to_name: toName,
        training_type: trainingType,
        message: `Your request is pending review. You will be notified once it has been processed.`,
      }, CONFIG.EMAILJS_PUBLIC_KEY);
      return true;
    } catch(e) { console.error('Confirmation email error:', e); return false; }
  },

  // ── 2. Approval WITH exam password ──────────────────────
  async sendApprovalWithExam(toEmail, toName, trainingType, examPassword) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        Thank you again for submitting a training request for <strong>${trainingType}</strong>. The ZMA Training Team is working hard to work through the training queue as fast as possible while upholding our high training quality standards. We thank you in advance for your patience.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        While you are waiting for training, we ask that you begin studying. Per VATUSA policy, students are required to be fully prepared for training sessions by having at a minimum a baseline understanding of all concepts, phraseology, and procedures. This includes studying your
        <a href="https://docs.google.com/spreadsheets/d/1jNdQUqyKbeZa6OU1Ls02p0aRNfrI2TN6rE1SxlhE9Wo/edit?gid=1564061930#gid=1564061930" style="color:#1e4d8c;font-weight:600;">ratings competencies</a>
        and our local SOPs found on the ZMA website. This is also a fantastic opportunity to observe the network, meet members of the Miami ARTCC, and ask any questions you may have.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#eff6ff;border:1px solid #bfdbf7;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:6px;">Your Training Program</div>
          <div style="font-size:16px;font-weight:600;color:#0a1628;">${trainingType}</div>
        </td></tr>
      </table>

      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        In addition to studying, you are also required to complete your rating exam prior to your training start — completing this exam as soon as possible is recommended so as not to delay your training. The exam is completed on the ZMA Academy website.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#16a34a;margin-bottom:6px;">Your Exam Password</div>
          <div style="font-size:26px;font-weight:700;color:#0a1628;letter-spacing:3px;">${examPassword}</div>
        </td></tr>
      </table>

      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Please email <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> or ask in the Miami ARTCC Discord should you have any questions. Thank you for your patience and we greatly look forward to training with you!
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Training Request Approved', 'Training Request Approved', body);
  },

  // ── 3. Approval WITHOUT exam password ───────────────────
  async sendApprovalNoExam(toEmail, toName, trainingType) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        Thank you again for submitting a training request for <strong>${trainingType}</strong>. The ZMA Training Team is working hard to work through the training queue as fast as possible while upholding our high training quality standards. We thank you in advance for your patience.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        While you are waiting for training, we ask that you begin studying. Per VATUSA policy, students are required to be fully prepared for training sessions by having at a minimum a baseline understanding of all concepts, phraseology, and procedures. This includes studying your
        <a href="https://docs.google.com/spreadsheets/d/1jNdQUqyKbeZa6OU1Ls02p0aRNfrI2TN6rE1SxlhE9Wo/edit?gid=1564061930#gid=1564061930" style="color:#1e4d8c;font-weight:600;">ratings competencies</a>
        and our local SOPs found on the ZMA website. This is also a fantastic opportunity to observe the network, meet members of the Miami ARTCC, and ask any questions you may have.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#eff6ff;border:1px solid #bfdbf7;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:6px;">Your Training Program</div>
          <div style="font-size:16px;font-weight:600;color:#0a1628;">${trainingType}</div>
        </td></tr>
      </table>

      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Please email <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> or ask in the Miami ARTCC Discord should you have any questions. Thank you for your patience and we greatly look forward to training with you!
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Training Request Approved', 'Training Request Approved', body);
  },

  // ── 4. Rejection ────────────────────────────────────────
  async sendRejection(toEmail, toName, trainingType, customMessage) {
    const reason = customMessage ||
      `After reviewing your request, we are unable to accommodate your training request for <strong>${trainingType}</strong> at this time. Please contact us at <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> for more information.`;
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">${reason}</p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Training Request Update', 'Training Request Update', body);
  },

  // ── 5. Training completion ───────────────────────────────
  async sendCompletion(toEmail, toName, trainingType) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Congratulations on completing your <strong>${trainingType}</strong> training with ZMA Miami ARTCC! We are proud of the hard work and dedication you have shown throughout this process.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        If you have any questions about next steps, please reach out at <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> or in the Miami ARTCC Discord.
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Training Completed', 'Training Completed', body);
  },

  // ── 6. Exam password request approved ───────────────────
  async sendExamPasswordApproved(toEmail, toName, examTitle, examPassword) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Your request for the <strong>${examTitle}</strong> exam password has been approved. Please find your exam password below.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#16a34a;margin-bottom:6px;">Exam Password</div>
          <div style="font-size:26px;font-weight:700;color:#0a1628;letter-spacing:3px;">${examPassword}</div>
        </td></tr>
      </table>

      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Please complete your exam on the ZMA Academy website as soon as possible. If you have any questions, contact us at
        <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a>.
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — Exam Password: ${examTitle}`, 'Exam Password Approved', body);
  },

  // ── 7. Exam password request denied ─────────────────────
  async sendExamPasswordDenied(toEmail, toName, examTitle, customMessage) {
    const reason = customMessage ||
      `We are unable to provide the exam password for <strong>${examTitle}</strong> at this time. Please contact us at <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> for more information.`;
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">${reason}</p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — Exam Request Update: ${examTitle}`, 'Exam Request Update', body);
  },

  // ── 8. Trainer assigned — sent to STUDENT ───────────────
  async sendTrainerAssignedStudent(toEmail, studentName, trainerName, trainingType) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${studentName}</strong>,</p>
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        Thank you for your patience! You have been assigned <strong>${trainerName}</strong> as your primary trainer for your <strong>${trainingType}</strong> training.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#eff6ff;border:1px solid #bfdbf7;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="display:flex;gap:32px;flex-wrap:wrap;">
            <div>
              <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:4px;">Assigned Trainer</div>
              <div style="font-size:16px;font-weight:600;color:#0a1628;">${trainerName}</div>
            </div>
            <div>
              <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:4px;">Training Program</div>
              <div style="font-size:16px;font-weight:600;color:#0a1628;">${trainingType}</div>
            </div>
          </div>
        </td></tr>
      </table>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#fefce8;border:1px solid #fde68a;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:16px 20px;">
          <div style="font-size:13px;font-weight:700;color:#92400e;margin-bottom:4px;">⚠ Action Required Within 96 Hours</div>
          <div style="font-size:14px;color:#78350f;line-height:1.6;">A Discord thread will be created within 24 hours. Please provide your availability in the Discord thread within <strong>96 hours</strong>; otherwise, you will be unassigned and removed from the training queue.</div>
        </td></tr>
      </table>

      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        Please also remember that per VATUSA policy, all students must be adequately prepared for all training sessions. Failure to prepare will result in removal from the training queue at the discretion of the TA. If you were required to complete any exams and have not yet done so, please do so as soon as possible.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        We look forward to working with you and are excited to work with you through your certification!
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Trainer Assigned', 'Trainer Assigned', body);
  },

  // ── 9. Trainer assigned — sent to TRAINER ───────────────
  async sendTrainerAssignedTrainer(toEmail, trainerName, studentName, trainingType) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${trainerName}</strong>,</p>
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        You have been assigned as the primary trainer for <strong>${studentName}</strong> for their <strong>${trainingType}</strong> training.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#eff6ff;border:1px solid #bfdbf7;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="display:flex;gap:32px;flex-wrap:wrap;">
            <div>
              <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:4px;">Student</div>
              <div style="font-size:16px;font-weight:600;color:#0a1628;">${studentName}</div>
            </div>
            <div>
              <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:4px;">Training Program</div>
              <div style="font-size:16px;font-weight:600;color:#0a1628;">${trainingType}</div>
            </div>
          </div>
        </td></tr>
      </table>

      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Please create a Discord thread for this student within 24 hours. The student has been asked to provide their availability within 96 hours of trainer assignment.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        If you have any questions, please contact <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a>.
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — New Student Assigned: ${studentName}`, 'New Student Assigned', body);
  },

  // ── 10. Removal — Ineligibility (T01) ───────────────────
  async sendRemovalIneligibility(toEmail, studentName) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${studentName}</strong>,</p>
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        This email is to inform you that you have been removed from the Miami ARTCC Training Queue due to ineligibility per T01.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        If you wish to resume training in the future, you may submit a new training request once you are able to fully commit to the expectations outlined in VATUSA and ZMA training policies.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        If you have questions regarding this decision, you may contact the Training Administrator at <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> for clarification.
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Training Queue Removal', 'Training Queue Removal', body);
  },

  // ── 11. Removal — Inactivity ─────────────────────────────
  async sendRemovalInactivity(toEmail, studentName) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${studentName}</strong>,</p>
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        This email is to inform you that you have been removed from the Miami ARTCC Training Queue due to inactivity.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        If you wish to resume training in the future, you may submit a new training request once you are able to fully commit to the expectations outlined in VATUSA and ZMA training policies.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        If you have questions regarding this decision, you may contact the Training Administrator at <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> for clarification.
      </p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, 'ZMA ARTCC — Training Queue Removal', 'Training Queue Removal', body);
  },

  // ── 12. Training completion (updated) ───────────────────
  async sendCompletion(toEmail, studentName, trainingType) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${studentName}</strong>,</p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#0a1628,#1e4d8c);border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:28px 24px;text-align:center;">
          <div style="font-size:40px;margin-bottom:8px;">🎉</div>
          <div style="color:#fff;font-size:20px;font-weight:700;letter-spacing:-0.3px;">Congratulations!</div>
          <div style="color:#93c5fd;font-size:14px;margin-top:6px;">You have earned your certification</div>
          <div style="display:inline-block;background:rgba(255,255,255,0.15);color:#fff;font-size:15px;font-weight:700;padding:8px 20px;border-radius:100px;margin-top:12px;">${trainingType}</div>
        </td></tr>
      </table>

      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">
        Once you see your certification(s) on the ZMA Website, you are free to utilize your certification on the network.
      </p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Please do not hesitate to contact <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a> should you have any questions or concerns.
      </p>
      <p style="margin:0 0 8px;font-size:15px;color:#334155;line-height:1.6;">Once again, congratulations!!</p>
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — Training Complete: ${trainingType}`, 'Training Complete! 🎉', body);
  },

  // ── 13. Exam password approved (simple) ─────────────────
  async sendExamPasswordApproved(toEmail, toName, examTitle, examPassword) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">The exam password for the <strong>${examTitle}</strong> is:</p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:24px;text-align:center;">
          <div style="font-size:32px;font-weight:700;color:#0a1628;letter-spacing:6px;">${examPassword}</div>
        </td></tr>
      </table>

      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — Exam Password: ${examTitle}`, 'Exam Password', body);
  },

  // ── 14a. Exam password denied ────────────────────────────
  async sendExamRejection(toEmail, toName, examTitle, reason) {
    const body = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        Your exam password request for <strong>${examTitle}</strong> has been denied.
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:20px 24px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#dc2626;margin-bottom:8px;">Reason</div>
          <div style="font-size:15px;color:#334155;line-height:1.6;">${reason}</div>
        </td></tr>
      </table>

      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — Exam Request Denied: ${examTitle}`, 'Exam Request Denied', body);
  },

  // ── 14b. Exam score result ───────────────────────────────
  async sendExamScore(toEmail, toName, examTitle, score, totalQuestions, passed) {
    const pct = Math.round((score / totalQuestions) * 100);
    const passColor = passed ? '#16a34a' : '#dc2626';
    const passBg    = passed ? '#f0fdf4'  : '#fef2f2';
    const passBorder= passed ? '#bbf7d0'  : '#fecaca';
    const passLabel = passed ? 'PASSED'   : 'FAILED';
    const passEmoji = passed ? '✅'        : '❌';

    // Score arc graphic using inline SVG (works in most email clients)
    const arcBody = `
      <p style="margin:0 0 16px;font-size:15px;color:#334155;line-height:1.6;">Hello <strong>${toName}</strong>,</p>
      <p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">
        You have completed the <strong>${examTitle}</strong> exam. Here are your results:
      </p>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:${passBg};border:1px solid ${passBorder};border-radius:12px;margin-bottom:24px;">
        <tr><td style="padding:28px 24px;text-align:center;">
          <div style="font-size:48px;font-weight:800;color:${passColor};line-height:1;">${pct}%</div>
          <div style="font-size:13px;color:#64748b;margin-top:4px;">${score} / ${totalQuestions} correct</div>
          <div style="display:inline-block;background:${passColor};color:#fff;font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;padding:4px 16px;border-radius:100px;margin-top:12px;">${passEmoji} ${passLabel}</div>
        </td></tr>
      </table>

      <table width="100%" cellpadding="0" cellspacing="0" style="background:#eff6ff;border:1px solid #bfdbf7;border-radius:10px;margin-bottom:24px;">
        <tr><td style="padding:16px 20px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2563a8;margin-bottom:4px;">Exam</div>
          <div style="font-size:15px;font-weight:600;color:#0a1628;">${examTitle}</div>
        </td></tr>
      </table>

      ${passed
        ? `<p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">Great work! Your result has been recorded. If you have any questions, contact <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a>.</p>`
        : `<p style="margin:0 0 24px;font-size:15px;color:#334155;line-height:1.6;">Unfortunately you did not pass this attempt. Please review the material and try again. If you have any questions, contact <a href="mailto:ta@zmaartcc.net" style="color:#1e4d8c;font-weight:600;">ta@zmaartcc.net</a>.</p>`
      }
      <p style="margin:0;font-size:15px;color:#334155;line-height:1.6;">Best regards,<br><strong>ZMA Training Team</strong></p>`;

    return this._send(toEmail, `ZMA ARTCC — Exam Results: ${examTitle}`, 'Exam Results', arcBody);
  },
};

// ============================================================
// Shared UI helpers
// ============================================================
function toast(msg, type = 'success') {
  const el = document.createElement('div');
  el.className = `toast toast--${type}`;
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.classList.add('toast--show'), 10);
  setTimeout(() => { el.classList.remove('toast--show'); setTimeout(() => el.remove(), 300); }, 3500);
}

function formatDate(d) {
  if (!d) return '—';
  return new Date(d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function timeSince(d) {
  if (!d) return '—';
  const diff = Date.now() - new Date(d).getTime();
  const days = Math.floor(diff / 86400000);
  if (days === 0) return 'Today';
  if (days === 1) return 'Yesterday';
  return `${days} days ago`;
}
