// ============================================================
// ZMA ARTCC TRAINING PORTAL — CONFIG
// ============================================================

const CONFIG = {
  SUPABASE_URL: 'https://rnfvvvxsesxcmejonndyu.supabase.co',
  SUPABASE_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJuZnZ2dnhzZXN4Y21lam9uZHl1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNTY1MjAsImV4cCI6MjA4ODgzMjUyMH0.Az7plNF6vX4aWho-hc7eqjQdP1TdCyYVfaIuxEPdr_8',
  VATUSA_API: 'https://api.vatusa.net/v2',
  VATUSA_API_KEY: 'UAobqIiRBtq2cOnu',
  FACILITY: 'ZMA',
  EMAILJS_PUBLIC_KEY: 'vkV3HRSeyW9Tz34t7',
  EMAILJS_SERVICE_ID: 'service_zma', // UPDATE with your actual EmailJS Service ID
  EMAILJS_TEMPLATE_TRAINING_CONFIRM: 'template_rbuk3tw',
  EMAILJS_TEMPLATE_REJECTION: 'template_rbuk3tw',     // UPDATE when you create this template
  EMAILJS_TEMPLATE_COMPLETION: 'template_rbuk3tw',    // UPDATE when you create this template
  EMAILJS_TEMPLATE_EXAM_APPROVED: 'template_rbuk3tw', // UPDATE when you create this template
  EMAILJS_TEMPLATE_EXAM_REJECTED: 'template_rbuk3tw', // UPDATE when you create this template
  SUPPORT_EMAIL: 'ta@zmaartcc.net',
};

// Training types matching the screenshot
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
// Supabase REST client (no library needed)
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
    // Polling fallback for realtime (no websocket lib needed)
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

  async getControllerEmail(cid) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/user/${cid}/email?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return null;
      const data = await res.json();
      return data.email || null;
    } catch { return null; }
  },

  async getFacilityRoster(facility) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/facility/${facility}/roster?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return null;
      const data = await res.json();
      return data.data || [];
    } catch { return []; }
  },

  async getFacilityStaff(facility) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/facility/${facility}/staff?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return null;
      const data = await res.json();
      return data.data || [];
    } catch { return []; }
  },

  async isMemberOfFacility(cid, facility) {
    const ctrl = await this.getController(cid);
    if (!ctrl) return false;
    // Check home facility or visiting
    if (ctrl.facility === facility) return { type: 'home', controller: ctrl };
    if (ctrl.visiting_facilities && ctrl.visiting_facilities.some(f => f.facility === facility)) {
      return { type: 'visitor', controller: ctrl };
    }
    return false;
  },

  async getExamGrades(cid) {
    try {
      const res = await fetch(`${CONFIG.VATUSA_API}/user/${cid}/exam?apikey=${CONFIG.VATUSA_API_KEY}`);
      if (!res.ok) return [];
      const data = await res.json();
      return data.data || [];
    } catch { return []; }
  },

  ratingName(rating) {
    const map = { 1: 'OBS', 2: 'S1', 3: 'S2', 4: 'S3', 5: 'C1', 6: 'C3', 7: 'I1', 8: 'I3', 9: 'SUP', 10: 'ADM' };
    return map[rating] || rating;
  },
};

// ============================================================
// EmailJS helper
// ============================================================
const EM = {
  async send(templateId, params) {
    try {
      await emailjs.send(CONFIG.EMAILJS_SERVICE_ID, templateId, params, CONFIG.EMAILJS_PUBLIC_KEY);
      return true;
    } catch (e) {
      console.error('EmailJS error:', e);
      return false;
    }
  },

  async sendTrainingConfirmation(to, name, trainingType) {
    return this.send(CONFIG.EMAILJS_TEMPLATE_TRAINING_CONFIRM, {
      to_email: to, to_name: name, training_type: trainingType,
      message: `Thank you for submitting your training request for ${trainingType}. Your request is pending approval. You will receive another email once your request has been reviewed.`,
    });
  },

  async sendRejection(to, name, reason) {
    return this.send(CONFIG.EMAILJS_TEMPLATE_REJECTION, {
      to_email: to, to_name: name,
      message: reason || `[PLACEHOLDER — Update your rejection email template in EmailJS]`,
    });
  },

  async sendCompletion(to, name, trainingType) {
    return this.send(CONFIG.EMAILJS_TEMPLATE_COMPLETION, {
      to_email: to, to_name: name, training_type: trainingType,
      message: `[PLACEHOLDER — Update your completion email template in EmailJS] Congratulations on completing your ${trainingType} training!`,
    });
  },

  async sendExamApproved(to, name, examTitle, password) {
    return this.send(CONFIG.EMAILJS_TEMPLATE_EXAM_APPROVED, {
      to_email: to, to_name: name, exam_title: examTitle, exam_password: password,
      message: `[PLACEHOLDER — Update your exam approval template in EmailJS] Your exam password for ${examTitle} is: ${password}`,
    });
  },

  async sendExamRejected(to, name, examTitle, reason) {
    return this.send(CONFIG.EMAILJS_TEMPLATE_EXAM_REJECTED, {
      to_email: to, to_name: name, exam_title: examTitle,
      message: reason || `[PLACEHOLDER — Update your exam rejection template in EmailJS]`,
    });
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
