<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Admin Portal — ZMA ARTCC</title>
  <link rel="stylesheet" href="css/styles.css">
  <script src="https://cdn.jsdelivr.net/npm/@emailjs/browser@3/dist/email.min.js"></script>
  <style>
    .queue-section { margin-bottom: 36px; }
    .queue-title {
      font-size: .8rem; font-weight: 700; letter-spacing: .07em;
      text-transform: uppercase; color: var(--gray-400);
      margin-bottom: 12px; padding-bottom: 8px;
      border-bottom: 1px solid var(--gray-200);
      display: flex; align-items: center; justify-content: space-between;
    }
    .exam-cat-header {
      font-size: .8rem; font-weight: 700; letter-spacing: .07em;
      text-transform: uppercase; color: var(--blue-600);
      background: var(--blue-50); border-radius: var(--radius-sm);
      padding: 6px 12px; margin: 20px 0 12px;
    }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    .trainer-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 16px; }
    .cert-toggle { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
    .cert-chip {
      padding: 4px 10px; border-radius: 4px; font-size: .75rem; font-weight: 600;
      border: 1px solid var(--gray-200); background: var(--white);
      cursor: pointer; transition: all .12s; color: var(--gray-600);
    }
    .cert-chip.selected { background: var(--blue-600); color: var(--white); border-color: var(--blue-600); }
    .action-row { display: flex; gap: 6px; flex-wrap: wrap; }
    .move-btn { font-size: .75rem; padding: 4px 8px; }
  </style>
</head>
<body>
  <nav class="top-nav">
    <a class="nav-brand" href="#">
      <span class="nav-brand-badge">ZMA</span>
      Miami ARTCC Training Portal
    </a>
    <div class="nav-links">
      <a class="nav-link" href="signup.html">Training Signup</a>
      <a class="nav-link" href="exam-center.html">Exam Center</a>
      <a class="nav-link active" href="admin.html">Admin Portal</a>
    </div>
  </nav>

  <div class="page-wrapper">
    <div class="admin-layout">
      <!-- SIDEBAR -->
      <nav class="admin-sidebar">
        <div class="sidebar-section">
          <div class="sidebar-label">Training</div>
          <button class="sidebar-btn active" onclick="switchTab('queues')">Queues</button>
          <button class="sidebar-btn" onclick="switchTab('completed')">Completed</button>
          <button class="sidebar-btn" onclick="switchTab('removed')">Removed</button>
        </div>
        <div class="sidebar-section" style="margin-top:16px;">
          <div class="sidebar-label">Exams</div>
          <button class="sidebar-btn" onclick="switchTab('exam-requests')" id="sidebarExamRequests">
            Exam Requests <span id="examNotifDot" class="notif-dot" style="display:none;">0</span>
          </button>
          <button class="sidebar-btn" onclick="switchTab('exam-center')">Exam Center</button>
        </div>
        <div class="sidebar-section" style="margin-top:16px;">
          <div class="sidebar-label">Settings</div>
          <button class="sidebar-btn" onclick="switchTab('admin')">Admin</button>
        </div>
      </nav>

      <!-- MAIN CONTENT -->
      <div class="admin-content">

        <!-- ======================== QUEUES TAB ======================== -->
        <div class="tab-content active" id="tab-queues">
          <div class="page-header">
            <div class="flex items-center gap-3">
              <h1>Training Queues</h1>
              <button class="btn btn-primary btn-sm" onclick="openAddStudentModal()">+ Add Student</button>
            </div>
          </div>

          <!-- Stats -->
          <div class="stats-grid" id="queueStats"></div>

          <!-- Pending Approval -->
          <div id="pendingSection" style="display:none;" class="queue-section">
            <div class="queue-title">Pending Approval</div>
            <div class="card" style="padding:0;">
              <div class="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Name</th><th>CID</th><th>Training Type</th><th>Requested</th>
                      <th>Exam Grades</th><th>Actions</th>
                    </tr>
                  </thead>
                  <tbody id="pendingBody"></tbody>
                </table>
              </div>
            </div>
          </div>

          <!-- Home Queue -->
          <div class="queue-section">
            <div class="queue-title">Home Controller Queue</div>
            <div class="card" style="padding:0;">
              <div class="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th></th><th>Name</th><th>CID</th><th>Training Type</th>
                      <th>Requested</th><th>Status</th><th>Exam Grades</th>
                      <th>Trainer</th><th>Actions</th>
                    </tr>
                  </thead>
                  <tbody id="homeBody"></tbody>
                </table>
              </div>
            </div>
          </div>

          <!-- Visitor Queue -->
          <div class="queue-section">
            <div class="queue-title">Visitor Queue</div>
            <div class="card" style="padding:0;">
              <div class="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th></th><th>Name</th><th>CID</th><th>Training Type</th>
                      <th>Requested</th><th>Status</th><th>Exam Grades</th>
                      <th>Trainer</th><th>Actions</th>
                    </tr>
                  </thead>
                  <tbody id="visitorBody"></tbody>
                </table>
              </div>
            </div>
          </div>
        </div>

        <!-- ======================== COMPLETED TAB ======================== -->
        <div class="tab-content" id="tab-completed">
          <div class="page-header"><h1>Completed Training</h1></div>
          <div class="card" style="padding:0;">
            <div class="table-wrap">
              <table>
                <thead>
                  <tr><th>Name</th><th>CID</th><th>Training Type</th><th>Start Date</th><th>End Date</th><th>Trainer</th></tr>
                </thead>
                <tbody id="completedBody"></tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- ======================== REMOVED TAB ======================== -->
        <div class="tab-content" id="tab-removed">
          <div class="page-header"><h1>Removed Students</h1></div>
          <div class="card" style="padding:0;">
            <div class="table-wrap">
              <table>
                <thead>
                  <tr><th>Name</th><th>CID</th><th>Training Type</th><th>Joined</th><th>Removed</th><th>Notes</th></tr>
                </thead>
                <tbody id="removedBody"></tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- ======================== EXAM REQUESTS TAB ======================== -->
        <div class="tab-content" id="tab-exam-requests">
          <div class="page-header"><h1>Exam Requests</h1></div>

          <!-- Pending requests -->
          <div class="card mb-4" style="padding:0;">
            <div class="table-wrap">
              <table>
                <thead>
                  <tr><th>Name</th><th>CID</th><th>VATSIM Rating</th><th>Exam Requested</th><th>Requested</th><th>Actions</th></tr>
                </thead>
                <tbody id="examRequestsBody"></tbody>
              </table>
            </div>
          </div>

          <!-- Code send history -->
          <div class="page-header"><h1 style="font-size:1.1rem;">Exam Code Send History</h1></div>
          <div id="examCodeHistory"></div>
        </div>

        <!-- ======================== EXAM CENTER TAB (Admin) ======================== -->
        <div class="tab-content" id="tab-exam-center">
          <div class="page-header">
            <div class="flex items-center gap-3">
              <h1>Exam Center</h1>
              <button class="btn btn-primary btn-sm" onclick="openExamBuilder()">+ Create Exam</button>
            </div>
          </div>

          <!-- Exam list by category -->
          <div id="adminExamList"></div>

          <!-- Exam Stats -->
          <div class="section-divider"></div>
          <div class="page-header"><h1 style="font-size:1.1rem;">Exam Statistics</h1></div>
          <div class="card mb-4">
            <label class="form-label">Select Exam</label>
            <select class="form-select" id="statsExamSelect" onchange="loadExamStats()">
              <option value="">Choose an exam...</option>
            </select>
          </div>
          <div id="examStatsArea"></div>
        </div>

        <!-- ======================== ADMIN TAB ======================== -->
        <div class="tab-content" id="tab-admin">
          <div class="page-header"><h1>Admin Dashboard</h1></div>

          <!-- Training Stats -->
          <div class="card mb-4">
            <div class="card-header"><span class="card-title">Training Statistics</span></div>
            <div id="trainingStatsGrid"></div>
          </div>

          <!-- Trainers -->
          <div class="page-header" style="margin-top:24px;">
            <div class="flex items-center gap-3">
              <h1 style="font-size:1.1rem;">Trainers</h1>
              <button class="btn btn-secondary btn-sm" onclick="syncTrainersFromVATUSA()">Sync from VATUSA</button>
              <button class="btn btn-primary btn-sm" onclick="openAddTrainerModal()">+ Add Manually</button>
            </div>
          </div>
          <div class="trainer-grid" id="trainerGrid"></div>
        </div>

      </div><!-- /admin-content -->
    </div><!-- /admin-layout -->
  </div>

  <!-- =================== MODALS =================== -->

  <!-- Reject Modal -->
  <div class="modal-backdrop" id="rejectModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Reject Training Request</span>
        <button class="modal-close" onclick="closeModal('rejectModal')">&times;</button>
      </div>
      <p class="text-muted mb-4">Select a pre-written reason or write a custom message.</p>
      <div class="form-group">
        <label class="form-label">Pre-written Reasons</label>
        <select class="form-select" id="rejectPrefab" onchange="fillPrefabReject()">
          <option value="">Select a reason...</option>
          <option value="[PLACEHOLDER] Coursework not completed. Please complete all required coursework before resubmitting.">Coursework not completed</option>
          <option value="[PLACEHOLDER] Your VATSIM rating does not meet the requirements for this training type at this time.">Rating requirement not met</option>
          <option value="[PLACEHOLDER] Our training queue is currently at capacity. Please resubmit when space opens up.">Queue at capacity</option>
          <option value="custom">Custom message...</option>
        </select>
      </div>
      <div class="form-group" id="customRejectWrap" style="display:none;">
        <label class="form-label">Custom Message</label>
        <textarea class="form-textarea" id="rejectCustomMsg" placeholder="Write your rejection reason..."></textarea>
      </div>
      <div id="rejectMsgPreview" class="note-chip" style="display:none;"></div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-danger" onclick="confirmReject()">Send Rejection</button>
        <button class="btn btn-secondary" onclick="closeModal('rejectModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Note Modal -->
  <div class="modal-backdrop" id="noteModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Add Note</span>
        <button class="modal-close" onclick="closeModal('noteModal')">&times;</button>
      </div>
      <div class="form-group">
        <label class="form-label">Note</label>
        <textarea class="form-textarea" id="noteText" placeholder="Enter your note..."></textarea>
      </div>
      <div id="existingNotes" class="mb-4"></div>
      <div class="flex gap-2">
        <button class="btn btn-primary" onclick="submitNote()">Save Note</button>
        <button class="btn btn-secondary" onclick="closeModal('noteModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Trainer Assignment Modal -->
  <div class="modal-backdrop" id="trainerModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Assign Trainer</span>
        <button class="modal-close" onclick="closeModal('trainerModal')">&times;</button>
      </div>
      <div class="form-group">
        <label class="form-label">Select Trainer</label>
        <select class="form-select" id="trainerSelect">
          <option value="">Choose trainer...</option>
        </select>
      </div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-primary" onclick="confirmTrainerAssign()">Assign</button>
        <button class="btn btn-secondary" onclick="closeModal('trainerModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Add Student Modal -->
  <div class="modal-backdrop" id="addStudentModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Manually Add Student to Queue</span>
        <button class="modal-close" onclick="closeModal('addStudentModal')">&times;</button>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">First Name</label>
          <input type="text" class="form-input" id="addFirstName">
        </div>
        <div class="form-group">
          <label class="form-label">Last Name</label>
          <input type="text" class="form-input" id="addLastName">
        </div>
      </div>
      <div class="form-group">
        <label class="form-label">CID</label>
        <input type="text" class="form-input" id="addCid">
      </div>
      <div class="form-group">
        <label class="form-label">Training Type</label>
        <select class="form-select" id="addTrainingType">
          <option value="">Select...</option>
          <option value="S1 Training (Tier 2)">S1 Training (Tier 2)</option>
          <option value="S1 Training (Unrestricted)">S1 Training (Unrestricted)</option>
          <option value="S2 Training (Tier 2)">S2 Training (Tier 2)</option>
          <option value="S2 Training (Unrestricted)">S2 Training (Unrestricted)</option>
          <option value="S3 Training (Tier 2)">S3 Training (Tier 2)</option>
          <option value="S3 Training (Unrestricted)">S3 Training (Unrestricted)</option>
          <option value="C1 Training">C1 Training</option>
          <option value="Miami Center Oceanic">Miami Center Oceanic</option>
          <option value="Tier 1 MIA GND/DEL">Tier 1 MIA GND/DEL</option>
          <option value="Tier 1 MIA LCL">Tier 1 MIA LCL</option>
          <option value="Tier 1 MIA TRACON">Tier 1 MIA TRACON</option>
          <option value="Visitor/Transfer — Miami Center Domestic">Visitor/Transfer — Miami Center Domestic</option>
          <option value="Visitor/Transfer — Miami CAB">Visitor/Transfer — Miami CAB</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Queue</label>
        <select class="form-select" id="addQueueType">
          <option value="home">Home Queue</option>
          <option value="visitor">Visitor Queue</option>
        </select>
      </div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-primary" onclick="confirmAddStudent()">Add to Queue</button>
        <button class="btn btn-secondary" onclick="closeModal('addStudentModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Remove Student Modal -->
  <div class="modal-backdrop" id="removeModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Remove Student</span>
        <button class="modal-close" onclick="closeModal('removeModal')">&times;</button>
      </div>
      <p class="text-muted mb-4">Add a note for this removal (optional).</p>
      <div class="form-group">
        <textarea class="form-textarea" id="removeNotes" placeholder="Reason for removal..."></textarea>
      </div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-danger" onclick="confirmRemove()">Remove Student</button>
        <button class="btn btn-secondary" onclick="closeModal('removeModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Exam Reject Modal -->
  <div class="modal-backdrop" id="examRejectModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Deny Exam Request</span>
        <button class="modal-close" onclick="closeModal('examRejectModal')">&times;</button>
      </div>
      <div class="form-group">
        <label class="form-label">Pre-written Reasons</label>
        <select class="form-select" id="examRejectPrefab" onchange="fillExamRejectPrefab()">
          <option value="">Select a reason...</option>
          <option value="[PLACEHOLDER] You are not eligible for this exam at this time. Please contact your trainer.">Not eligible at this time</option>
          <option value="[PLACEHOLDER] Please complete the required prerequisites before requesting this exam.">Prerequisites not met</option>
          <option value="custom">Custom message...</option>
        </select>
      </div>
      <div class="form-group" id="examCustomRejectWrap" style="display:none;">
        <textarea class="form-textarea" id="examRejectCustomMsg" placeholder="Custom reason..."></textarea>
      </div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-danger" onclick="confirmExamReject()">Send Denial</button>
        <button class="btn btn-secondary" onclick="closeModal('examRejectModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Exam Builder Modal -->
  <div class="modal-backdrop" id="examBuilderModal">
    <div class="modal modal-xl">
      <div class="modal-header">
        <span class="modal-title" id="examBuilderTitle">Create Exam</span>
        <button class="modal-close" onclick="closeModal('examBuilderModal')">&times;</button>
      </div>
      <div id="examBuilderContent"></div>
    </div>
  </div>

  <!-- Add Trainer Modal -->
  <div class="modal-backdrop" id="addTrainerModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Add Trainer Manually</span>
        <button class="modal-close" onclick="closeModal('addTrainerModal')">&times;</button>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label class="form-label">First Name</label>
          <input class="form-input" id="trainerFirstName">
        </div>
        <div class="form-group">
          <label class="form-label">Last Name</label>
          <input class="form-input" id="trainerLastName">
        </div>
      </div>
      <div class="form-group">
        <label class="form-label">CID</label>
        <input class="form-input" id="trainerCid">
      </div>
      <div class="form-group">
        <label class="form-label">Rating</label>
        <input class="form-input" id="trainerRating" placeholder="e.g. S3">
      </div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-primary" onclick="confirmAddTrainer()">Add Trainer</button>
        <button class="btn btn-secondary" onclick="closeModal('addTrainerModal')">Cancel</button>
      </div>
    </div>
  </div>

  <!-- Manual Grade Modal -->
  <div class="modal-backdrop" id="manualGradeModal">
    <div class="modal">
      <div class="modal-header">
        <span class="modal-title">Add Manual Exam Grade</span>
        <button class="modal-close" onclick="closeModal('manualGradeModal')">&times;</button>
      </div>
      <div class="form-group">
        <label class="form-label">CID</label>
        <input class="form-input" id="manualGradeCid" placeholder="Student CID">
      </div>
      <div class="form-group">
        <label class="form-label">Student Name</label>
        <input class="form-input" id="manualGradeName" placeholder="Full name">
      </div>
      <div class="form-group">
        <label class="form-label">Score (%)</label>
        <input class="form-input" type="number" id="manualGradeScore" min="0" max="100">
      </div>
      <div class="form-group">
        <label class="form-label">Graded By</label>
        <input class="form-input" id="manualGradeBy" placeholder="Your name">
      </div>
      <div class="flex gap-2 mt-4">
        <button class="btn btn-primary" onclick="confirmManualGrade()">Save Grade</button>
        <button class="btn btn-secondary" onclick="closeModal('manualGradeModal')">Cancel</button>
      </div>
    </div>
  </div>

  <script src="js/config.js"></script>
  <script>
    emailjs.init(CONFIG.EMAILJS_PUBLIC_KEY);

    // ===================== STATE =====================
    let state = {
      pending: [], home: [], visitor: [],
      completed: [], removed: [],
      examRequests: [], exams: [], trainers: [],
      examCodeHistory: [],
      currentRejectId: null,
      currentRejectEmail: null,
      currentRejectName: null,
      currentNoteQueueId: null,
      currentTrainerQueueId: null,
      currentRemoveId: null,
      currentExamRejectId: null,
      currentExamRejectEmail: null,
      currentExamRejectName: null,
      currentExamRejectExamTitle: null,
      editingExamId: null,
      manualGradeExamId: null,
    };

    // ===================== TAB SWITCHING =====================
    function switchTab(name) {
      document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.sidebar-btn').forEach(b => b.classList.remove('active'));
      document.getElementById(`tab-${name}`).classList.add('active');
      event.target.classList.add('active');
      loadTab(name);
    }

    function loadTab(name) {
      if (name === 'queues') loadQueues();
      if (name === 'completed') loadCompleted();
      if (name === 'removed') loadRemoved();
      if (name === 'exam-requests') loadExamRequests();
      if (name === 'exam-center') loadAdminExams();
      if (name === 'admin') loadAdmin();
    }

    // ===================== MODAL HELPERS =====================
    function openModal(id) { document.getElementById(id).classList.add('open'); }
    function closeModal(id) { document.getElementById(id).classList.remove('open'); }

    // ===================== QUEUES =====================
    async function loadQueues() {
      try {
        state.pending = await SB.query('training_requests', '?status=eq.pending&order=created_at.asc');
        state.home = await SB.query('training_queue', '?queue_type=eq.home&status=not.in.(completed,removed)&order=sort_order.asc,requested_at.asc');
        state.visitor = await SB.query('training_queue', '?queue_type=eq.visitor&status=not.in.(completed,removed)&order=sort_order.asc,requested_at.asc');
        renderQueueStats();
        renderPending();
        renderQueueTable('homeBody', state.home, 'home');
        renderQueueTable('visitorBody', state.visitor, 'visitor');
      } catch (e) { console.error(e); toast('Failed to load queues', 'error'); }
    }

    function renderQueueStats() {
      const total = state.home.length + state.visitor.length;
      document.getElementById('queueStats').innerHTML = `
        <div class="stat-box">
          <div class="stat-label">Home Queue</div>
          <div class="stat-value">${state.home.length}</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">Visitor Queue</div>
          <div class="stat-value">${state.visitor.length}</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">Total Active</div>
          <div class="stat-value">${total}</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">Pending Approval</div>
          <div class="stat-value">${state.pending.length}</div>
        </div>
      `;
    }

    function renderPending() {
      const section = document.getElementById('pendingSection');
      section.style.display = state.pending.length ? 'block' : 'none';
      document.getElementById('pendingBody').innerHTML = state.pending.map(r => `
        <tr>
          <td>${r.first_name} ${r.last_name}</td>
          <td><a class="cid-link" href="https://stats.vatsim.net/dashboard/${r.cid}" target="_blank">${r.cid}</a></td>
          <td>${r.training_type}</td>
          <td>${formatDate(r.requested_at)}</td>
          <td><button class="btn btn-ghost btn-sm" onclick="loadExamGrades('${r.cid}', this)">View</button></td>
          <td class="action-row">
            <button class="btn btn-success btn-sm" onclick="approveRequest('${r.id}','${r.cid}','${r.first_name}','${r.last_name}','${r.training_type}','${r.email || ''}')">Approve</button>
            <button class="btn btn-danger btn-sm" onclick="openRejectModal('${r.id}','${r.email || ''}','${r.first_name} ${r.last_name}')">Reject</button>
          </td>
        </tr>
      `).join('') || '<tr><td colspan="6" class="text-muted" style="text-align:center;padding:20px;">No pending requests</td></tr>';
    }

    async function loadExamGrades(cid, btn) {
      btn.textContent = 'Loading...';
      const grades = await VATUSA.getExamGrades(cid);
      btn.textContent = grades.length ? grades.map(g => `${g.exam_id}: ${g.score}%`).join(', ') : 'None found';
    }

    async function approveRequest(id, cid, firstName, lastName, trainingType, email) {
      const membership = await VATUSA.isMemberOfFacility(cid, CONFIG.FACILITY);
      if (!membership) {
        toast('This controller is not a home or visiting member of ZMA.', 'error');
        return;
      }
      const queueType = membership.type;
      try {
        const [req] = await SB.query('training_requests', `?id=eq.${id}`);
        await SB.update('training_requests', `id=eq.${id}`, { status: 'approved', approved_at: new Date().toISOString(), queue_type: queueType });
        const maxOrder = queueType === 'home'
          ? Math.max(0, ...state.home.map(e => e.sort_order || 0))
          : Math.max(0, ...state.visitor.map(e => e.sort_order || 0));
        await SB.insert('training_queue', {
          request_id: id, first_name: firstName, last_name: lastName,
          cid: cid, email: email, training_type: trainingType,
          queue_type: queueType, status: 'waiting',
          sort_order: maxOrder + 1,
          requested_at: req.requested_at || new Date().toISOString(),
        });
        toast(`${firstName} ${lastName} approved and added to ${queueType} queue.`);
        loadQueues();
      } catch (e) { console.error(e); toast('Error approving request', 'error'); }
    }

    function openRejectModal(id, email, name) {
      state.currentRejectId = id;
      state.currentRejectEmail = email;
      state.currentRejectName = name;
      document.getElementById('rejectPrefab').value = '';
      document.getElementById('customRejectWrap').style.display = 'none';
      document.getElementById('rejectMsgPreview').style.display = 'none';
      openModal('rejectModal');
    }

    function fillPrefabReject() {
      const val = document.getElementById('rejectPrefab').value;
      if (val === 'custom') {
        document.getElementById('customRejectWrap').style.display = 'block';
        document.getElementById('rejectMsgPreview').style.display = 'none';
      } else if (val) {
        document.getElementById('customRejectWrap').style.display = 'none';
        document.getElementById('rejectMsgPreview').style.display = 'block';
        document.getElementById('rejectMsgPreview').textContent = val;
      } else {
        document.getElementById('customRejectWrap').style.display = 'none';
        document.getElementById('rejectMsgPreview').style.display = 'none';
      }
    }

    async function confirmReject() {
      const prefab = document.getElementById('rejectPrefab').value;
      const custom = document.getElementById('rejectCustomMsg').value.trim();
      const reason = prefab === 'custom' ? custom : prefab;
      if (!reason) { toast('Please select or write a reason.', 'error'); return; }
      try {
        await SB.update('training_requests', `id=eq.${state.currentRejectId}`, {
          status: 'rejected', rejected_at: new Date().toISOString(), rejection_reason: reason
        });
        if (state.currentRejectEmail) {
          await EM.sendRejection(state.currentRejectEmail, state.currentRejectName, reason);
        }
        toast('Request rejected and email sent.');
        closeModal('rejectModal');
        loadQueues();
      } catch (e) { toast('Error rejecting request', 'error'); }
    }

    function renderQueueTable(bodyId, entries, queueType) {
      const tbody = document.getElementById(bodyId);
      if (!entries.length) {
        tbody.innerHTML = '<tr><td colspan="9" class="text-muted" style="text-align:center;padding:20px;">No students in queue</td></tr>';
        return;
      }
      tbody.innerHTML = entries.map((e, i) => `
        <tr draggable="true" data-id="${e.id}" data-queue="${queueType}" data-idx="${i}">
          <td class="drag-handle">&#8597;</td>
          <td>${e.first_name} ${e.last_name}</td>
          <td><a class="cid-link" href="https://stats.vatsim.net/dashboard/${e.cid}" target="_blank">${e.cid}</a></td>
          <td>${e.training_type}</td>
          <td>${formatDate(e.requested_at)}</td>
          <td>
            <span class="badge ${e.status === 'training' ? 'badge-blue' : 'badge-gray'}">
              ${e.status === 'training' ? 'In Training' : 'Waiting'}
            </span>
          </td>
          <td><button class="btn btn-ghost btn-sm" onclick="loadExamGrades('${e.cid}', this)">View</button></td>
          <td>
            ${e.trainer_name ? `<span class="badge badge-blue">${e.trainer_name}</span>` : '—'}
            <button class="btn btn-ghost btn-sm" onclick="openTrainerModal('${e.id}')">Assign</button>
          </td>
          <td>
            <div class="action-row">
              <button class="btn btn-success btn-sm" onclick="markComplete('${e.id}','${e.cid}','${e.first_name}','${e.last_name}','${e.training_type}','${e.trainer_name || ''}','${e.requested_at}','${e.email || ''}')">Complete</button>
              <button class="btn btn-ghost btn-sm" onclick="openNoteModal('${e.id}')">Note</button>
              <button class="btn btn-danger btn-sm" onclick="openRemoveModal('${e.id}')">Remove</button>
              <select class="form-select" style="width:auto;padding:4px 8px;font-size:.75rem;" onchange="moveStudent('${e.id}', this.value, '${queueType}')">
                <option value="">Move to...</option>
                ${queueType === 'home' ? '<option value="visitor">Visitor Queue</option>' : '<option value="home">Home Queue</option>'}
              </select>
            </div>
          </td>
        </tr>
      `).join('');
      initDragDrop(bodyId, queueType);
    }

    // Drag and drop within/between queues
    function initDragDrop(bodyId, queueType) {
      const tbody = document.getElementById(bodyId);
      let dragSrc = null;
      tbody.querySelectorAll('tr').forEach(row => {
        row.addEventListener('dragstart', e => { dragSrc = row; row.classList.add('dragging'); });
        row.addEventListener('dragend', e => { row.classList.remove('dragging'); tbody.querySelectorAll('tr').forEach(r => r.classList.remove('drag-over')); });
        row.addEventListener('dragover', e => { e.preventDefault(); row.classList.add('drag-over'); });
        row.addEventListener('dragleave', e => row.classList.remove('drag-over'));
        row.addEventListener('drop', async e => {
          e.preventDefault();
          row.classList.remove('drag-over');
          if (dragSrc === row) return;
          const arr = Array.from(tbody.querySelectorAll('tr'));
          const srcIdx = arr.indexOf(dragSrc);
          const tgtIdx = arr.indexOf(row);
          tbody.insertBefore(dragSrc, srcIdx < tgtIdx ? row.nextSibling : row);
          // Update sort_order in DB
          const newOrder = Array.from(tbody.querySelectorAll('tr')).map(r => r.dataset.id);
          for (let i = 0; i < newOrder.length; i++) {
            await SB.update('training_queue', `id=eq.${newOrder[i]}`, { sort_order: i });
          }
        });
      });
    }

    async function moveStudent(id, toQueue, fromQueue) {
      if (!toQueue) return;
      await SB.update('training_queue', `id=eq.${id}`, { queue_type: toQueue });
      toast(`Moved to ${toQueue} queue.`);
      loadQueues();
    }

    // Trainer assignment
    async function openTrainerModal(queueId) {
      state.currentTrainerQueueId = queueId;
      const sel = document.getElementById('trainerSelect');
      sel.innerHTML = '<option value="">Choose trainer...</option>';
      const trainers = await SB.query('trainers', '');
      trainers.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = `${t.first_name} ${t.last_name}`;
        sel.appendChild(opt);
      });
      openModal('trainerModal');
    }

    async function confirmTrainerAssign() {
      const sel = document.getElementById('trainerSelect');
      const trainerId = sel.value;
      const trainerName = sel.options[sel.selectedIndex].textContent;
      if (!trainerId) { toast('Select a trainer.', 'error'); return; }
      await SB.update('training_queue', `id=eq.${state.currentTrainerQueueId}`, {
        trainer_id: trainerId, trainer_name: trainerName, status: 'training', training_started_at: new Date().toISOString()
      });
      toast('Trainer assigned.');
      closeModal('trainerModal');
      loadQueues();
    }

    // Notes
    async function openNoteModal(queueId) {
      state.currentNoteQueueId = queueId;
      document.getElementById('noteText').value = '';
      const notes = await SB.query('queue_notes', `?queue_entry_id=eq.${queueId}&order=created_at.desc`);
      document.getElementById('existingNotes').innerHTML = notes.length
        ? notes.map(n => `<div class="note-chip">${formatDate(n.created_at)}: ${n.note}</div>`).join('')
        : '<p class="text-muted">No notes yet.</p>';
      openModal('noteModal');
    }

    async function submitNote() {
      const text = document.getElementById('noteText').value.trim();
      if (!text) { toast('Enter a note.', 'error'); return; }
      await SB.insert('queue_notes', { queue_entry_id: state.currentNoteQueueId, note: text });
      toast('Note saved.');
      closeModal('noteModal');
    }

    // Complete
    async function markComplete(id, cid, firstName, lastName, trainingType, trainerName, requestedAt, email) {
      if (!confirm(`Mark ${firstName} ${lastName} as training complete?`)) return;
      await SB.update('training_queue', `id=eq.${id}`, { status: 'completed', completed_at: new Date().toISOString() });
      await SB.insert('completed_training', {
        queue_entry_id: id, first_name: firstName, last_name: lastName,
        cid, training_type: trainingType, trainer_name: trainerName,
        training_started_at: requestedAt, completed_at: new Date().toISOString()
      });
      if (email) await EM.sendCompletion(email, `${firstName} ${lastName}`, trainingType);
      toast('Marked complete.');
      loadQueues();
    }

    // Remove
    function openRemoveModal(queueId) {
      state.currentRemoveId = queueId;
      document.getElementById('removeNotes').value = '';
      openModal('removeModal');
    }

    async function confirmRemove() {
      const notes = document.getElementById('removeNotes').value.trim();
      const [entry] = await SB.query('training_queue', `?id=eq.${state.currentRemoveId}`);
      await SB.update('training_queue', `id=eq.${state.currentRemoveId}`, { status: 'removed', removed_at: new Date().toISOString() });
      await SB.insert('removed_students', {
        queue_entry_id: state.currentRemoveId,
        first_name: entry.first_name, last_name: entry.last_name,
        cid: entry.cid, training_type: entry.training_type,
        queue_type: entry.queue_type, joined_at: entry.requested_at,
        removed_at: new Date().toISOString(), removal_notes: notes
      });
      toast('Student removed.');
      closeModal('removeModal');
      loadQueues();
    }

    // Add Student manually
    function openAddStudentModal() { openModal('addStudentModal'); }

    async function confirmAddStudent() {
      const firstName = document.getElementById('addFirstName').value.trim();
      const lastName = document.getElementById('addLastName').value.trim();
      const cid = document.getElementById('addCid').value.trim();
      const trainingType = document.getElementById('addTrainingType').value;
      const queueType = document.getElementById('addQueueType').value;
      if (!firstName || !lastName || !cid || !trainingType) { toast('Fill all fields.', 'error'); return; }
      await SB.insert('training_queue', {
        first_name: firstName, last_name: lastName, cid,
        training_type: trainingType, queue_type: queueType,
        status: 'waiting', requested_at: new Date().toISOString(), sort_order: 999
      });
      toast('Student added to queue.');
      closeModal('addStudentModal');
      loadQueues();
    }

    // ===================== COMPLETED =====================
    async function loadCompleted() {
      const data = await SB.query('completed_training', '?order=completed_at.desc');
      document.getElementById('completedBody').innerHTML = data.map(r => `
        <tr>
          <td>${r.first_name} ${r.last_name}</td>
          <td><a class="cid-link" href="https://stats.vatsim.net/dashboard/${r.cid}" target="_blank">${r.cid}</a></td>
          <td>${r.training_type}</td>
          <td>${formatDate(r.training_started_at)}</td>
          <td>${formatDate(r.completed_at)}</td>
          <td>${r.trainer_name || '—'}</td>
        </tr>
      `).join('') || '<tr><td colspan="6" class="text-muted" style="text-align:center;padding:20px;">No completed training records</td></tr>';
    }

    // ===================== REMOVED =====================
    async function loadRemoved() {
      const data = await SB.query('removed_students', '?order=removed_at.desc');
      document.getElementById('removedBody').innerHTML = data.map(r => `
        <tr>
          <td>${r.first_name} ${r.last_name}</td>
          <td><a class="cid-link" href="https://stats.vatsim.net/dashboard/${r.cid}" target="_blank">${r.cid}</a></td>
          <td>${r.training_type}</td>
          <td>${formatDate(r.joined_at)}</td>
          <td>${formatDate(r.removed_at)}</td>
          <td>
            <span class="text-muted">${r.removal_notes || '—'}</span>
            <button class="btn btn-ghost btn-sm" onclick="editRemovalNote('${r.id}','${(r.removal_notes||'').replace(/'/g,"\\'")}')">Edit</button>
          </td>
        </tr>
      `).join('') || '<tr><td colspan="6" class="text-muted" style="text-align:center;padding:20px;">No removed students</td></tr>';
    }

    async function editRemovalNote(id, current) {
      const note = prompt('Edit removal note:', current);
      if (note === null) return;
      await SB.update('removed_students', `id=eq.${id}`, { removal_notes: note });
      toast('Note updated.');
      loadRemoved();
    }

    // ===================== EXAM REQUESTS =====================
    async function loadExamRequests() {
      state.examRequests = await SB.query('exam_requests', '?status=eq.pending&order=requested_at.asc');
      const history = await SB.query('exam_code_history', '?order=sent_at.desc');
      renderExamRequests();
      renderCodeHistory(history);
      updateExamNotifDot();
    }

    function updateExamNotifDot() {
      const dot = document.getElementById('examNotifDot');
      if (state.examRequests.length > 0) {
        dot.style.display = 'flex';
        dot.textContent = state.examRequests.length;
      } else {
        dot.style.display = 'none';
      }
    }

    function renderExamRequests() {
      document.getElementById('examRequestsBody').innerHTML = state.examRequests.map(r => `
        <tr>
          <td>${r.student_name || '—'}</td>
          <td><a class="cid-link" href="https://stats.vatsim.net/dashboard/${r.cid}" target="_blank">${r.cid}</a></td>
          <td>${r.vatsim_rating || '—'}</td>
          <td>${r.exam_title || '—'}</td>
          <td>${formatDate(r.requested_at)}</td>
          <td>
            <div class="action-row">
              <button class="btn btn-success btn-sm" onclick="approveExamRequest('${r.id}','${r.cid}','${r.student_name}','${r.exam_id}','${r.exam_title}')">Approve</button>
              <button class="btn btn-danger btn-sm" onclick="openExamRejectModal('${r.id}','${r.cid}','${r.student_name}','${r.exam_title}')">Deny</button>
            </div>
          </td>
        </tr>
      `).join('') || '<tr><td colspan="6" class="text-muted" style="text-align:center;padding:20px;">No pending exam requests</td></tr>';
    }

    async function approveExamRequest(reqId, cid, studentName, examId, examTitle) {
      const [exam] = await SB.query('exams', `?id=eq.${examId}`);
      if (!exam) { toast('Exam not found.', 'error'); return; }
      const password = exam.password || '—';
      await SB.update('exam_requests', `id=eq.${reqId}`, {
        status: 'approved', resolved_at: new Date().toISOString(), password_sent: password
      });
      await SB.insert('exam_code_history', {
        cid, student_name: studentName, exam_id: examId, exam_title: examTitle,
        password_sent: password, sent_at: new Date().toISOString()
      });
      // Get email
      const ctrl = await VATUSA.getController(cid);
      if (ctrl) {
        try {
          const emailRes = await fetch(`${CONFIG.VATUSA_API}/user/${cid}/email?apikey=${CONFIG.VATUSA_API_KEY}`);
          const emailData = await emailRes.json();
          if (emailData.email) await EM.sendExamApproved(emailData.email, studentName, examTitle, password);
        } catch {}
      }
      toast('Exam approved. Password sent.');
      loadExamRequests();
    }

    function openExamRejectModal(reqId, cid, name, examTitle) {
      state.currentExamRejectId = reqId;
      state.currentExamRejectName = name;
      state.currentExamRejectExamTitle = examTitle;
      openModal('examRejectModal');
    }

    function fillExamRejectPrefab() {
      const val = document.getElementById('examRejectPrefab').value;
      document.getElementById('examCustomRejectWrap').style.display = val === 'custom' ? 'block' : 'none';
    }

    async function confirmExamReject() {
      const prefab = document.getElementById('examRejectPrefab').value;
      const custom = document.getElementById('examRejectCustomMsg').value.trim();
      const reason = prefab === 'custom' ? custom : prefab;
      if (!reason) { toast('Select a reason.', 'error'); return; }
      await SB.update('exam_requests', `id=eq.${state.currentExamRejectId}`, {
        status: 'rejected', resolved_at: new Date().toISOString(), rejection_reason: reason
      });
      toast('Exam request denied.');
      closeModal('examRejectModal');
      loadExamRequests();
    }

    function renderCodeHistory(history) {
      if (!history.length) {
        document.getElementById('examCodeHistory').innerHTML = '<p class="text-muted">No exam codes sent yet.</p>';
        return;
      }
      // Group by CID
      const grouped = {};
      history.forEach(h => {
        if (!grouped[h.cid]) grouped[h.cid] = { name: h.student_name, items: [] };
        grouped[h.cid].items.push(h);
      });
      document.getElementById('examCodeHistory').innerHTML = Object.entries(grouped).map(([cid, g]) => `
        <div>
          <div class="collapsible-header" onclick="toggleCollapse(this)">
            <span><strong>${g.name || cid}</strong> — CID ${cid}</span>
            <span class="text-muted">${g.items.length} sent</span>
          </div>
          <div class="collapsible-body">
            <div class="table-wrap">
              <table>
                <thead><tr><th>Exam</th><th>Password</th><th>Sent</th></tr></thead>
                <tbody>
                  ${g.items.map(i => `<tr><td>${i.exam_title}</td><td class="text-mono">${i.password_sent}</td><td>${formatDate(i.sent_at)}</td></tr>`).join('')}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      `).join('');
    }

    function toggleCollapse(header) {
      const body = header.nextElementSibling;
      body.classList.toggle('open');
    }

    // ===================== EXAM CENTER (Admin) =====================
    async function loadAdminExams() {
      state.exams = await SB.query('exams', '?order=created_at.desc');
      renderAdminExamList();
      populateStatsExamSelect();
    }

    function renderAdminExamList() {
      const categories = [
        { id: 'tier2', label: 'Tier 2 Exams' },
        { id: 'tier1', label: 'Tier 1 Exams' },
        { id: 'training', label: 'Training Exams' },
      ];
      document.getElementById('adminExamList').innerHTML = categories.map(cat => {
        const exams = state.exams.filter(e => e.category === cat.id);
        return `
          <div class="exam-cat-header">${cat.label}</div>
          ${exams.length ? exams.map(ex => `
            <div class="card mb-4" style="padding:16px 20px;">
              <div class="flex items-center gap-3">
                <div style="flex:1;">
                  <strong>${ex.title}</strong>
                  <span class="badge ${ex.is_published ? 'badge-green' : 'badge-gray'}" style="margin-left:8px;">${ex.is_published ? 'Published' : 'Draft'}</span>
                  <div class="text-muted" style="margin-top:2px;font-size:.8rem;">${ex.description || 'No description'} &bull; ${ex.time_limit ? ex.time_limit + ' min' : 'No time limit'} &bull; Pass: ${ex.passing_grade}% &bull; Attempts: ${ex.allowed_attempts}</div>
                </div>
                <div class="action-row">
                  <button class="btn btn-secondary btn-sm" onclick="openExamBuilder('${ex.id}')">Edit</button>
                  <button class="btn btn-ghost btn-sm" onclick="openManualGradeModal('${ex.id}')">+ Grade</button>
                  <button class="btn btn-danger btn-sm" onclick="deleteExam('${ex.id}')">Delete</button>
                </div>
              </div>
            </div>
          `).join('') : '<p class="text-muted" style="padding:12px;">No exams in this category.</p>'}
        `;
      }).join('');
    }

    function populateStatsExamSelect() {
      const sel = document.getElementById('statsExamSelect');
      sel.innerHTML = '<option value="">Choose an exam...</option>';
      state.exams.forEach(e => {
        const opt = document.createElement('option');
        opt.value = e.id;
        opt.textContent = e.title;
        sel.appendChild(opt);
      });
    }

    async function loadExamStats() {
      const examId = document.getElementById('statsExamSelect').value;
      if (!examId) { document.getElementById('examStatsArea').innerHTML = ''; return; }
      const attempts = await SB.query('exam_attempts', `?exam_id=eq.${examId}&order=submitted_at.desc`);
      const questions = await SB.query('exam_questions', `?exam_id=eq.${examId}`);
      if (!attempts.length) {
        document.getElementById('examStatsArea').innerHTML = '<div class="card"><p class="text-muted">No attempts yet for this exam.</p></div>';
        return;
      }
      const avgScore = Math.round(attempts.filter(a => a.score !== null).reduce((s, a) => s + (a.score||0), 0) / attempts.length);
      const passCount = attempts.filter(a => a.passed).length;
      // Most missed questions
      const missCount = {};
      attempts.forEach(a => {
        if (!a.answers) return;
        questions.forEach(q => {
          const ans = a.answers[q.id];
          if (!isAnswerCorrect(q, ans, a.question_overrides)) {
            missCount[q.id] = (missCount[q.id] || 0) + 1;
          }
        });
      });
      const sortedMissed = Object.entries(missCount).sort((a,b) => b[1]-a[1]).slice(0,5);
      document.getElementById('examStatsArea').innerHTML = `
        <div class="stats-grid">
          <div class="stat-box"><div class="stat-label">Total Attempts</div><div class="stat-value">${attempts.length}</div></div>
          <div class="stat-box"><div class="stat-label">Pass Rate</div><div class="stat-value">${Math.round(passCount/attempts.length*100)}%</div></div>
          <div class="stat-box"><div class="stat-label">Avg Score</div><div class="stat-value">${avgScore}%</div></div>
        </div>
        <div class="card mb-4">
          <div class="card-header"><span class="card-title">Most Missed Questions</span></div>
          ${sortedMissed.map(([qid, count]) => {
            const q = questions.find(x => x.id === qid);
            return `<div style="padding:8px 0;border-bottom:1px solid var(--gray-100);font-size:.875rem;"><strong>${count}x missed:</strong> ${q ? q.question_text.slice(0,100) : qid}</div>`;
          }).join('') || '<p class="text-muted">No missed question data.</p>'}
        </div>
        <div class="card">
          <div class="card-header"><span class="card-title">All Attempts</span></div>
          <div class="table-wrap">
            <table>
              <thead><tr><th>Student</th><th>CID</th><th>Score</th><th>Passed</th><th>Date</th></tr></thead>
              <tbody>
                ${attempts.map(a => `
                  <tr>
                    <td>${a.student_name || '—'}</td>
                    <td class="text-mono">${a.cid}</td>
                    <td>${a.score !== null ? a.score + '%' : '—'}</td>
                    <td><span class="badge ${a.passed ? 'badge-green' : 'badge-red'}">${a.passed ? 'Pass' : 'Fail'}</span></td>
                    <td>${formatDate(a.submitted_at)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    }

    function isAnswerCorrect(question, answer, overrides) {
      if (overrides && overrides[question.id] !== undefined) return overrides[question.id] > 0;
      if (!answer) return false;
      const correct = question.correct_answers;
      if (question.question_type === 'true_false') return answer === correct;
      if (question.question_type === 'short_response') return answer.toLowerCase().trim() === (question.correct_text||'').toLowerCase().trim();
      if (question.question_type === 'multiple_choice_single') return answer === correct[0];
      if (question.question_type === 'multiple_choice_multi') {
        if (!Array.isArray(answer) || !Array.isArray(correct)) return false;
        return answer.sort().join(',') === correct.sort().join(',');
      }
      return false;
    }

    // Exam Builder
    async function openExamBuilder(examId = null) {
      state.editingExamId = examId;
      document.getElementById('examBuilderTitle').textContent = examId ? 'Edit Exam' : 'Create Exam';
      let exam = null, questions = [];
      if (examId) {
        [exam] = await SB.query('exams', `?id=eq.${examId}`);
        questions = await SB.query('exam_questions', `?exam_id=eq.${examId}&order=sort_order.asc`);
      }
      renderExamBuilder(exam, questions);
      openModal('examBuilderModal');
    }

    function renderExamBuilder(exam, questions) {
      document.getElementById('examBuilderContent').innerHTML = `
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Exam Title</label>
            <input class="form-input" id="eb_title" value="${exam?.title || ''}" placeholder="e.g. S1 Ground Control Exam">
          </div>
          <div class="form-group">
            <label class="form-label">Category</label>
            <select class="form-select" id="eb_category">
              <option value="tier2" ${exam?.category==='tier2'?'selected':''}>Tier 2 Exams</option>
              <option value="tier1" ${exam?.category==='tier1'?'selected':''}>Tier 1 Exams</option>
              <option value="training" ${exam?.category==='training'?'selected':''}>Training Exams</option>
            </select>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">Description</label>
          <textarea class="form-textarea" id="eb_description" style="min-height:60px;">${exam?.description || ''}</textarea>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Time Limit (minutes, blank = unlimited)</label>
            <input class="form-input" id="eb_time_limit" type="number" value="${exam?.time_limit || ''}" placeholder="Leave blank for unlimited">
          </div>
          <div class="form-group">
            <label class="form-label">Passing Grade (%)</label>
            <input class="form-input" id="eb_passing_grade" type="number" value="${exam?.passing_grade || 70}" min="0" max="100">
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Allowed Attempts</label>
            <input class="form-input" id="eb_allowed_attempts" type="number" value="${exam?.allowed_attempts || 1}" min="1">
          </div>
          <div class="form-group">
            <label class="form-label">Exam Password</label>
            <input class="form-input" id="eb_password" value="${exam?.password || ''}" placeholder="Set password">
          </div>
        </div>
        <div class="form-group">
          <label class="checkbox-wrap">
            <input type="checkbox" id="eb_allow_request" ${exam?.allow_password_request?'checked':''}>
            <span class="checkbox-label">Allow students to request this exam password from the Exam Center</span>
          </label>
        </div>
        <hr class="section-divider">
        <div class="flex items-center gap-3 mb-4">
          <span class="card-title">Questions</span>
          <button class="btn btn-secondary btn-sm" onclick="addQuestion()">+ Add Question</button>
        </div>
        <div id="questionBank">${questions.map((q,i) => renderQuestionCard(q, i)).join('')}</div>
        <hr class="section-divider">
        <div class="flex gap-2">
          <button class="btn btn-primary" onclick="saveExam(true)">Publish Exam</button>
          <button class="btn btn-secondary" onclick="saveExam(false)">Save as Draft</button>
          <button class="btn btn-ghost" onclick="closeModal('examBuilderModal')">Cancel</button>
        </div>
      `;
    }

    let qCounter = 0;
    function addQuestion() {
      const bank = document.getElementById('questionBank');
      const idx = bank.querySelectorAll('.question-card').length;
      const div = document.createElement('div');
      div.innerHTML = renderQuestionCard(null, idx);
      bank.appendChild(div.firstElementChild);
    }

    function renderQuestionCard(q, idx) {
      const id = q?.id || `new_${++qCounter}`;
      const type = q?.question_type || 'multiple_choice_single';
      const text = q?.question_text || '';
      const points = q?.points || 1;
      return `
        <div class="question-card" data-qid="${id}">
          <div class="question-card-header">
            <span class="question-num">Q${idx+1}</span>
            <select class="form-select" style="width:auto;" onchange="changeQuestionType(this)">
              <option value="multiple_choice_single" ${type==='multiple_choice_single'?'selected':''}>Multiple Choice (Single)</option>
              <option value="multiple_choice_multi" ${type==='multiple_choice_multi'?'selected':''}>Multiple Choice (Multi)</option>
              <option value="true_false" ${type==='true_false'?'selected':''}>True / False</option>
              <option value="short_response" ${type==='short_response'?'selected':''}>Short Response</option>
              <option value="matching" ${type==='matching'?'selected':''}>Matching</option>
            </select>
            <input type="number" class="form-input" style="width:80px;" value="${points}" placeholder="Pts" title="Points">
            <button class="btn btn-danger btn-sm" onclick="this.closest('.question-card').remove()">Remove</button>
          </div>
          <div class="form-group">
            <label class="form-label">Question Text</label>
            <textarea class="form-textarea" style="min-height:60px;" placeholder="Enter question...">${text}</textarea>
          </div>
          <div class="form-group">
            <label class="form-label">Image URL (optional)</label>
            <input class="form-input" placeholder="https://..." value="${q?.image_url||''}">
          </div>
          ${renderQuestionOptions(q, type)}
        </div>
      `;
    }

    function changeQuestionType(sel) {
      const card = sel.closest('.question-card');
      const optArea = card.querySelector('.q-options-area');
      optArea.innerHTML = renderQuestionOptionsHTML(null, sel.value);
    }

    function renderQuestionOptions(q, type) {
      return `<div class="q-options-area">${renderQuestionOptionsHTML(q, type)}</div>`;
    }

    function renderQuestionOptionsHTML(q, type) {
      const opts = q?.options || [];
      const correct = q?.correct_answers || [];
      if (type === 'true_false') {
        const ans = q?.correct_answers || 'true';
        return `
          <div class="form-group">
            <label class="form-label">Correct Answer</label>
            <select class="form-select" data-field="correct_tf">
              <option value="true" ${ans==='true'?'selected':''}>True</option>
              <option value="false" ${ans==='false'?'selected':''}>False</option>
            </select>
          </div>`;
      }
      if (type === 'short_response') {
        return `<div class="form-group"><label class="form-label">Correct Answer</label><input class="form-input" data-field="correct_text" value="${q?.correct_text||''}" placeholder="Exact answer (case-insensitive)"></div>`;
      }
      if (type === 'multiple_choice_single' || type === 'multiple_choice_multi') {
        const isMulti = type === 'multiple_choice_multi';
        const optHtml = (opts.length ? opts : [{id:'a',text:''},{id:'b',text:''},{id:'c',text:''},{id:'d',text:''}]).map(o => `
          <div class="flex items-center gap-2 mb-2">
            <input type="${isMulti?'checkbox':'radio'}" name="correct_${q?.id||'new'}" value="${o.id}" ${(isMulti ? correct.includes(o.id) : correct[0]===o.id)?'checked':''}>
            <input class="form-input" placeholder="Option text" value="${o.text}" data-opt-id="${o.id}">
            <button class="btn btn-ghost btn-sm" onclick="this.parentElement.remove()">&#x2715;</button>
          </div>
        `).join('');
        return `
          <div class="form-group">
            <label class="form-label">Options <span class="text-muted">(check correct answer${isMulti?'s':''})</span></label>
            <div class="q-options-list">${optHtml}</div>
            <button class="btn btn-ghost btn-sm mt-4" onclick="addOption(this)">+ Add Option</button>
          </div>`;
      }
      if (type === 'matching') {
        const pairs = (opts.length ? opts : [{left:'',right:''}]);
        return `
          <div class="form-group">
            <label class="form-label">Matching Pairs (Left &#8594; Right)</label>
            <div class="q-matching-list">
              ${pairs.map(p => `
                <div class="matching-row">
                  <input class="form-input" placeholder="Left side" value="${p.left||''}">
                  <span>&#8594;</span>
                  <input class="form-input" placeholder="Right side" value="${p.right||''}">
                  <button class="btn btn-ghost btn-sm" onclick="this.parentElement.remove()">&#x2715;</button>
                </div>`).join('')}
            </div>
            <button class="btn btn-ghost btn-sm mt-4" onclick="addMatchingPair(this)">+ Add Pair</button>
          </div>`;
      }
      return '';
    }

    function addOption(btn) {
      const list = btn.previousElementSibling;
      const id = 'opt_' + Date.now();
      const div = document.createElement('div');
      div.className = 'flex items-center gap-2 mb-2';
      div.innerHTML = `<input type="radio"><input class="form-input" placeholder="Option text" data-opt-id="${id}"><button class="btn btn-ghost btn-sm" onclick="this.parentElement.remove()">&#x2715;</button>`;
      list.appendChild(div);
    }

    function addMatchingPair(btn) {
      const list = btn.previousElementSibling;
      const div = document.createElement('div');
      div.className = 'matching-row';
      div.innerHTML = `<input class="form-input" placeholder="Left side"><span>&#8594;</span><input class="form-input" placeholder="Right side"><button class="btn btn-ghost btn-sm" onclick="this.parentElement.remove()">&#x2715;</button>`;
      list.appendChild(div);
    }

    function collectExamData() {
      const questions = [];
      document.querySelectorAll('#questionBank .question-card').forEach((card, idx) => {
        const typeSel = card.querySelector('select');
        const type = typeSel.value;
        const questionText = card.querySelectorAll('textarea')[0].value.trim();
        const imageUrl = card.querySelectorAll('input[type="text"], input:not([type="number"]):not([type="checkbox"]):not([type="radio"])')[1]?.value.trim() || '';
        const pointsInput = card.querySelector('input[type="number"]');
        const points = parseInt(pointsInput?.value || '1');
        let options = null, correctAnswers = null, correctText = null;
        if (type === 'true_false') {
          const tfSel = card.querySelector('[data-field="correct_tf"]');
          correctAnswers = tfSel?.value || 'true';
        } else if (type === 'short_response') {
          correctText = card.querySelector('[data-field="correct_text"]')?.value.trim() || '';
        } else if (type === 'multiple_choice_single' || type === 'multiple_choice_multi') {
          options = [];
          correctAnswers = [];
          card.querySelectorAll('.q-options-list .flex').forEach(row => {
            const inp = row.querySelector('input[data-opt-id]');
            const checkbox = row.querySelector('input[type="radio"], input[type="checkbox"]');
            if (inp) {
              const optId = inp.dataset.optId;
              options.push({ id: optId, text: inp.value });
              if (checkbox?.checked) correctAnswers.push(optId);
            }
          });
          if (type === 'multiple_choice_single') correctAnswers = correctAnswers.slice(0,1);
        } else if (type === 'matching') {
          options = [];
          correctAnswers = [];
          card.querySelectorAll('.q-matching-list .matching-row').forEach((row, i) => {
            const inputs = row.querySelectorAll('input');
            const leftId = `left_${i}`, rightId = `right_${i}`;
            options.push({ id: leftId, text: inputs[0]?.value, side: 'left' });
            options.push({ id: rightId, text: inputs[1]?.value, side: 'right' });
            correctAnswers.push({ left_id: leftId, right_id: rightId });
          });
        }
        questions.push({ question_type: type, question_text: questionText, image_url: imageUrl, options, correct_answers: correctAnswers, correct_text: correctText, points, sort_order: idx });
      });
      return questions;
    }

    async function saveExam(publish) {
      const title = document.getElementById('eb_title').value.trim();
      if (!title) { toast('Enter an exam title.', 'error'); return; }
      const examData = {
        title,
        category: document.getElementById('eb_category').value,
        description: document.getElementById('eb_description').value.trim(),
        time_limit: parseInt(document.getElementById('eb_time_limit').value) || null,
        passing_grade: parseInt(document.getElementById('eb_passing_grade').value) || 70,
        allowed_attempts: parseInt(document.getElementById('eb_allowed_attempts').value) || 1,
        password: document.getElementById('eb_password').value.trim(),
        allow_password_request: document.getElementById('eb_allow_request').checked,
        is_published: publish,
        updated_at: new Date().toISOString(),
      };
      let examId = state.editingExamId;
      try {
        if (examId) {
          await SB.update('exams', `id=eq.${examId}`, examData);
          await SB.delete('exam_questions', `exam_id=eq.${examId}`);
        } else {
          const [newExam] = await SB.insert('exams', examData);
          examId = newExam.id;
        }
        const questions = collectExamData();
        if (questions.length) {
          await SB.insert('exam_questions', questions.map(q => ({ ...q, exam_id: examId })));
        }
        toast(publish ? 'Exam published.' : 'Exam saved as draft.');
        closeModal('examBuilderModal');
        loadAdminExams();
      } catch (e) { console.error(e); toast('Error saving exam.', 'error'); }
    }

    async function deleteExam(id) {
      if (!confirm('Delete this exam and all its data?')) return;
      await SB.delete('exams', `id=eq.${id}`);
      toast('Exam deleted.');
      loadAdminExams();
    }

    // Manual Grade
    function openManualGradeModal(examId) {
      state.manualGradeExamId = examId;
      openModal('manualGradeModal');
    }

    async function confirmManualGrade() {
      const cid = document.getElementById('manualGradeCid').value.trim();
      const name = document.getElementById('manualGradeName').value.trim();
      const score = parseInt(document.getElementById('manualGradeScore').value);
      const by = document.getElementById('manualGradeBy').value.trim();
      const [exam] = await SB.query('exams', `?id=eq.${state.manualGradeExamId}`);
      await SB.insert('exam_attempts', {
        exam_id: state.manualGradeExamId,
        cid, student_name: name, score,
        passed: score >= (exam?.passing_grade || 70),
        is_manual: true, graded_by: by,
        submitted_at: new Date().toISOString()
      });
      toast('Grade saved.');
      closeModal('manualGradeModal');
    }

    // ===================== ADMIN TAB =====================
    async function loadAdmin() {
      loadTrainingStats();
      loadTrainers();
    }

    async function loadTrainingStats() {
      const completed = await SB.query('completed_training', '?order=completed_at.desc');
      const queue = await SB.query('training_queue', '');
      // Avg wait time by training type
      const byType = {};
      completed.forEach(c => {
        if (!byType[c.training_type]) byType[c.training_type] = [];
        if (c.training_started_at && c.completed_at) {
          const days = Math.round((new Date(c.completed_at) - new Date(c.training_started_at)) / 86400000);
          byType[c.training_type].push(days);
        }
      });
      // By trainer
      const byTrainer = {};
      completed.forEach(c => {
        if (!c.trainer_name) return;
        if (!byTrainer[c.trainer_name]) byTrainer[c.trainer_name] = 0;
        byTrainer[c.trainer_name]++;
      });
      document.getElementById('trainingStatsGrid').innerHTML = `
        <div class="table-wrap mb-4">
          <table>
            <thead><tr><th>Training Type</th><th>Avg Days to Complete</th><th>Total Completed</th></tr></thead>
            <tbody>
              ${Object.entries(byType).map(([type, days]) => `
                <tr><td>${type}</td><td>${Math.round(days.reduce((a,b)=>a+b,0)/days.length)} days</td><td>${days.length}</td></tr>
              `).join('') || '<tr><td colspan="3" class="text-muted">No data yet.</td></tr>'}
            </tbody>
          </table>
        </div>
        <div class="card-title mb-4">Completions by Trainer</div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Trainer</th><th>Completions</th></tr></thead>
            <tbody>
              ${Object.entries(byTrainer).sort((a,b)=>b[1]-a[1]).map(([name,count]) => `
                <tr><td>${name}</td><td>${count}</td></tr>
              `).join('') || '<tr><td colspan="2" class="text-muted">No data yet.</td></tr>'}
            </tbody>
          </table>
        </div>
      `;
    }

    async function loadTrainers() {
      state.trainers = await SB.query('trainers', '?order=last_name.asc');
      // Get active queue counts
      const queueEntries = await SB.query('training_queue', '?status=not.in.(completed,removed)');
      const countByTrainer = {};
      queueEntries.forEach(e => {
        if (e.trainer_name) countByTrainer[e.trainer_name] = (countByTrainer[e.trainer_name]||0)+1;
      });
      renderTrainerGrid(countByTrainer);
    }

    function renderTrainerGrid(countByTrainer) {
      const grid = document.getElementById('trainerGrid');
      if (!state.trainers.length) {
        grid.innerHTML = '<p class="text-muted">No trainers yet. Sync from VATUSA or add manually.</p>';
        return;
      }
      grid.innerHTML = state.trainers.map(t => {
        const certs = (t.certifications || []);
        // Filter: if T1S1 selected, hide S1, etc.
        const displayCerts = certs.filter(c => {
          if (c === 'S1' && certs.includes('T1S1')) return false;
          if (c === 'S2' && certs.includes('T1S2')) return false;
          if (c === 'S3' && certs.includes('T1S3')) return false;
          return true;
        });
        const studentCount = countByTrainer[`${t.first_name} ${t.last_name}`] || 0;
        return `
          <div class="trainer-card">
            <div class="trainer-name">${t.first_name} ${t.last_name}</div>
            <span class="trainer-rating">${t.rating || '—'}</span>
            <div class="trainer-certs">${displayCerts.map(c => `<span class="trainer-cert">${c}</span>`).join('') || '<span class="text-muted" style="font-size:.75rem;">No certs</span>'}</div>
            <div class="trainer-students">${studentCount}</div>
            <div class="trainer-students-label">Students in queue</div>
            <div class="cert-toggle mt-4" id="certs_${t.id}">
              ${CERTIFICATIONS.map(c => `<span class="cert-chip ${certs.includes(c)?'selected':''}" onclick="toggleCert('${t.id}','${c}',this)">${c}</span>`).join('')}
            </div>
            <button class="btn btn-danger btn-sm mt-4" onclick="removeTrainer('${t.id}')">Remove</button>
          </div>
        `;
      }).join('');
    }

    async function toggleCert(trainerId, cert, el) {
      const trainer = state.trainers.find(t => t.id === trainerId);
      if (!trainer) return;
      const certs = [...(trainer.certifications || [])];
      const idx = certs.indexOf(cert);
      if (idx >= 0) certs.splice(idx, 1); else certs.push(cert);
      await SB.update('trainers', `id=eq.${trainerId}`, { certifications: certs });
      trainer.certifications = certs;
      el.classList.toggle('selected');
    }

    async function removeTrainer(id) {
      if (!confirm('Remove this trainer?')) return;
      await SB.delete('trainers', `id=eq.${id}`);
      toast('Trainer removed.');
      loadTrainers();
    }

    async function syncTrainersFromVATUSA() {
      toast('Syncing from VATUSA...');
      const staff = await VATUSA.getFacilityStaff(CONFIG.FACILITY);
      if (!staff || !staff.length) { toast('No staff found or API error.', 'error'); return; }
      for (const s of staff) {
        const existing = await SB.query('trainers', `?cid=eq.${s.cid}`);
        if (!existing.length) {
          await SB.insert('trainers', {
            cid: String(s.cid),
            first_name: s.fname || '',
            last_name: s.lname || '',
            rating: VATUSA.ratingName(s.rating),
            certifications: [],
            is_manual: false,
          });
        }
      }
      toast('Trainers synced from VATUSA.');
      loadTrainers();
    }

    function openAddTrainerModal() { openModal('addTrainerModal'); }

    async function confirmAddTrainer() {
      const first = document.getElementById('trainerFirstName').value.trim();
      const last = document.getElementById('trainerLastName').value.trim();
      const cid = document.getElementById('trainerCid').value.trim();
      const rating = document.getElementById('trainerRating').value.trim();
      if (!first || !last || !cid) { toast('Fill all required fields.', 'error'); return; }
      await SB.insert('trainers', { cid, first_name: first, last_name: last, rating, certifications: [], is_manual: true });
      toast('Trainer added.');
      closeModal('addTrainerModal');
      loadTrainers();
    }

    // ===================== POLLING =====================
    setInterval(() => {
      // Check active tab and refresh
      const active = document.querySelector('.tab-content.active');
      if (!active) return;
      const id = active.id.replace('tab-','');
      loadTab(id);
    }, 10000);

    // Check exam notif dot
    async function checkExamNotif() {
      const pending = await SB.query('exam_requests', '?status=eq.pending');
      const dot = document.getElementById('examNotifDot');
      if (pending.length) {
        dot.style.display = 'flex';
        dot.textContent = pending.length;
      } else {
        dot.style.display = 'none';
      }
    }
    setInterval(checkExamNotif, 15000);

    // ===================== INIT =====================
    loadQueues();
    checkExamNotif();
  </script>
</body>
</html>
