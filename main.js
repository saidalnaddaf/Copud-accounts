/* ============================================
   CUPOD REGISTRATION — MAIN SCRIPT
   ============================================ */

(function() {
  'use strict';

  // ---------- Element refs ----------
  const form = document.getElementById('registrationForm');
  const submitBtn = document.getElementById('submitBtn');
  const resetBtn = document.getElementById('resetBtn');
  const enrollDateInput = document.getElementById('enrollDate');
  const previewId = document.getElementById('previewId');
  const pvTotal = document.getElementById('pvTotal');
  const pvPaid = document.getElementById('pvPaid');
  const pvBalance = document.getElementById('pvBalance');
  const pvAccess = document.getElementById('pvAccess');
  const toast = document.getElementById('toast');
  const toastIcon = document.getElementById('toastIcon');
  const toastTitle = document.getElementById('toastTitle');
  const toastMsg = document.getElementById('toastMsg');
  const successModal = document.getElementById('successModal');
  const modalDetails = document.getElementById('modalDetails');
  const modalCloseBtn = document.getElementById('modalCloseBtn');
  const addAnotherBtn = document.getElementById('addAnotherBtn');
  const viewSheetLink = document.getElementById('viewSheetLink');

  // ---------- Init ----------
  document.addEventListener('DOMContentLoaded', () => {
    setDefaultDate();
    setSheetLink();
    setupRevealAnimations();
    setupLivePreview();
    setupValidation();
    setupSubmit();
    setupModal();
  });

  // ---------- Set today's date as default ----------
  function setDefaultDate() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    enrollDateInput.value = `${yyyy}-${mm}-${dd}`;
  }

  // ---------- Master sheet link ----------
  function setSheetLink() {
    if (CUPOD_CONFIG.SHEET_VIEW_URL && !CUPOD_CONFIG.SHEET_VIEW_URL.includes('PASTE')) {
      viewSheetLink.href = CUPOD_CONFIG.SHEET_VIEW_URL;
    } else {
      viewSheetLink.style.display = 'none';
    }
  }

  // ---------- Reveal animation ----------
  function setupRevealAnimations() {
    const reveals = document.querySelectorAll('.reveal');
    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('visible');
        }
      });
    }, { threshold: 0.05 });
    reveals.forEach(el => observer.observe(el));
  }

  // ---------- Live preview ----------
  function setupLivePreview() {
    const inputs = ['coursePrice', 'firstPayment', 'enrollDate'];
    inputs.forEach(id => {
      document.getElementById(id).addEventListener('input', updatePreview);
    });
    updatePreview();
  }

  function updatePreview() {
    const price = parseFloat(document.getElementById('coursePrice').value) || 0;
    const paid = parseFloat(document.getElementById('firstPayment').value) || 0;
    const balance = Math.max(0, price - paid);
    const dateStr = enrollDateInput.value;

    pvTotal.textContent = formatMoney(price);
    pvPaid.textContent = formatMoney(paid);
    pvBalance.textContent = formatMoney(balance);

    if (dateStr) {
      const d = new Date(dateStr + 'T00:00:00');
      d.setDate(d.getDate() + (CUPOD_CONFIG.ACCESS_DAYS || 90));
      pvAccess.textContent = d.toLocaleDateString('en-US', {
        year: 'numeric', month: 'short', day: 'numeric'
      });
    } else {
      pvAccess.textContent = '—';
    }
  }

  function formatMoney(n) {
    return '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  // ---------- Validation ----------
  function setupValidation() {
    form.querySelectorAll('input').forEach(input => {
      input.addEventListener('blur', () => validateField(input));
      input.addEventListener('input', () => {
        if (input.classList.contains('invalid')) {
          validateField(input);
        }
      });
    });
  }

  function validateField(input) {
    const id = input.id;
    const errEl = document.getElementById('err-' + id);
    let valid = true;
    let message = '';

    const val = (input.value || '').trim();

    if (input.required && !val) {
      valid = false;
      message = 'This field is required';
    } else {
      switch (id) {
        case 'studentName':
          if (val.length < 2) { valid = false; message = 'Name must be at least 2 characters'; }
          else if (val.length > 80) { valid = false; message = 'Name is too long'; }
          break;
        case 'phone':
          // Allow digits, spaces, +, -, parentheses; need at least 6 digits
          const digitCount = (val.match(/\d/g) || []).length;
          if (digitCount < 6) { valid = false; message = 'Enter a valid phone number'; }
          else if (val.length > 30) { valid = false; message = 'Phone is too long'; }
          break;
        case 'email':
          if (val && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val)) {
            valid = false; message = 'Enter a valid email or leave blank';
          }
          break;
        case 'coursePrice':
          if (isNaN(parseFloat(val))) { valid = false; message = 'Must be a number'; }
          else if (parseFloat(val) < 0) { valid = false; message = 'Cannot be negative'; }
          else if (parseFloat(val) > 10000) { valid = false; message = 'Price seems too high'; }
          break;
        case 'firstPayment':
          if (val && isNaN(parseFloat(val))) { valid = false; message = 'Must be a number'; }
          else if (val && parseFloat(val) < 0) { valid = false; message = 'Cannot be negative'; }
          else {
            const price = parseFloat(document.getElementById('coursePrice').value) || 0;
            if (parseFloat(val) > price) { valid = false; message = 'Cannot exceed course price'; }
          }
          break;
        case 'enrollDate':
          if (!val) { valid = false; message = 'Date is required'; }
          break;
      }
    }

    if (valid) {
      input.classList.remove('invalid');
      if (errEl) {
        errEl.textContent = '';
        errEl.classList.remove('visible');
      }
    } else {
      input.classList.add('invalid');
      if (errEl) {
        errEl.textContent = message;
        errEl.classList.add('visible');
      }
    }

    return valid;
  }

  function validateAll() {
    let allValid = true;
    let firstInvalid = null;
    form.querySelectorAll('input').forEach(input => {
      const ok = validateField(input);
      if (!ok && !firstInvalid) firstInvalid = input;
      if (!ok) allValid = false;
    });
    if (firstInvalid) firstInvalid.focus();
    return allValid;
  }

  // ---------- Submit ----------
  function setupSubmit() {
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (!validateAll()) {
        showToast('error', 'Check the form', 'Some fields need fixing.');
        return;
      }

      // Check config
      if (!CUPOD_CONFIG.APPS_SCRIPT_URL || CUPOD_CONFIG.APPS_SCRIPT_URL.includes('PASTE')) {
        showToast('error', 'Not configured', 'Please paste your Apps Script URL into js/config.js.');
        return;
      }

      submitBtn.classList.add('loading');
      submitBtn.disabled = true;

      const payload = {
        studentName: document.getElementById('studentName').value.trim(),
        email: document.getElementById('email').value.trim(),
        phone: document.getElementById('phone').value.trim(),
        coursePrice: parseFloat(document.getElementById('coursePrice').value) || 0,
        firstPayment: parseFloat(document.getElementById('firstPayment').value) || 0,
        enrollDate: document.getElementById('enrollDate').value,
        accessDays: CUPOD_CONFIG.ACCESS_DAYS || 90,
      };

      try {
        // Apps Script doPost expects URL-encoded form data when sent from a browser
        // (avoids CORS preflight issues). We send JSON in the body of a form param.
        const formData = new FormData();
        formData.append('payload', JSON.stringify(payload));

        const response = await fetch(CUPOD_CONFIG.APPS_SCRIPT_URL, {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          throw new Error('Server returned ' + response.status);
        }

        const result = await response.json();

        if (result.status === 'success') {
          showSuccess(result, payload);
        } else {
          throw new Error(result.message || 'Unknown error');
        }

      } catch (err) {
        console.error(err);
        showToast('error', 'Could not save', err.message || 'Check your connection and try again.');
      } finally {
        submitBtn.classList.remove('loading');
        submitBtn.disabled = false;
      }
    });

    resetBtn.addEventListener('click', () => {
      // Wait a tick so the form actually clears, then reset defaults
      setTimeout(() => {
        setDefaultDate();
        document.getElementById('coursePrice').value = '350';
        document.getElementById('firstPayment').value = '0';
        form.querySelectorAll('input').forEach(i => {
          i.classList.remove('invalid');
        });
        form.querySelectorAll('.error-msg').forEach(e => {
          e.textContent = '';
          e.classList.remove('visible');
        });
        updatePreview();
      }, 0);
    });
  }

  // ---------- Success modal ----------
  function showSuccess(result, payload) {
    const balance = (payload.coursePrice - payload.firstPayment).toFixed(2);
    const accessDate = new Date(payload.enrollDate + 'T00:00:00');
    accessDate.setDate(accessDate.getDate() + payload.accessDays);

    modalDetails.innerHTML = `
      <div class="id-row"><strong>${result.studentId}</strong></div>
      <div><strong>${escapeHtml(payload.studentName)}</strong></div>
      <div>Enrolled: ${formatDate(payload.enrollDate)}</div>
      <div>Course: ${formatMoney(payload.coursePrice)} · Paid: ${formatMoney(payload.firstPayment)} · Balance: ${formatMoney(parseFloat(balance))}</div>
      <div>Access ends: ${accessDate.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'})}</div>
    `;

    successModal.classList.add('show');
    showToast('success', 'Saved!', `${result.studentId} added to the master sheet.`);
  }

  function setupModal() {
    modalCloseBtn.addEventListener('click', () => {
      successModal.classList.remove('show');
    });
    addAnotherBtn.addEventListener('click', () => {
      successModal.classList.remove('show');
      form.reset();
      setTimeout(() => {
        setDefaultDate();
        document.getElementById('coursePrice').value = '350';
        document.getElementById('firstPayment').value = '0';
        updatePreview();
        document.getElementById('studentName').focus();
      }, 0);
    });
    successModal.querySelector('.modal-backdrop').addEventListener('click', () => {
      successModal.classList.remove('show');
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && successModal.classList.contains('show')) {
        successModal.classList.remove('show');
      }
    });
  }

  // ---------- Toast ----------
  let toastTimer;
  function showToast(type, title, msg) {
    clearTimeout(toastTimer);
    toast.classList.remove('show', 'error');
    if (type === 'error') toast.classList.add('error');
    toastIcon.textContent = type === 'error' ? '!' : '✓';
    toastTitle.textContent = title;
    toastMsg.textContent = msg;
    requestAnimationFrame(() => toast.classList.add('show'));
    toastTimer = setTimeout(() => toast.classList.remove('show'), 4500);
  }

  // ---------- Helpers ----------
  function formatDate(iso) {
    const d = new Date(iso + 'T00:00:00');
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

})();
