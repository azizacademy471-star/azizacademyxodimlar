const form = document.getElementById('employeeForm');
const alertBox = document.getElementById('formAlert');
const submitBtn = document.getElementById('submitBtn');
const filledCount = document.getElementById('filledCount');

const monthSelect = document.getElementById('workMonth');
const yearSelect = document.getElementById('workYear');

const TEACHER_POSITION = 'Ustoz';

const monthNames = [
  'Yanvar', 'Fevral', 'Mart', 'Aprel', 'May', 'Iyun',
  'Iyul', 'Avgust', 'Sentabr', 'Oktabr', 'Noyabr', 'Dekabr'
];

function getApiBaseUrl() {
  const cfg = window.APP_CONFIG || {};
  const raw = String(cfg.apiBaseUrl || '').trim();
  if (raw) {
    return raw.replace(/\/$/, '');
  }
  return window.location.origin.replace(/\/$/, '');
}


function getAdminUrl() {
  const cfg = window.APP_CONFIG || {};
  const raw = String(cfg.adminUrl || '').trim();
  if (raw) {
    return raw;
  }
  return `${getApiBaseUrl()}/admin`;
}

function populateWorkDateOptions() {
  const currentYear = new Date().getFullYear();

  for (let year = currentYear; year >= 1990; year -= 1) {
    const option = document.createElement('option');
    option.value = String(year);
    option.textContent = String(year);
    yearSelect.appendChild(option);
  }

  monthNames.forEach((month, index) => {
    const option = document.createElement('option');
    option.value = String(index + 1).padStart(2, '0');
    option.textContent = month;
    monthSelect.appendChild(option);
  });
}

function setAlert(type, message) {
  alertBox.className = `form-alert ${type}`;
  alertBox.textContent = message;
}

function hideAlert() {
  alertBox.className = 'form-alert hidden';
  alertBox.textContent = '';
}

function getStateValue(name) {
  const input = form.querySelector(`input[name="${name}"]`);
  return input ? input.value : 'false';
}

function setStateValue(name, value) {
  const input = form.querySelector(`input[name="${name}"]`);
  if (input) input.value = String(value);
}

function toggleButtonState(button, isActive) {
  button.classList.toggle('active', isActive);
}

function setupGenderToggle() {
  const hiddenInput = document.getElementById('gender');
  document.querySelectorAll('[data-toggle-group="gender"] .toggle-btn').forEach((button) => {
    button.addEventListener('click', () => {
      document.querySelectorAll('[data-toggle-group="gender"] .toggle-btn').forEach((btn) => btn.classList.remove('active'));
      button.classList.add('active');
      hiddenInput.value = button.dataset.value;
      updateFilledCount();
    });
  });
}

function setElementHiddenState(element, shouldHide) {
  element.classList.toggle('hidden', shouldHide);
  element.querySelectorAll('input, select, textarea, button').forEach((control) => {
    if (control.type === 'hidden') return;
    control.disabled = shouldHide;
  });
}

function clearValuesInside(element) {
  element.querySelectorAll('input, select, textarea').forEach((control) => {
    if (control.type === 'hidden') {
      if (control.name !== 'gender') control.value = 'false';
      return;
    }
    if (control.type === 'date' || control.type === 'text' || control.type === 'tel' || control.tagName === 'TEXTAREA') {
      control.value = '';
    }
    if (control.tagName === 'SELECT') {
      control.selectedIndex = 0;
    }
  });
  element.querySelectorAll('.toggle-btn, .mini-toggle').forEach((button) => button.classList.remove('active'));
}

function updateConditionalVisibility() {
  document.querySelectorAll('[data-conditional]').forEach((element) => {
    const stateName = element.dataset.conditional;
    const hideWhen = element.dataset.hideWhen;
    const shouldHide = getStateValue(stateName) === hideWhen;
    if (shouldHide) clearValuesInside(element);
    setElementHiddenState(element, shouldHide);
  });

  document.querySelectorAll('[data-conditional-combo]').forEach((element) => {
    const rules = element.dataset.conditionalCombo.split(',').map((item) => item.trim());
    const matches = rules.every((rule) => {
      const [name, expectedValue] = rule.split(':');
      return getStateValue(name) === expectedValue;
    });
    const shouldHide = !matches;
    if (shouldHide) clearValuesInside(element);
    setElementHiddenState(element, shouldHide);
  });

  document.querySelectorAll('[data-conditional-any]').forEach((element) => {
    const rules = element.dataset.conditionalAny.split(',').map((item) => item.trim());
    const shouldHide = rules.some((rule) => {
      const [name, expectedValue] = rule.split(':');
      return getStateValue(name) === expectedValue;
    });
    if (shouldHide) clearValuesInside(element);
    setElementHiddenState(element, shouldHide);
  });

  syncPrimaryRelationInputs();
}

function syncPrimaryRelationInputs() {
  const fatherName = form.querySelector('[name="father_name"]');
  const motherName = form.querySelector('[name="mother_name"]');
  const spouseName = form.querySelector('[name="spouse_name"]');
  const fatherInLawName = form.querySelector('[name="father_in_law_name"]');
  const motherInLawName = form.querySelector('[name="mother_in_law_name"]');
  const teacherSubject = document.getElementById('teacherSubject');

  const disableControl = (control, shouldDisable, placeholder) => {
    if (!control) return;
    control.disabled = shouldDisable;
    control.classList.toggle('disabled-input', shouldDisable);
    if (!control.dataset.defaultPlaceholder) {
      control.dataset.defaultPlaceholder = control.placeholder || '';
    }
    control.placeholder = shouldDisable ? placeholder : control.dataset.defaultPlaceholder;
    if (shouldDisable) control.value = '';
  };

  disableControl(fatherName, getStateValue('father_deceased') === 'true', 'Vafot etgan deb belgilandi');
  disableControl(motherName, getStateValue('mother_deceased') === 'true', 'Vafot etgan deb belgilandi');
  disableControl(fatherInLawName, getStateValue('father_in_law_deceased') === 'true', 'Vafot etgan deb belgilandi');
  disableControl(motherInLawName, getStateValue('mother_in_law_deceased') === 'true', 'Vafot etgan deb belgilandi');

  const spouseDisabled = getStateValue('spouse_deceased') === 'true'
    || getStateValue('spouse_unmarried') === 'true'
    || getStateValue('spouse_divorced') === 'true';
  const spousePlaceholder = getStateValue('spouse_unmarried') === 'true'
    ? 'Turmush qurmaganman deb belgilandi'
    : getStateValue('spouse_divorced') === 'true'
      ? 'Ajrashganman deb belgilandi'
      : 'Vafot etgan deb belgilandi';
  disableControl(spouseName, spouseDisabled, spousePlaceholder);

  if (teacherSubject && teacherSubject.closest('.field')?.classList.contains('hidden')) {
    teacherSubject.value = '';
  }
}

function setupStateToggles() {
  document.querySelectorAll('.mini-toggle[data-state-target]').forEach((button) => {
    button.addEventListener('click', () => {
      const target = button.dataset.stateTarget;
      const currentValue = getStateValue(target) === 'true';
      const nextValue = !currentValue;
      setStateValue(target, nextValue);
      toggleButtonState(button, nextValue);

      if (['spouse_deceased', 'spouse_unmarried', 'spouse_divorced'].includes(target) && nextValue) {
        ['spouse_deceased', 'spouse_unmarried', 'spouse_divorced']
          .filter((name) => name !== target)
          .forEach((name) => {
            setStateValue(name, false);
            const relatedButton = document.querySelector(`[data-state-target="${name}"]`);
            if (relatedButton) relatedButton.classList.remove('active');
          });
      }

      updateConditionalVisibility();
      updateFilledCount();
    });
  });
}

function formatPhone(value) {
  const digits = value.replace(/\D/g, '').slice(0, 12);
  if (!digits) return '';

  let normalized = digits;
  if (!normalized.startsWith('998')) {
    normalized = `998${normalized.slice(-9)}`.slice(0, 12);
  }

  const parts = [];
  if (normalized.length > 0) parts.push(`+${normalized.slice(0, 3)}`);
  if (normalized.length > 3) parts.push(normalized.slice(3, 5));
  if (normalized.length > 5) parts.push(normalized.slice(5, 8));
  if (normalized.length > 8) parts.push(normalized.slice(8, 10));
  if (normalized.length > 10) parts.push(normalized.slice(10, 12));
  return parts.join(' ');
}

function attachInputMasks() {
  form.querySelectorAll('input[type="tel"]').forEach((input) => {
    input.addEventListener('input', () => {
      input.value = formatPhone(input.value);
      updateFilledCount();
    });
  });

  form.querySelectorAll('input[type="date"]').forEach((input) => {
    input.addEventListener('change', updateFilledCount);
  });
}

function setupPositionToggle() {
  const positionSelect = document.getElementById('position');
  const teacherSubjectField = document.getElementById('teacherSubjectField');
  const teacherSubjectSelect = document.getElementById('teacherSubject');

  if (!positionSelect || !teacherSubjectField || !teacherSubjectSelect) return;

  const syncTeacherField = () => {
    const isTeacher = positionSelect.value === TEACHER_POSITION;
    teacherSubjectField.classList.toggle('hidden', !isTeacher);
    teacherSubjectSelect.disabled = !isTeacher;
    teacherSubjectSelect.required = isTeacher;

    if (!isTeacher) {
      teacherSubjectSelect.value = '';
    }

    updateFilledCount();
  };

  positionSelect.addEventListener('change', syncTeacherField);
  syncTeacherField();
}

function updateFilledCount() {
  const controls = [...form.querySelectorAll('input, select')]
    .filter((element) => element.type !== 'hidden' && !element.disabled);
  const filled = controls.filter((element) => String(element.value || '').trim() !== '').length;
  filledCount.textContent = String(filled);
}

function validateForm() {
  hideAlert();

  const requiredSelectors = [
    '[name="full_name"]',
    '[name="phone"]',
    '[name="branch"]',
    '[name="position"]',
    '#gender',
    '[name="birth_date"]',
    '[name="work_start_year"]',
    '[name="work_start_month"]',
  ];

  if (form.querySelector('[name="position"]')?.value === TEACHER_POSITION) {
    requiredSelectors.push('[name="teacher_subject"]');
  }

  for (const selector of requiredSelectors) {
    const element = form.querySelector(selector);
    if (!element || element.disabled) continue;
    if (!String(element.value || '').trim()) {
      setAlert('error', 'Iltimos, barcha majburiy maydonlarni to‘ldiring.');
      const fieldBox = element.closest('.field') || element;
      fieldBox.scrollIntoView({ behavior: 'smooth', block: 'center' });
      if (typeof element.focus === 'function') element.focus();
      return false;
    }
  }

  return true;
}

function formDataToObject(formData) {
  const object = {};
  for (const [key, value] of formData.entries()) {
    object[key] = value;
  }
  return object;
}

function resetFormState() {
  form.reset();
  document.querySelectorAll('.toggle-btn, .mini-toggle').forEach((button) => button.classList.remove('active'));
  document.querySelectorAll('input[type="hidden"]').forEach((input) => {
    if (input.name !== 'gender') input.value = 'false';
  });
  document.getElementById('gender').value = '';

  const teacherSubjectField = document.getElementById('teacherSubjectField');
  const teacherSubjectSelect = document.getElementById('teacherSubject');
  if (teacherSubjectField && teacherSubjectSelect) {
    teacherSubjectField.classList.add('hidden');
    teacherSubjectSelect.disabled = true;
    teacherSubjectSelect.required = false;
    teacherSubjectSelect.value = '';
  }

  updateConditionalVisibility();
  updateFilledCount();
}

async function submitForm(event) {
  event.preventDefault();

  if (!validateForm()) return;

  submitBtn.disabled = true;
  submitBtn.textContent = 'Yuborilmoqda...';

  try {
    const payload = formDataToObject(new FormData(form));
    const response = await fetch(`${getApiBaseUrl()}/api/submit`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    let result = {};
    try {
      result = await response.json();
    } catch (jsonError) {
      result = {};
    }

    if (!response.ok || !result.success) {
      throw new Error(result.error || result.message || 'Serverga yuborishda xatolik yuz berdi.');
    }

    const totalText = Number.isFinite(Number(result.total_rows)) ? ` Jami yozuvlar: ${result.total_rows}.` : '';
    setAlert('success', `Ma’lumotlar muvaffaqiyatli yuborildi.${totalText} Excel fayl yangilandi.`);
    resetFormState();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  } catch (error) {
    setAlert('error', error.message || 'Yuborishda xatolik yuz berdi.');
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = 'Yuborish';
  }
}

function setupAutoCount() {
  form.addEventListener('input', updateFilledCount);
  form.addEventListener('change', updateFilledCount);
}

populateWorkDateOptions();
setupGenderToggle();
setupStateToggles();
setupPositionToggle();
attachInputMasks();
setupAutoCount();
updateConditionalVisibility();
updateFilledCount();
form.addEventListener('submit', submitForm);

const adminLink = document.getElementById('adminPanelLink');
if (adminLink) {
  adminLink.href = getAdminUrl();
}
