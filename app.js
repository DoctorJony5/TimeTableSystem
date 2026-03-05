// Global constants shared by form logic, rendering, and export features.
// ToDo: Make a separate config file & perhaps work on separating things a bit.
const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const PRESET_TYPES = ["Class", "Lecture", "Lab", "Workshop", "Seminar", "Tutorial"];
const LEGACY_STORAGE_KEY = "timetable-studio:v1";
const STORAGE_KEY = "timetable-studio:v2";
const THEME_KEY = "timetable-studio:theme";
const PANEL_STATE_KEY = "timetable-studio:panelstate";
const DEFAULT_SUBJECT_COLOR = "#159f8b"; // Ugly colour that I can easily find in the UI to change or use as a default.
const SEMESTER_WEEKS = 16;

const DEFAULT_DISPLAY_SETTINGS = Object.freeze({
  showWeekends: true,
  showTeacher: true,
  showLocation: true,
  compactCards: false,
  dayStartHour: 7,
  dayEndHour: 22,
  hourHeight: 56
});

const DEFAULT_EXPORT_SETTINGS = Object.freeze({
  theme: "current",
  pngScale: 3,
  includeLegend: true,
  includeClassList: true,
  pdfOrientation: "landscape",
  pdfPageSize: "a4",
  filePrefix: "timetable"
});

const DEFAULT_PANEL_STATE = Object.freeze({
  subject: true,
  classSession: true,
  settings: true
});

// This is probably a bad idea, because it is just a massive block of things
// But this should be fine, hopefully ?
// Maybe at some point make this smarter / better? Helper functions or whatever. Need to ask about it a bit more.
const els = {
  classForm: document.getElementById("classForm"),
  subjectForm: document.getElementById("subjectForm"),
  formTitle: document.getElementById("formTitle"),
  subjectFormTitle: document.getElementById("subjectFormTitle"),
  saveClassBtn: document.getElementById("saveClassBtn"),
  saveSubjectBtn: document.getElementById("saveSubjectBtn"),
  cancelEditBtn: document.getElementById("cancelEditBtn"),
  cancelSubjectEditBtn: document.getElementById("cancelSubjectEditBtn"),
  resetFormBtn: document.getElementById("resetFormBtn"),
  resetSubjectBtn: document.getElementById("resetSubjectBtn"),
  formError: document.getElementById("formError"),
  subjectFormError: document.getElementById("subjectFormError"),
  settingsStatus: document.getElementById("settingsStatus"),
  subjectList: document.getElementById("subjectList"),
  classList: document.getElementById("classList"),
  summaryText: document.getElementById("summaryText"),
  timetableGrid: document.getElementById("timetableGrid"),
  timetableCapture: document.getElementById("timetableCapture"),
  exportStatus: document.getElementById("exportStatus"),
  themeToggle: document.getElementById("themeToggle"),
  clearAllBtn: document.getElementById("clearAllBtn"),
  typeSelect: document.getElementById("typeSelect"),
  customTypeWrap: document.getElementById("customTypeWrap"),
  customTypeInput: document.getElementById("customTypeInput"),
  importJsonInput: document.getElementById("importJsonInput"),
  exportPngBtn: document.getElementById("exportPngBtn"),
  exportPdfBtn: document.getElementById("exportPdfBtn"),
  exportIcsBtn: document.getElementById("exportIcsBtn"),
  exportJsonBtn: document.getElementById("exportJsonBtn"),
  panelToggles: document.querySelectorAll("[data-panel-toggle]"),
  fields: {
    title: document.getElementById("titleInput"),
    subject: document.getElementById("subjectSelect"),
    type: document.getElementById("typeSelect"),
    customType: document.getElementById("customTypeInput"),
    day: document.getElementById("daySelect"),
    color: document.getElementById("colorInput"),
    inheritSubjectColor: document.getElementById("inheritSubjectColorInput"),
    startTime: document.getElementById("startTimeInput"),
    endTime: document.getElementById("endTimeInput"),
    location: document.getElementById("locationInput"),
    teacher: document.getElementById("teacherInput"),
    notes: document.getElementById("notesInput")
  },
  subjectFields: {
    name: document.getElementById("subjectNameInput"),
    code: document.getElementById("subjectCodeInput"),
    teacher: document.getElementById("subjectTeacherInput"),
    room: document.getElementById("subjectRoomInput"),
    color: document.getElementById("subjectColorInput")
  },
  displayInputs: {
    showWeekends: document.getElementById("showWeekendsInput"),
    showTeacher: document.getElementById("showTeacherInput"),
    showLocation: document.getElementById("showLocationInput"),
    compactCards: document.getElementById("compactCardsInput"),
    dayStartHour: document.getElementById("dayStartHourSelect"),
    dayEndHour: document.getElementById("dayEndHourSelect"),
    hourHeight: document.getElementById("hourHeightSelect")
  },
  exportInputs: {
    theme: document.getElementById("exportThemeSelect"),
    pngScale: document.getElementById("pngScaleSelect"),
    includeLegend: document.getElementById("includeLegendInput"),
    includeClassList: document.getElementById("includeClassListInput"),
    pdfOrientation: document.getElementById("pdfOrientationSelect"),
    pdfPageSize: document.getElementById("pdfPageSizeSelect"),
    filePrefix: document.getElementById("filePrefixInput")
  }
};

const state = {
  subjects: [],
  classes: [],
  editSubjectId: null,
  editClassId: null,
  theme: "light",
  display: { ...DEFAULT_DISPLAY_SETTINGS },
  exportPrefs: { ...DEFAULT_EXPORT_SETTINGS },
  panelState: { ...DEFAULT_PANEL_STATE }
};

init();

// ToDo: This init function is doing a lot, but it needs to wait until the DOM is loaded and it needs to be in one place. 
// Maybe split the logic into separate functions or something, but this should be fine for now.
function init() {
  hydrateTheme();
  populateHourSelects();
  hydrateAppData();
  hydratePanelState();
  wireEvents();
  applyPanelStates();
  applyStateToControls();
  applyDisplayCssVars();
  updateCustomTypeVisibility();
  updateColorInheritanceUI();
  renderAll();
  syncClassDefaultsFromSubject();
}

function wireEvents() {
  els.panelToggles.forEach((toggle) => {
    toggle.addEventListener("click", () => {
      togglePanel(toggle.dataset.panelToggle);
    });
  });

  els.subjectForm.addEventListener("submit", handleSubjectSubmit);
  els.cancelSubjectEditBtn.addEventListener("click", resetSubjectFormState);
  els.resetSubjectBtn.addEventListener("click", () => {
    clearSubjectFormError();
    if (state.editSubjectId) {
      resetSubjectFormState();
      return;
    }
    els.subjectForm.reset();
    els.subjectFields.color.value = DEFAULT_SUBJECT_COLOR;
  });

  els.subjectList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }

    const subjectId = button.dataset.id;
    if (button.dataset.action === "edit-subject") {
      startEditSubject(subjectId);
      return;
    }

    if (button.dataset.action === "delete-subject") {
      deleteSubject(subjectId);
    }
  });

  els.classForm.addEventListener("submit", handleClassSubmit);
  els.cancelEditBtn.addEventListener("click", resetClassFormState);
  els.resetFormBtn.addEventListener("click", () => {
    clearClassFormError();
    if (state.editClassId) {
      resetClassFormState();
      return;
    }

    els.classForm.reset();
    els.fields.startTime.value = "09:00";
    els.fields.endTime.value = "10:00";
    els.fields.inheritSubjectColor.checked = true;
    updateCustomTypeVisibility();
    updateColorInheritanceUI();
    syncClassDefaultsFromSubject();
  });

  els.classList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }

    const classId = button.dataset.id;
    if (button.dataset.action === "edit-class") {
      startEditClass(classId);
      return;
    }

    if (button.dataset.action === "delete-class") {
      deleteClass(classId);
    }
  });

  els.typeSelect.addEventListener("change", updateCustomTypeVisibility);
  els.fields.subject.addEventListener("change", syncClassDefaultsFromSubject);
  els.fields.inheritSubjectColor.addEventListener("change", () => {
    updateColorInheritanceUI();
    syncClassDefaultsFromSubject();
  });

  [
    els.displayInputs.showWeekends,
    els.displayInputs.showTeacher,
    els.displayInputs.showLocation,
    els.displayInputs.compactCards,
    els.displayInputs.dayStartHour,
    els.displayInputs.dayEndHour,
    els.displayInputs.hourHeight
  ].forEach((input) => {
    input.addEventListener("change", handleDisplaySettingsChange);
  });

  [
    els.exportInputs.theme,
    els.exportInputs.pngScale,
    els.exportInputs.includeLegend,
    els.exportInputs.includeClassList,
    els.exportInputs.pdfOrientation,
    els.exportInputs.pdfPageSize
  ].forEach((input) => {
    input.addEventListener("change", handleExportSettingsChange);
  });

  els.exportInputs.filePrefix.addEventListener("input", handleExportSettingsChange);

  els.themeToggle.addEventListener("click", toggleTheme);
  els.clearAllBtn.addEventListener("click", clearAllClasses);
  els.exportPngBtn.addEventListener("click", exportAsPng);
  els.exportPdfBtn.addEventListener("click", exportAsPdf);
  els.exportIcsBtn.addEventListener("click", exportAsIcs);
  els.exportJsonBtn.addEventListener("click", exportAsJson);
  els.importJsonInput.addEventListener("change", importFromJson);
}

function populateHourSelects() {
  els.displayInputs.dayStartHour.innerHTML = "";
  els.displayInputs.dayEndHour.innerHTML = "";

  for (let hour = 6; hour <= 23; hour += 1) {
    const startOption = document.createElement("option");
    startOption.value = String(hour);
    startOption.textContent = `${pad2(hour)}:00`;
    els.displayInputs.dayStartHour.append(startOption);

    const endOption = document.createElement("option");
    endOption.value = String(hour);
    endOption.textContent = `${pad2(hour)}:00`;
    els.displayInputs.dayEndHour.append(endOption);
  }
}

function applyStateToControls() {
  els.displayInputs.showWeekends.checked = state.display.showWeekends;
  els.displayInputs.showTeacher.checked = state.display.showTeacher;
  els.displayInputs.showLocation.checked = state.display.showLocation;
  els.displayInputs.compactCards.checked = state.display.compactCards;
  els.displayInputs.dayStartHour.value = String(state.display.dayStartHour);
  els.displayInputs.dayEndHour.value = String(state.display.dayEndHour);
  els.displayInputs.hourHeight.value = String(state.display.hourHeight);

  els.exportInputs.theme.value = state.exportPrefs.theme;
  els.exportInputs.pngScale.value = String(state.exportPrefs.pngScale);
  els.exportInputs.includeLegend.checked = state.exportPrefs.includeLegend;
  els.exportInputs.includeClassList.checked = state.exportPrefs.includeClassList;
  els.exportInputs.pdfOrientation.value = state.exportPrefs.pdfOrientation;
  els.exportInputs.pdfPageSize.value = state.exportPrefs.pdfPageSize;
  els.exportInputs.filePrefix.value = state.exportPrefs.filePrefix;
}

function handleSubjectSubmit(event) {
  event.preventDefault();
  clearSubjectFormError();

  const draft = getSubjectFromForm();
  const validationError = validateSubjectDraft(draft);
  if (validationError) {
    setSubjectFormError(validationError);
    return;
  }

  const normalized = normalizeSubjectForStorage(draft, state.editSubjectId);
  const wasEditing = Boolean(state.editSubjectId);

  if (wasEditing) {
    const index = state.subjects.findIndex((item) => item.id === state.editSubjectId);
    if (index >= 0) {
      state.subjects[index] = normalized;
    }
  } else {
    state.subjects.push(normalized);
  }

  persistAppState();
  renderAll();
  resetSubjectFormState();
  setStatus(wasEditing ? "Subject updated." : "Subject created.", "success");
}

function getSubjectFromForm() {
  return {
    id: state.editSubjectId,
    name: els.subjectFields.name.value.trim(),
    code: els.subjectFields.code.value.trim(),
    teacher: els.subjectFields.teacher.value.trim(),
    defaultRoom: els.subjectFields.room.value.trim(),
    color: normalizeHexColor(els.subjectFields.color.value)
  };
}

function validateSubjectDraft(item) {
  if (!item.name) {
    return "Subject name is required.";
  }

  const duplicate = state.subjects.find(
    (subject) =>
      subject.id !== state.editSubjectId &&
      subject.name.trim().toLowerCase() === item.name.trim().toLowerCase()
  );

  if (duplicate) {
    return "A subject with that name already exists.";
  }

  return null;
}

function normalizeSubjectForStorage(item, forcedId) {
  return {
    id: forcedId || createId(),
    name: item.name.slice(0, 80),
    code: item.code.slice(0, 20),
    teacher: item.teacher.slice(0, 60),
    defaultRoom: item.defaultRoom.slice(0, 60),
    color: normalizeHexColor(item.color)
  };
}

function startEditSubject(subjectId) {
  const subject = getSubjectById(subjectId);
  if (!subject) {
    return;
  }

  ensurePanelExpanded("subject");

  state.editSubjectId = subject.id;
  els.subjectFormTitle.textContent = "Edit Subject";
  els.saveSubjectBtn.textContent = "Save Subject";
  els.cancelSubjectEditBtn.classList.remove("hidden");

  els.subjectFields.name.value = subject.name;
  els.subjectFields.code.value = subject.code;
  els.subjectFields.teacher.value = subject.teacher;
  els.subjectFields.room.value = subject.defaultRoom;
  els.subjectFields.color.value = subject.color;

  clearSubjectFormError();
  setStatus(`Editing subject: ${subject.name}.`, "success");
}

function resetSubjectFormState() {
  state.editSubjectId = null;
  els.subjectForm.reset();
  els.subjectFields.color.value = DEFAULT_SUBJECT_COLOR;
  els.subjectFormTitle.textContent = "Add Subject";
  els.saveSubjectBtn.textContent = "Add Subject";
  els.cancelSubjectEditBtn.classList.add("hidden");
  clearSubjectFormError();
}

function deleteSubject(subjectId) {
  const subject = getSubjectById(subjectId);
  if (!subject) {
    return;
  }

  const linkedClasses = state.classes.filter((entry) => entry.subjectId === subject.id);

  if (linkedClasses.length > 0) {
    if (state.subjects.length <= 1) {
      setStatus("Cannot delete the only subject while classes still reference it.", "error");
      return;
    }

    const replacement = state.subjects.find((entry) => entry.id !== subject.id);
    const confirmedReassign = window.confirm(
      `"${subject.name}" has ${linkedClasses.length} class(es). Delete subject and move these classes to "${replacement.name}"?`
    );

    if (!confirmedReassign) {
      return;
    }

    state.classes = state.classes.map((entry) =>
      entry.subjectId === subject.id
        ? {
            ...entry,
            subjectId: replacement.id
          }
        : entry
    );
  } else {
    const confirmed = window.confirm(`Delete subject "${subject.name}"?`);
    if (!confirmed) {
      return;
    }
  }

  state.subjects = state.subjects.filter((entry) => entry.id !== subject.id);

  if (state.editSubjectId === subject.id) {
    resetSubjectFormState();
  }

  if (state.editClassId) {
    const editingClass = state.classes.find((entry) => entry.id === state.editClassId);
    if (editingClass && editingClass.subjectId === subject.id) {
      resetClassFormState();
    }
  }

  persistAppState();
  renderAll();
  setStatus("Subject deleted.", "success");
}

function handleClassSubmit(event) {
  event.preventDefault();
  clearClassFormError();

  const draft = getClassFromForm();
  const validationError = validateClassDraft(draft);
  if (validationError) {
    setClassFormError(validationError);
    return;
  }

  const normalized = normalizeClassForStorage(draft, state.editClassId);
  const wasEditing = Boolean(state.editClassId);

  if (wasEditing) {
    const index = state.classes.findIndex((entry) => entry.id === state.editClassId);
    if (index >= 0) {
      state.classes[index] = normalized;
    }
  } else {
    state.classes.push(normalized);
  }

  persistAppState();
  renderAll();
  resetClassFormState();
  setStatus(wasEditing ? "Class updated." : "Class added.", "success");
}

function getClassFromForm() {
  const chosenType = els.fields.type.value.trim();
  const customType = els.fields.customType.value.trim();

  return {
    id: state.editClassId,
    subjectId: els.fields.subject.value,
    title: els.fields.title.value.trim(),
    type: chosenType === "Custom" ? customType : chosenType,
    day: els.fields.day.value,
    color: normalizeHexColor(els.fields.color.value),
    useSubjectColor: els.fields.inheritSubjectColor.checked,
    startTime: normalizeTime(els.fields.startTime.value),
    endTime: normalizeTime(els.fields.endTime.value),
    location: els.fields.location.value.trim(),
    teacherOverride: els.fields.teacher.value.trim(),
    notes: els.fields.notes.value.trim()
  };
}

function validateClassDraft(item) {
  if (!state.subjects.length) {
    return "Create at least one subject before adding classes.";
  }

  if (!item.subjectId || !getSubjectById(item.subjectId)) {
    return "Select a valid subject.";
  }

  if (!item.title) {
    return "Session title is required.";
  }

  if (!item.type) {
    return "Class type is required.";
  }

  if (!DAYS.includes(item.day)) {
    return "Select a valid day.";
  }

  if (!isValidTime(item.startTime) || !isValidTime(item.endTime)) {
    return "Enter valid start and end times.";
  }

  if (toMinutes(item.endTime) <= toMinutes(item.startTime)) {
    return "End time must be after start time.";
  }

  return null;
}

function normalizeClassForStorage(item, forcedId) {
  return {
    id: forcedId || createId(),
    subjectId: item.subjectId,
    title: item.title.slice(0, 80),
    type: item.type.slice(0, 40),
    day: item.day,
    color: normalizeHexColor(item.color),
    useSubjectColor: Boolean(item.useSubjectColor),
    startTime: item.startTime,
    endTime: item.endTime,
    location: item.location.slice(0, 60),
    teacherOverride: item.teacherOverride.slice(0, 60),
    notes: item.notes.slice(0, 220)
  };
}

function startEditClass(classId) {
  const entry = state.classes.find((item) => item.id === classId);
  if (!entry) {
    return;
  }

  ensurePanelExpanded("classSession");

  state.editClassId = entry.id;
  els.formTitle.textContent = "Edit Class Session";
  els.saveClassBtn.textContent = "Save Changes";
  els.cancelEditBtn.classList.remove("hidden");

  renderSubjectSelectOptions(entry.subjectId);
  els.fields.subject.value = entry.subjectId;
  els.fields.title.value = entry.title;

  if (PRESET_TYPES.includes(entry.type)) {
    els.fields.type.value = entry.type;
    els.fields.customType.value = "";
  } else {
    els.fields.type.value = "Custom";
    els.fields.customType.value = entry.type;
  }

  els.fields.day.value = entry.day;
  els.fields.color.value = entry.color;
  els.fields.inheritSubjectColor.checked = Boolean(entry.useSubjectColor);
  els.fields.startTime.value = entry.startTime;
  els.fields.endTime.value = entry.endTime;
  els.fields.location.value = entry.location;
  els.fields.teacher.value = entry.teacherOverride;
  els.fields.notes.value = entry.notes;

  updateCustomTypeVisibility();
  updateColorInheritanceUI();
  syncClassDefaultsFromSubject();
  clearClassFormError();
  setStatus("Editing selected class.", "success");
}

function resetClassFormState() {
  state.editClassId = null;
  els.classForm.reset();
  els.fields.startTime.value = "09:00";
  els.fields.endTime.value = "10:00";
  els.fields.inheritSubjectColor.checked = true;
  els.formTitle.textContent = "Add Class Session";
  els.saveClassBtn.textContent = "Add Class";
  els.cancelEditBtn.classList.add("hidden");
  clearClassFormError();
  renderSubjectSelectOptions();
  updateCustomTypeVisibility();
  updateColorInheritanceUI();
  syncClassDefaultsFromSubject();
}

function deleteClass(classId) {
  const entry = state.classes.find((item) => item.id === classId);
  if (!entry) {
    return;
  }

  const resolved = resolveClassDetails(entry);
  const confirmed = window.confirm(`Delete "${resolved.subjectName} - ${entry.title}"?`);
  if (!confirmed) {
    return;
  }

  state.classes = state.classes.filter((item) => item.id !== classId);

  if (state.editClassId === classId) {
    resetClassFormState();
  }

  persistAppState();
  renderAll();
  setStatus("Class deleted.", "success");
}

function clearAllClasses() {
  if (!state.classes.length) {
    setStatus("No classes to clear.", "info");
    return;
  }

  const confirmed = window.confirm("Clear all class sessions from this timetable? This cannot be undone!!");
  if (!confirmed) {
    return;
  }

  state.classes = [];
  persistAppState();
  renderAll();
  resetClassFormState();
  setStatus("All classes cleared.", "success");
}

function updateCustomTypeVisibility() {
  const useCustom = els.typeSelect.value === "Custom";
  els.customTypeWrap.classList.toggle("hidden", !useCustom);
  els.customTypeInput.required = useCustom;

  if (!useCustom) {
    els.customTypeInput.value = "";
  }
}

function updateColorInheritanceUI() {
  const useSubjectColor = els.fields.inheritSubjectColor.checked;
  els.fields.color.disabled = useSubjectColor;
}

function syncClassDefaultsFromSubject() {
  const selectedSubject = getSubjectById(els.fields.subject.value);

  if (!selectedSubject) {
    els.fields.color.value = DEFAULT_SUBJECT_COLOR;
    els.fields.location.placeholder = "Use default room";
    els.fields.teacher.placeholder = "Use subject teacher";
    return;
  }

  if (els.fields.inheritSubjectColor.checked) {
    els.fields.color.value = selectedSubject.color;
  }

  els.fields.location.placeholder = selectedSubject.defaultRoom
    ? `Default room: ${selectedSubject.defaultRoom}`
    : "Use default room";
  els.fields.teacher.placeholder = selectedSubject.teacher
    ? `Default teacher: ${selectedSubject.teacher}`
    : "Use subject teacher";
}

function handleDisplaySettingsChange() {
  const next = {
    showWeekends: els.displayInputs.showWeekends.checked,
    showTeacher: els.displayInputs.showTeacher.checked,
    showLocation: els.displayInputs.showLocation.checked,
    compactCards: els.displayInputs.compactCards.checked,
    dayStartHour: Number(els.displayInputs.dayStartHour.value),
    dayEndHour: Number(els.displayInputs.dayEndHour.value),
    hourHeight: Number(els.displayInputs.hourHeight.value)
  };

  if (!Number.isFinite(next.dayStartHour) || !Number.isFinite(next.dayEndHour)) {
    setSettingsStatus("Invalid day range.", "error");
    applyStateToControls();
    return;
  }

  if (next.dayEndHour <= next.dayStartHour + 1) {
    setSettingsStatus("Day end must be at least 2 hours after day start.", "error");
    applyStateToControls();
    return;
  }

  state.display = sanitizeDisplaySettings(next);
  applyDisplayCssVars();
  persistAppState();
  renderSummary();
  renderTimetable();
  setSettingsStatus("Display settings saved.", "success");
}

function applyDisplayCssVars() {
  const root = document.documentElement;
  const totalHours = state.display.dayEndHour - state.display.dayStartHour;
  const computedDayHeight = totalHours * state.display.hourHeight;

  root.style.setProperty("--hour-height", `${state.display.hourHeight}px`);
  root.style.setProperty("--day-height", `${computedDayHeight}px`);
}

function handleExportSettingsChange() {
  const next = {
    theme: els.exportInputs.theme.value,
    pngScale: Number(els.exportInputs.pngScale.value),
    includeLegend: els.exportInputs.includeLegend.checked,
    includeClassList: els.exportInputs.includeClassList.checked,
    pdfOrientation: els.exportInputs.pdfOrientation.value,
    pdfPageSize: els.exportInputs.pdfPageSize.value,
    filePrefix: els.exportInputs.filePrefix.value
  };

  state.exportPrefs = sanitizeExportSettings(next);
  applyStateToControls();
  persistAppState();
  setSettingsStatus("Export settings saved.", "success");
}

function renderAll() {
  renderSubjectList();
  renderSubjectSelectOptions();
  renderSummary();
  renderClassList();
  renderTimetable();
}

function renderSubjectList() {
  els.subjectList.innerHTML = "";

  if (!state.subjects.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "No subjects yet. Add a subject first, then attach classes to it.";
    els.subjectList.append(empty);
    return;
  }

  const fragment = document.createDocumentFragment();

  state.subjects
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((subject) => {
      const classCount = state.classes.filter((entry) => entry.subjectId === subject.id).length;

      const card = document.createElement("article");
      card.className = "subject-item";

      const main = document.createElement("div");
      main.className = "subject-main";

      const head = document.createElement("div");
      head.className = "subject-head";

      const swatch = document.createElement("span");
      swatch.className = "swatch";
      swatch.style.backgroundColor = subject.color;

      const name = document.createElement("span");
      name.className = "subject-name";
      name.textContent = subject.name;

      head.append(swatch, name);

      if (subject.code) {
        const code = document.createElement("span");
        code.className = "subject-code";
        code.textContent = subject.code;
        head.append(code);
      }

      const meta = document.createElement("p");
      meta.className = "subject-meta";
      meta.textContent = `Teacher: ${subject.teacher || "-"} | Room: ${subject.defaultRoom || "-"} | ${classCount} class(es)`;

      main.append(head, meta);

      const actions = document.createElement("div");
      actions.className = "subject-actions";

      const editBtn = document.createElement("button");
      editBtn.type = "button";
      editBtn.className = "secondary";
      editBtn.dataset.action = "edit-subject";
      editBtn.dataset.id = subject.id;
      editBtn.textContent = "Edit";

      const deleteBtn = document.createElement("button");
      deleteBtn.type = "button";
      deleteBtn.className = "danger";
      deleteBtn.dataset.action = "delete-subject";
      deleteBtn.dataset.id = subject.id;
      deleteBtn.textContent = "Delete";

      actions.append(editBtn, deleteBtn);
      card.append(main, actions);
      fragment.append(card);
    });

  els.subjectList.append(fragment);
}

function renderSubjectSelectOptions(preferredId) {
  const currentValue = preferredId || els.fields.subject.value;
  els.fields.subject.innerHTML = "";

  if (!state.subjects.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "Create a subject first";
    els.fields.subject.append(option);
    els.fields.subject.disabled = true;
    syncClassDefaultsFromSubject();
    return;
  }

  els.fields.subject.disabled = false;

  state.subjects
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((subject) => {
      const option = document.createElement("option");
      option.value = subject.id;
      option.textContent = subject.code ? `${subject.name} (${subject.code})` : subject.name;
      els.fields.subject.append(option);
    });

  const hasCurrent = state.subjects.some((subject) => subject.id === currentValue);
  els.fields.subject.value = hasCurrent ? currentValue : state.subjects[0].id;
  syncClassDefaultsFromSubject();
}

function renderSummary() {
  const totalSubjects = state.subjects.length;
  const totalClasses = state.classes.length;
  const visibleDays = getVisibleDays();

  if (!totalClasses) {
    els.summaryText.textContent =
      totalSubjects > 0
        ? `${totalSubjects} subject(s) set up. Add classes to populate the timetable.`
        : "No subjects or classes yet. Start by creating a subject.";
    return;
  }

  els.summaryText.textContent = `${totalClasses} class session(s) across ${totalSubjects} subject(s), showing ${visibleDays.length} day(s).`;
}

function renderClassList() {
  els.classList.innerHTML = "";

  if (!state.classes.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "No classes added yet. Create subjects and then add sessions.";
    els.classList.append(empty);
    return;
  }

  const fragment = document.createDocumentFragment();

  getSortedClasses().forEach((item) => {
    const resolved = resolveClassDetails(item);

    const card = document.createElement("article");
    card.className = "class-item";

    const main = document.createElement("div");
    main.className = "item-main";

    const line1 = document.createElement("div");
    line1.className = "item-line-1";

    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.backgroundColor = resolved.color;

    const title = document.createElement("span");
    title.className = "item-title";
    title.textContent = item.title;

    const subjectChip = document.createElement("span");
    subjectChip.className = "item-subject";
    subjectChip.textContent = resolved.subjectName;

    line1.append(swatch, title, subjectChip);

    const line2 = document.createElement("p");
    line2.className = "item-meta";
    line2.textContent = `${item.day}, ${item.startTime}-${item.endTime} | ${item.type}`;

    const line3 = document.createElement("p");
    line3.className = "item-meta";
    line3.textContent = `Room: ${resolved.location || "-"} | Teacher: ${resolved.teacher || "-"}`;

    main.append(line1, line2, line3);

    if (item.notes) {
      const notes = document.createElement("p");
      notes.className = "item-notes";
      notes.textContent = `Notes: ${item.notes}`;
      main.append(notes);
    }

    const actions = document.createElement("div");
    actions.className = "item-actions";

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "secondary";
    editBtn.dataset.action = "edit-class";
    editBtn.dataset.id = item.id;
    editBtn.textContent = "Edit";

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "danger";
    deleteBtn.dataset.action = "delete-class";
    deleteBtn.dataset.id = item.id;
    deleteBtn.textContent = "Delete";

    actions.append(editBtn, deleteBtn);
    card.append(main, actions);
    fragment.append(card);
  });

  els.classList.append(fragment);
}

function renderTimetable() {
  const visibleDays = getVisibleDays();
  const dayStartMinutes = state.display.dayStartHour * 60;
  const dayEndMinutes = state.display.dayEndHour * 60;
  const dayTotalMinutes = dayEndMinutes - dayStartMinutes;

  els.timetableGrid.innerHTML = "";
  els.timetableGrid.style.setProperty("--day-columns", String(visibleDays.length));
  els.timetableGrid.style.minWidth = `${78 + visibleDays.length * 145}px`;

  const corner = document.createElement("div");
  corner.className = "grid-corner";
  corner.textContent = "Time";
  els.timetableGrid.append(corner);

  visibleDays.forEach((day) => {
    const header = document.createElement("div");
    header.className = "day-header";
    header.textContent = day;
    els.timetableGrid.append(header);
  });

  const timeColumn = document.createElement("div");
  timeColumn.className = "time-column";

  for (let hour = state.display.dayStartHour; hour <= state.display.dayEndHour; hour += 1) {
    const label = document.createElement("span");
    label.className = "hour-label";
    label.textContent = `${pad2(hour)}:00`;
    label.style.top = `${((hour * 60 - dayStartMinutes) / dayTotalMinutes) * 100}%`;
    timeColumn.append(label);
  }

  els.timetableGrid.append(timeColumn);

  const groupedByDay = groupClassesByDay(visibleDays);

  visibleDays.forEach((day) => {
    const column = document.createElement("div");
    column.className = "day-column";

    const laidOutItems = layoutDayItems(groupedByDay[day]);
    laidOutItems.forEach((item) => {
      const node = createEventNode(item, dayStartMinutes, dayEndMinutes, dayTotalMinutes);
      if (node) {
        column.append(node);
      }
    });

    els.timetableGrid.append(column);
  });
}

function groupClassesByDay(visibleDays) {
  const grouped = {};
  visibleDays.forEach((day) => {
    grouped[day] = [];
  });

  state.classes.forEach((item) => {
    if (grouped[item.day]) {
      grouped[item.day].push(item);
    }
  });

  return grouped;
}

function layoutDayItems(items) {
  if (!items.length) {
    return [];
  }

  const sorted = items
    .map((item) => ({
      ...item,
      startMinutes: toMinutes(item.startTime),
      endMinutes: toMinutes(item.endTime)
    }))
    .sort((a, b) => a.startMinutes - b.startMinutes || a.endMinutes - b.endMinutes);

  let clusterId = -1;
  let clusterEnd = -1;
  const active = [];

  sorted.forEach((item) => {
    if (item.startMinutes >= clusterEnd) {
      clusterId += 1;
      clusterEnd = item.endMinutes;
    } else {
      clusterEnd = Math.max(clusterEnd, item.endMinutes);
    }

    item.clusterId = clusterId;

    for (let index = active.length - 1; index >= 0; index -= 1) {
      if (active[index].endMinutes <= item.startMinutes) {
        active.splice(index, 1);
      }
    }

    // Keep overlapping classes visible by assigning each overlap to its own lane.
    const occupiedLanes = new Set(active.map((entry) => entry.lane));
    let lane = 0;
    while (occupiedLanes.has(lane)) {
      lane += 1;
    }

    item.lane = lane;
    active.push(item);
  });

  const lanesByCluster = new Map();
  sorted.forEach((item) => {
    const previousMax = lanesByCluster.get(item.clusterId) || 0;
    lanesByCluster.set(item.clusterId, Math.max(previousMax, item.lane + 1));
  });

  sorted.forEach((item) => {
    item.clusterLanes = lanesByCluster.get(item.clusterId) || 1;
  });

  return sorted;
}

function createEventNode(item, dayStartMinutes, dayEndMinutes, dayTotalMinutes) {
  const visibleStart = Math.max(item.startMinutes, dayStartMinutes);
  const visibleEnd = Math.min(item.endMinutes, dayEndMinutes);

  if (visibleEnd <= visibleStart) {
    return null;
  }

  const resolved = resolveClassDetails(item);
  const topPct = ((visibleStart - dayStartMinutes) / dayTotalMinutes) * 100;
  const heightPct = ((visibleEnd - visibleStart) / dayTotalMinutes) * 100;
  const leftPct = (item.lane / item.clusterLanes) * 100;
  const widthPct = 100 / item.clusterLanes;

  const node = document.createElement("article");
  node.className = "class-event";

  if (state.display.compactCards || heightPct < 9) {
    node.classList.add("compact");
  }

  node.style.top = `${topPct}%`;
  node.style.height = `calc(${heightPct}% - 3px)`;
  node.style.left = `calc(${leftPct}% + 2px)`;
  node.style.width = `calc(${widthPct}% - 4px)`;

  const foreground = pickReadableTextColor(resolved.color);
  node.style.color = foreground;
  node.style.background = `linear-gradient(160deg, ${hexToRgba(resolved.color, 0.9)}, ${hexToRgba(
    resolved.color,
    0.68
  )})`;

  const subjectName = document.createElement("h4");
  subjectName.textContent = resolved.subjectName;

  const sessionTitle = document.createElement("p");
  sessionTitle.className = "event-title";
  sessionTitle.textContent = item.title;

  const typeAndTime = document.createElement("p");
  typeAndTime.className = "event-type";
  typeAndTime.textContent = `${item.startTime}-${item.endTime} | ${item.type}`;

  node.append(subjectName);

  if (!isDuplicateSessionTitle(item.title, resolved.subjectName)) {
    node.append(sessionTitle);
  }

  node.append(typeAndTime);

  if (state.display.showLocation) {
    const location = document.createElement("p");
    location.className = "event-location";
    location.textContent = resolved.location || "Room: -";
    node.append(location);
  }

  if (state.display.showTeacher) {
    const teacher = document.createElement("p");
    teacher.className = "event-teacher";
    teacher.textContent = resolved.teacher || "Teacher: -";
    node.append(teacher);
  }

  node.title = `${resolved.subjectName}\n${item.title}\n${item.day} ${item.startTime}-${item.endTime}\nType: ${item.type}\nRoom: ${resolved.location || "-"}\nTeacher: ${resolved.teacher || "-"}`;

  return node;
}

function resolveClassDetails(item) {
  const subject = getSubjectById(item.subjectId);
  const subjectName = subject ? subject.name : "Unknown Subject";
  const color = item.useSubjectColor
    ? subject
      ? subject.color
      : normalizeHexColor(item.color)
    : normalizeHexColor(item.color);

  return {
    subjectName,
    subjectCode: subject ? subject.code : "",
    teacher: item.teacherOverride || (subject ? subject.teacher : ""),
    location: item.location || (subject ? subject.defaultRoom : ""),
    color
  };
}

function isDuplicateSessionTitle(title, subjectName) {
  return title.trim().toLowerCase() === subjectName.trim().toLowerCase();
}

function getSortedClasses() {
  return [...state.classes].sort((a, b) => {
    const dayDiff = DAYS.indexOf(a.day) - DAYS.indexOf(b.day);
    if (dayDiff !== 0) {
      return dayDiff;
    }

    const timeDiff = toMinutes(a.startTime) - toMinutes(b.startTime);
    if (timeDiff !== 0) {
      return timeDiff;
    }

    const aSubject = getSubjectById(a.subjectId);
    const bSubject = getSubjectById(b.subjectId);
    const subjectDiff = (aSubject ? aSubject.name : "").localeCompare(bSubject ? bSubject.name : "");
    if (subjectDiff !== 0) {
      return subjectDiff;
    }

    return a.title.localeCompare(b.title);
  });
}

function getVisibleDays() {
  return state.display.showWeekends ? DAYS : DAYS.slice(0, 5);
}

function getSubjectById(subjectId) {
  return state.subjects.find((subject) => subject.id === subjectId) || null;
}

function handleExportTheme() {
  return state.exportPrefs.theme === "current" ? state.theme : state.exportPrefs.theme;
}

async function exportAsPng() {
  if (!state.classes.length) {
    setStatus("Add at least one class before exporting.", "error");
    return;
  }

  try {
    setStatus("Rendering PNG sheet...", "info");
    const { canvas, exportTheme } = await captureExportCanvas();
    const blob = await canvasToBlob(canvas, "image/png");
    downloadBlob(blob, `${getFilePrefix()}-${exportTheme}-${buildStamp()}.png`);
    setStatus(`PNG downloaded (${capitalize(exportTheme)} theme).`, "success");
  } catch (error) {
    setStatus(`PNG export failed: ${error.message}`, "error");
  }
}

async function exportAsPdf() {
  if (!state.classes.length) {
    setStatus("Add at least one class before exporting.", "error");
    return;
  }

  try {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error("PDF library not loaded.");
    }

    setStatus("Building PDF...", "info");

    const { canvas, exportTheme } = await captureExportCanvas();
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
      orientation: state.exportPrefs.pdfOrientation,
      unit: "pt",
      format: state.exportPrefs.pdfPageSize
    });

    const margin = 24;
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const targetWidth = pageWidth - margin * 2;
    const printableHeight = pageHeight - margin * 2 - 14;

    const ratio = targetWidth / canvas.width;
    const fullHeightOnPage = canvas.height * ratio;

    pdf.setFontSize(12);
    pdf.text(`Timetable Studio (${capitalize(exportTheme)} theme)`, margin, margin - 6);

    if (fullHeightOnPage <= printableHeight) {
      const imageData = canvas.toDataURL("image/png");
      pdf.addImage(imageData, "PNG", margin, margin + 8, targetWidth, fullHeightOnPage, undefined, "FAST");
    } else {
      // Slice large snapshots into multiple pages so text remains readable.
      const sliceHeightPx = Math.max(1, Math.floor(printableHeight / ratio));
      let renderedPx = 0;
      let pageIndex = 0;

      while (renderedPx < canvas.height) {
        const currentSliceHeight = Math.min(sliceHeightPx, canvas.height - renderedPx);
        const sliceCanvas = document.createElement("canvas");
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = currentSliceHeight;

        const context = sliceCanvas.getContext("2d");
        if (!context) {
          throw new Error("Could not prepare PDF image slice.");
        }

        context.drawImage(
          canvas,
          0,
          renderedPx,
          canvas.width,
          currentSliceHeight,
          0,
          0,
          canvas.width,
          currentSliceHeight
        );

        if (pageIndex > 0) {
          pdf.addPage();
          pdf.setFontSize(11);
          pdf.text(`Timetable Studio (${capitalize(exportTheme)} theme)`, margin, margin - 6);
        }

        const imageData = sliceCanvas.toDataURL("image/png");
        const drawHeight = currentSliceHeight * ratio;
        pdf.addImage(imageData, "PNG", margin, margin + 8, targetWidth, drawHeight, undefined, "FAST");

        renderedPx += currentSliceHeight;
        pageIndex += 1;
      }
    }

    pdf.save(`${getFilePrefix()}-${buildStamp()}.pdf`);
    setStatus("PDF downloaded.", "success");
  } catch (error) {
    setStatus(`PDF export failed: ${error.message}`, "error");
  }
}

function exportAsIcs() {
  if (!state.classes.length) {
    setStatus("Add at least one class before exporting.", "error");
    return;
  }

  try {
    const now = new Date();
    const stamp = toIcsUtc(now);
    const lines = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//TimetableStudio//EN",
      "CALSCALE:GREGORIAN",
      "METHOD:PUBLISH",
      "X-WR-CALNAME:Timetable Studio"
    ];

    getSortedClasses().forEach((item) => {
      const resolved = resolveClassDetails(item);
      const eventDate = getNextDateForDay(item.day);
      const startDate = mergeDateAndTime(eventDate, item.startTime);
      const endDate = mergeDateAndTime(eventDate, item.endTime);

      const summary = `${resolved.subjectName} - ${item.title}`;

      lines.push("BEGIN:VEVENT");
      lines.push(`UID:${escapeIcsText(item.id)}@timetablestudio.local`);
      lines.push(`DTSTAMP:${stamp}`);
      lines.push(`SUMMARY:${escapeIcsText(summary)}`);
      lines.push(`DTSTART:${toIcsLocal(startDate)}`);
      lines.push(`DTEND:${toIcsLocal(endDate)}`);
      lines.push(`RRULE:FREQ=WEEKLY;COUNT=${SEMESTER_WEEKS}`);
      lines.push(`LOCATION:${escapeIcsText(resolved.location || "TBD")}`);
      lines.push(
        `DESCRIPTION:${escapeIcsText(
          `Type: ${item.type}\nTeacher: ${resolved.teacher || "TBD"}\nNotes: ${item.notes || "-"}`
        )}`
      );
      lines.push("END:VEVENT");
    });

    lines.push("END:VCALENDAR", "");

    const blob = new Blob([lines.join("\r\n")], {
      type: "text/calendar;charset=utf-8"
    });

    downloadBlob(blob, `${getFilePrefix()}-${buildStamp()}.ics`);
    setStatus("ICS calendar downloaded.", "success");
  } catch (error) {
    setStatus(`ICS export failed: ${error.message}`, "error");
  }
}

function exportAsJson() {
  const payload = {
    version: 2,
    generatedAt: new Date().toISOString(),
    subjects: state.subjects,
    classes: getSortedClasses(),
    displaySettings: state.display,
    exportSettings: state.exportPrefs
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json;charset=utf-8"
  });

  downloadBlob(blob, `${getFilePrefix()}-${buildStamp()}.json`);
  setStatus("JSON exported.", "success");
}

function importFromJson(event) {
  const [file] = event.target.files || [];
  event.target.value = "";

  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || ""));
      const extracted = extractStateFromPayload(parsed);

      if (!extracted) {
        throw new Error("Unsupported JSON format. Expected version 2 payload or legacy class list.");
      }

      if (!extracted.classes.length && !extracted.subjects.length) {
        throw new Error("No valid subjects or classes found in file.");
      }

      const confirmed = window.confirm(
        `Import ${extracted.subjects.length} subject(s) and ${extracted.classes.length} class(es), replacing current data?`
      );
      if (!confirmed) {
        return;
      }

      state.subjects = extracted.subjects;
      state.classes = extracted.classes;
      state.display = extracted.display;
      state.exportPrefs = extracted.exportPrefs;
      state.editClassId = null;
      state.editSubjectId = null;

      applyStateToControls();
      applyDisplayCssVars();
      resetSubjectFormState();
      resetClassFormState();
      persistAppState();
      renderAll();
      setStatus(`Imported ${extracted.classes.length} classes.`, "success");
      setSettingsStatus("Imported display/export settings from file.", "success");
    } catch (error) {
      setStatus(`Import failed: ${error.message}`, "error");
    }
  };

  reader.onerror = () => {
    setStatus("Import failed: could not read file.", "error");
  };

  reader.readAsText(file);
}

async function captureExportCanvas() {
  if (typeof window.html2canvas !== "function") {
    throw new Error("Image export library not loaded.");
  }

  // Temporarily switch theme only for export, then restore the current UI theme.
  const exportTheme = handleExportTheme();
  const originalTheme = state.theme;
  const shouldRestoreTheme = exportTheme !== originalTheme;

  if (shouldRestoreTheme) {
    applyTheme(exportTheme, { persist: false });
    await waitForPaint();
  }

  const snapshot = buildExportSnapshot();
  document.body.append(snapshot);
  await waitForPaint();

  try {
    const canvas = await window.html2canvas(snapshot, {
      scale: state.exportPrefs.pngScale,
      useCORS: true,
      backgroundColor: getComputedStyle(snapshot).backgroundColor,
      logging: false
    });

    return {
      canvas,
      exportTheme
    };
  } finally {
    snapshot.remove();

    if (shouldRestoreTheme) {
      applyTheme(originalTheme, { persist: false });
      await waitForPaint();
    }
  }
}

function buildExportSnapshot() {
  const wrapper = document.createElement("section");
  wrapper.className = "export-snapshot";

  const sheet = document.createElement("article");
  sheet.className = "export-sheet";

  const header = document.createElement("header");
  header.className = "export-sheet-header";

  const titleWrap = document.createElement("div");

  const title = document.createElement("h2");
  title.className = "export-title";
  title.textContent = "Timetable Studio";

  const subtitle = document.createElement("p");
  subtitle.className = "export-subtitle";
  subtitle.textContent = "Subject-first weekly timetable";

  titleWrap.append(title, subtitle);

  const meta = document.createElement("div");
  meta.className = "export-meta";
  meta.innerHTML = `<div>${new Date().toLocaleString()}</div><div>${state.subjects.length} subject(s), ${state.classes.length} class(es)</div>`;

  header.append(titleWrap, meta);

  const main = document.createElement("div");
  main.className = "export-sheet-main";

  const gridWrap = document.createElement("div");
  gridWrap.className = "timetable-capture";

  const gridClone = els.timetableGrid.cloneNode(true);
  gridClone.removeAttribute("id");
  gridWrap.append(gridClone);
  main.append(gridWrap);

  if (state.exportPrefs.includeLegend) {
    main.append(buildExportLegend());
  }

  if (state.exportPrefs.includeClassList) {
    main.append(buildExportClassTable());
  }

  sheet.append(header, main);
  wrapper.append(sheet);
  return wrapper;
}

function buildExportLegend() {
  const container = document.createElement("section");
  container.className = "export-legend";

  const heading = document.createElement("h3");
  heading.textContent = "Subject Legend";

  const list = document.createElement("div");
  list.className = "export-legend-list";

  state.subjects
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach((subject) => {
      const item = document.createElement("div");
      item.className = "export-legend-item";

      const name = document.createElement("strong");
      name.textContent = subject.code ? `${subject.name} (${subject.code})` : subject.name;
      name.style.color = pickReadableTextColor(subject.color);
      name.style.background = `linear-gradient(180deg, ${hexToRgba(subject.color, 0.86)}, ${hexToRgba(
        subject.color,
        0.56
      )})`;
      name.style.padding = "4px 6px";
      name.style.borderRadius = "8px";

      const details = document.createElement("span");
      details.textContent = `Teacher: ${subject.teacher || "-"} | Room: ${subject.defaultRoom || "-"}`;

      item.append(name, details);
      list.append(item);
    });

  container.append(heading, list);
  return container;
}

function buildExportClassTable() {
  const wrap = document.createElement("section");
  wrap.className = "export-class-table-wrap";

  const heading = document.createElement("h3");
  heading.textContent = "Class Sessions";

  const table = document.createElement("table");
  table.className = "export-class-table";

  const thead = document.createElement("thead");
  thead.innerHTML =
    "<tr><th>Day</th><th>Time</th><th>Subject</th><th>Session</th><th>Type</th><th>Room</th><th>Teacher</th></tr>";

  const tbody = document.createElement("tbody");

  getSortedClasses().forEach((entry) => {
    const resolved = resolveClassDetails(entry);
    const row = document.createElement("tr");

    row.innerHTML = `<td>${escapeHtml(entry.day)}</td><td>${escapeHtml(entry.startTime)}-${escapeHtml(
      entry.endTime
    )}</td><td>${escapeHtml(resolved.subjectName)}</td><td>${escapeHtml(entry.title)}</td><td>${escapeHtml(
      entry.type
    )}</td><td>${escapeHtml(resolved.location || "-")}</td><td>${escapeHtml(resolved.teacher || "-")}</td>`;

    tbody.append(row);
  });

  table.append(thead, tbody);
  wrap.append(heading, table);
  return wrap;
}

function toggleTheme() {
  const next = state.theme === "light" ? "dark" : "light";
  applyTheme(next, { persist: true });
  setStatus(`Theme set to ${capitalize(next)}.`, "success");
}

function applyTheme(theme, options = {}) {
  const { persist } = { persist: true, ...options };
  if (theme !== "light" && theme !== "dark") {
    return;
  }

  state.theme = theme;
  document.documentElement.setAttribute("data-theme", theme);
  els.themeToggle.textContent = theme === "light" ? "Switch to Dark" : "Switch to Light";

  if (persist) {
    localStorage.setItem(THEME_KEY, theme);
  }
}

function hydrateTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  if (saved === "light" || saved === "dark") {
    state.theme = saved;
  }

  applyTheme(state.theme, { persist: false });
}

function hydratePanelState() {
  const saved = localStorage.getItem(PANEL_STATE_KEY);
  if (!saved) {
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    if (!parsed || typeof parsed !== "object") {
      return;
    }

    state.panelState = {
      subject:
        typeof parsed.subject === "boolean" ? parsed.subject : DEFAULT_PANEL_STATE.subject,
      classSession:
        typeof parsed.classSession === "boolean"
          ? parsed.classSession
          : DEFAULT_PANEL_STATE.classSession,
      settings:
        typeof parsed.settings === "boolean" ? parsed.settings : DEFAULT_PANEL_STATE.settings
    };
  } catch {
    state.panelState = { ...DEFAULT_PANEL_STATE };
  }
}

function persistPanelState() {
  localStorage.setItem(PANEL_STATE_KEY, JSON.stringify(state.panelState));
}

function applyPanelStates() {
  Object.keys(DEFAULT_PANEL_STATE).forEach((panelKey) => {
    applyPanelState(panelKey);
  });
}

function applyPanelState(panelKey) {
  const panel = document.querySelector(`.collapsible-panel[data-panel-key="${panelKey}"]`);
  if (!panel) {
    return;
  }

  // The `hidden` attribute is the source of truth for collapsed panel bodies.
  const isExpanded = Boolean(state.panelState[panelKey]);
  panel.classList.toggle("is-collapsed", !isExpanded);

  const toggle = panel.querySelector("[data-panel-toggle]");
  if (toggle) {
    toggle.setAttribute("aria-expanded", String(isExpanded));
    const text = toggle.querySelector(".panel-toggle-text");
    if (text) {
      text.textContent = isExpanded ? "Collapse" : "Expand";
    }
  }

  const body = panel.querySelector(".panel-collapsible-body");
  if (body) {
    body.hidden = !isExpanded;
  }
}

function togglePanel(panelKey) {
  if (!Object.prototype.hasOwnProperty.call(state.panelState, panelKey)) {
    return;
  }

  state.panelState[panelKey] = !state.panelState[panelKey];
  applyPanelState(panelKey);
  persistPanelState();
}

function ensurePanelExpanded(panelKey) {
  if (!Object.prototype.hasOwnProperty.call(state.panelState, panelKey)) {
    return;
  }

  if (state.panelState[panelKey]) {
    return;
  }

  state.panelState[panelKey] = true;
  applyPanelState(panelKey);
  persistPanelState();
}

function hydrateAppData() {
  const currentRaw = localStorage.getItem(STORAGE_KEY);
  if (currentRaw) {
    try {
      const parsed = JSON.parse(currentRaw);
      const extracted = extractStateFromPayload(parsed);
      if (extracted) {
        applyExtractedState(extracted);
        return;
      }
    } catch {
      setStatus("Could not load saved data. Storage payload was invalid.", "error");
    }
  }

  const legacyRaw = localStorage.getItem(LEGACY_STORAGE_KEY);
  if (!legacyRaw) {
    return;
  }

  try {
    const legacyParsed = JSON.parse(legacyRaw);
    const extracted = extractStateFromPayload(legacyParsed);
    if (!extracted) {
      return;
    }

    applyExtractedState(extracted);
    persistAppState();
    setStatus("Migrated your previous timetable into the new subject-based format.", "success");
  } catch {
    setStatus("Could not migrate previous saved data.", "error");
  }
}

function extractStateFromPayload(payload) {
  if (Array.isArray(payload)) {
    return migrateLegacyClassList(payload);
  }

  if (!payload || typeof payload !== "object") {
    return null;
  }

  const hasModernArrays = Array.isArray(payload.subjects) && Array.isArray(payload.classes);

  if (hasModernArrays) {
    const subjects = payload.subjects.map(sanitizeSubject).filter(Boolean);
    let classes = payload.classes.map(sanitizeClass).filter(Boolean);

    if (!subjects.length && classes.length) {
      const fallback = {
        id: createId(),
        name: "Imported Subject",
        code: "",
        teacher: "",
        defaultRoom: "",
        color: DEFAULT_SUBJECT_COLOR
      };

      subjects.push(fallback);
      classes = classes.map((entry) => ({
        ...entry,
        subjectId: fallback.id
      }));
    }

    const subjectIds = new Set(subjects.map((subject) => subject.id));
    classes = classes
      .map((entry) => {
        if (subjectIds.has(entry.subjectId)) {
          return entry;
        }

        if (subjects[0]) {
          return {
            ...entry,
            subjectId: subjects[0].id
          };
        }

        return null;
      })
      .filter(Boolean);

    return {
      subjects,
      classes,
      display: sanitizeDisplaySettings(payload.displaySettings || payload.display),
      exportPrefs: sanitizeExportSettings(payload.exportSettings || payload.exportPrefs)
    };
  }

  if (Array.isArray(payload.classes)) {
    const migrated = migrateLegacyClassList(payload.classes);
    migrated.display = sanitizeDisplaySettings(payload.displaySettings || payload.display);
    migrated.exportPrefs = sanitizeExportSettings(payload.exportSettings || payload.exportPrefs);
    return migrated;
  }

  return null;
}

function migrateLegacyClassList(list) {
  const source = Array.isArray(list) ? list : [];
  const subjectMap = new Map();
  const migratedClasses = [];

  source.forEach((item) => {
    const legacy = sanitizeLegacyClass(item);
    if (!legacy) {
      return;
    }

    const key = legacy.title.trim().toLowerCase();
    let subject = subjectMap.get(key);

    if (!subject) {
      subject = {
        id: createId(),
        name: legacy.title,
        code: "",
        teacher: legacy.lecturer || "",
        defaultRoom: legacy.location || "",
        color: legacy.color
      };

      subjectMap.set(key, subject);
    }

    migratedClasses.push({
      id: legacy.id,
      subjectId: subject.id,
      title: legacy.title,
      type: legacy.type,
      day: legacy.day,
      color: legacy.color,
      useSubjectColor: true,
      startTime: legacy.startTime,
      endTime: legacy.endTime,
      location: legacy.location,
      teacherOverride: legacy.lecturer,
      notes: legacy.notes
    });
  });

  return {
    subjects: Array.from(subjectMap.values()),
    classes: migratedClasses,
    display: { ...DEFAULT_DISPLAY_SETTINGS },
    exportPrefs: { ...DEFAULT_EXPORT_SETTINGS }
  };
}

function sanitizeLegacyClass(item) {
  if (!item || typeof item !== "object") {
    return null;
  }

  const title = typeof item.title === "string" ? item.title.trim() : "";
  const type = typeof item.type === "string" ? item.type.trim() : "Class";
  const day = typeof item.day === "string" ? item.day : "";
  const startTime = normalizeTime(item.startTime);
  const endTime = normalizeTime(item.endTime);

  if (!title || !DAYS.includes(day) || !isValidTime(startTime) || !isValidTime(endTime)) {
    return null;
  }

  if (toMinutes(endTime) <= toMinutes(startTime)) {
    return null;
  }

  return {
    id: typeof item.id === "string" && item.id ? item.id : createId(),
    title: title.slice(0, 80),
    type: type.slice(0, 40) || "Class",
    day,
    color: normalizeHexColor(typeof item.color === "string" ? item.color : DEFAULT_SUBJECT_COLOR),
    startTime,
    endTime,
    location: typeof item.location === "string" ? item.location.slice(0, 60).trim() : "",
    lecturer: typeof item.lecturer === "string" ? item.lecturer.slice(0, 60).trim() : "",
    notes: typeof item.notes === "string" ? item.notes.slice(0, 220).trim() : ""
  };
}

function sanitizeSubject(item) {
  if (!item || typeof item !== "object") {
    return null;
  }

  const name = typeof item.name === "string" ? item.name.trim() : "";
  if (!name) {
    return null;
  }

  return {
    id: typeof item.id === "string" && item.id ? item.id : createId(),
    name: name.slice(0, 80),
    code: typeof item.code === "string" ? item.code.slice(0, 20).trim() : "",
    teacher: typeof item.teacher === "string" ? item.teacher.slice(0, 60).trim() : "",
    defaultRoom: typeof item.defaultRoom === "string" ? item.defaultRoom.slice(0, 60).trim() : "",
    color: normalizeHexColor(typeof item.color === "string" ? item.color : DEFAULT_SUBJECT_COLOR)
  };
}

function sanitizeClass(item) {
  if (!item || typeof item !== "object") {
    return null;
  }

  const title = typeof item.title === "string" ? item.title.trim() : "";
  const type = typeof item.type === "string" ? item.type.trim() : "";
  const day = typeof item.day === "string" ? item.day : "";
  const startTime = normalizeTime(item.startTime);
  const endTime = normalizeTime(item.endTime);

  if (!title || !type || !DAYS.includes(day) || !isValidTime(startTime) || !isValidTime(endTime)) {
    return null;
  }

  if (toMinutes(endTime) <= toMinutes(startTime)) {
    return null;
  }

  return {
    id: typeof item.id === "string" && item.id ? item.id : createId(),
    subjectId: typeof item.subjectId === "string" ? item.subjectId : "",
    title: title.slice(0, 80),
    type: type.slice(0, 40),
    day,
    color: normalizeHexColor(typeof item.color === "string" ? item.color : DEFAULT_SUBJECT_COLOR),
    useSubjectColor: item.useSubjectColor !== false,
    startTime,
    endTime,
    location: typeof item.location === "string" ? item.location.slice(0, 60).trim() : "",
    teacherOverride:
      typeof item.teacherOverride === "string"
        ? item.teacherOverride.slice(0, 60).trim()
        : typeof item.teacher === "string"
          ? item.teacher.slice(0, 60).trim()
          : "",
    notes: typeof item.notes === "string" ? item.notes.slice(0, 220).trim() : ""
  };
}

function sanitizeDisplaySettings(raw) {
  const base = {
    ...DEFAULT_DISPLAY_SETTINGS,
    ...(raw && typeof raw === "object" ? raw : {})
  };

  const dayStartHour = clampInt(base.dayStartHour, 6, 21);
  const dayEndHour = clampInt(base.dayEndHour, dayStartHour + 2, 23);
  const hourHeight = [46, 56, 68].includes(Number(base.hourHeight)) ? Number(base.hourHeight) : 56;

  return {
    showWeekends: Boolean(base.showWeekends),
    showTeacher: Boolean(base.showTeacher),
    showLocation: Boolean(base.showLocation),
    compactCards: Boolean(base.compactCards),
    dayStartHour,
    dayEndHour,
    hourHeight
  };
}

function sanitizeExportSettings(raw) {
  const base = {
    ...DEFAULT_EXPORT_SETTINGS,
    ...(raw && typeof raw === "object" ? raw : {})
  };

  return {
    theme: ["current", "light", "dark"].includes(base.theme) ? base.theme : "current",
    pngScale: [2, 3, 4].includes(Number(base.pngScale)) ? Number(base.pngScale) : 3,
    includeLegend: Boolean(base.includeLegend),
    includeClassList: Boolean(base.includeClassList),
    pdfOrientation: ["landscape", "portrait"].includes(base.pdfOrientation)
      ? base.pdfOrientation
      : "landscape",
    pdfPageSize: ["a4", "a3", "letter"].includes(base.pdfPageSize) ? base.pdfPageSize : "a4",
    filePrefix: sanitizeFilePrefix(base.filePrefix)
  };
}

function applyExtractedState(extracted) {
  state.subjects = extracted.subjects;
  state.classes = extracted.classes;
  state.display = extracted.display;
  state.exportPrefs = extracted.exportPrefs;
}

function persistAppState() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      version: 2,
      subjects: state.subjects,
      classes: state.classes,
      displaySettings: state.display,
      exportSettings: state.exportPrefs
    })
  );
}

function setStatus(message, tone) {
  els.exportStatus.textContent = message;
  els.exportStatus.dataset.tone = tone;
}

function setSettingsStatus(message, tone) {
  els.settingsStatus.textContent = message;
  els.settingsStatus.dataset.tone = tone;
}

function setClassFormError(message) {
  els.formError.textContent = message;
}

function clearClassFormError() {
  els.formError.textContent = "";
}

function setSubjectFormError(message) {
  els.subjectFormError.textContent = message;
}

function clearSubjectFormError() {
  els.subjectFormError.textContent = "";
}

function getFilePrefix() {
  return sanitizeFilePrefix(state.exportPrefs.filePrefix) || "timetable";
}

function sanitizeFilePrefix(value) {
  const text = typeof value === "string" ? value.trim().toLowerCase() : "";
  const cleaned = text.replace(/[^a-z0-9-_]+/g, "-").replace(/-{2,}/g, "-").replace(/^-|-$/g, "");
  return cleaned.slice(0, 40) || "timetable";
}

function downloadBlob(blob, fileName) {
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);

  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1200);
}

function waitForPaint() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
}

function canvasToBlob(canvas, mimeType) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Could not serialize image."));
        return;
      }
      resolve(blob);
    }, mimeType);
  });
}

function toMinutes(time) {
  const [hours, minutes] = time.split(":").map((value) => Number(value));
  return hours * 60 + minutes;
}

function normalizeTime(value) {
  if (typeof value !== "string") {
    return "";
  }

  const match = value.trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!match) {
    return "";
  }

  const hours = Number(match[1]);
  const minutes = Number(match[2]);

  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
    return "";
  }

  return `${pad2(hours)}:${pad2(minutes)}`;
}

function isValidTime(value) {
  return /^([01]\d|2[0-3]):([0-5]\d)$/.test(value);
}

function normalizeHexColor(value) {
  if (typeof value !== "string") {
    return DEFAULT_SUBJECT_COLOR;
  }

  const shortHexMatch = value.match(/^#([\da-fA-F]{3})$/);
  if (shortHexMatch) {
    const [r, g, b] = shortHexMatch[1].split("");
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }

  if (/^#[\da-fA-F]{6}$/.test(value)) {
    return value.toLowerCase();
  }

  return DEFAULT_SUBJECT_COLOR;
}

function hexToRgba(hex, alpha) {
  const normalized = normalizeHexColor(hex).replace("#", "");
  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function pickReadableTextColor(hex) {
  const normalized = normalizeHexColor(hex).replace("#", "");
  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.58 ? "#0f1d30" : "#f4f7ff";
}

function createId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function buildStamp() {
  const now = new Date();
  return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
}

function dayToJsIndex(dayName) {
  const dayIndex = DAYS.indexOf(dayName);
  if (dayIndex < 0) {
    return 1;
  }
  return dayIndex === 6 ? 0 : dayIndex + 1;
}

function getNextDateForDay(dayName) {
  const today = new Date();
  const targetJsDay = dayToJsIndex(dayName);
  const result = new Date(today);
  result.setHours(0, 0, 0, 0);

  const diff = (targetJsDay - today.getDay() + 7) % 7;
  result.setDate(today.getDate() + diff);
  return result;
}

function mergeDateAndTime(date, time) {
  const [hours, minutes] = time.split(":").map((value) => Number(value));
  const merged = new Date(date);
  merged.setHours(hours, minutes, 0, 0);
  return merged;
}

function toIcsLocal(date) {
  return `${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}T${pad2(
    date.getHours()
  )}${pad2(date.getMinutes())}00`;
}

function toIcsUtc(date) {
  return `${date.getUTCFullYear()}${pad2(date.getUTCMonth() + 1)}${pad2(date.getUTCDate())}T${pad2(
    date.getUTCHours()
  )}${pad2(date.getUTCMinutes())}${pad2(date.getUTCSeconds())}Z`;
}

function escapeIcsText(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/,/g, "\\,")
    .replace(/;/g, "\\;");
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function clampInt(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return min;
  }

  return Math.min(max, Math.max(min, Math.round(number)));
}
