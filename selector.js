(function () {
  "use strict";

  const STORAGE_KEY = "student-selector-session-v1";
  const DEFAULT_STATE = {
    selectedClass: "",
    timerSeconds: 5,
    soundEnabled: true,
    slotEffectEnabled: true,
    selectedStudentsByClass: {},
    studentGradesByClass: {},
    studentUngradedByClass: {},
    absentStudentsByClass: {}
  };

  const AUDIO = {
    intro: "assets/welcome.mp3",
    closing: "assets/closing.mp3",
    slotShort: "assets/select_student.mp3",
    slotMedium: "assets/medium_slot.mp3",
    slotLong: "assets/long_slot.mp3",
    timeup: "assets/timeup.mp3",
    ratings: {
      "A*": "assets/sound_a_star.mp3",
      A: "assets/sound_a.mp3",
      B: "assets/sound_b.mp3",
      C: "assets/sound_c.mp3"
    }
  };

  const RATINGS = [
    ["A*", "Excellent"],
    ["A", "Strong"],
    ["B", "Secure"],
    ["C", "Needs support"]
  ];

  function getScriptBase() {
    const script =
      document.currentScript ||
      Array.from(document.scripts).find((item) => (item.src || "").endsWith("selector.js"));
    if (!script || !script.src) return "";
    return script.src.slice(0, script.src.lastIndexOf("/") + 1);
  }

  const SCRIPT_BASE = getScriptBase();

  function assetUrl(path, basePath) {
    if (/^(https?:)?\/\//.test(path) || path.startsWith("data:")) return path;
    return new URL(path, basePath || SCRIPT_BASE || window.location.href).toString();
  }

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let cell = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
      const ch = text[i];
      const next = text[i + 1];

      if (ch === "\"") {
        if (inQuotes && next === "\"") {
          cell += "\"";
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }

      if (ch === "," && !inQuotes) {
        row.push(cell);
        cell = "";
        continue;
      }

      if ((ch === "\n" || ch === "\r") && !inQuotes) {
        if (ch === "\r" && next === "\n") i += 1;
        row.push(cell);
        if (row.some((value) => value.trim() !== "")) rows.push(row);
        row = [];
        cell = "";
        continue;
      }

      cell += ch;
    }

    row.push(cell);
    if (row.some((value) => value.trim() !== "")) rows.push(row);
    return rows;
  }

  function loadStoredState() {
    try {
      const raw = window.sessionStorage.getItem(STORAGE_KEY);
      if (!raw) return { ...DEFAULT_STATE };
      return { ...DEFAULT_STATE, ...JSON.parse(raw) };
    } catch (_error) {
      return { ...DEFAULT_STATE };
    }
  }

  function uniquePush(list, value) {
    if (value && !list.includes(value)) list.push(value);
  }

  function removeFrom(list, value) {
    const index = list.indexOf(value);
    if (index >= 0) list.splice(index, 1);
  }

  function choice(list) {
    return list[Math.floor(Math.random() * list.length)];
  }

  function formatSeconds(totalSeconds) {
    const total = Math.max(1, Math.round(Number(totalSeconds) || 1));
    const minutes = Math.floor(total / 60);
    const seconds = total % 60;
    if (minutes && seconds) return `${minutes}m ${String(seconds).padStart(2, "0")}s`;
    if (minutes) return `${minutes} ${minutes === 1 ? "min" : "mins"}`;
    return `${seconds} sec`;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll("\"", "&quot;")
      .replaceAll("'", "&#039;");
  }

  class SoundManager {
    constructor(app) {
      this.app = app;
      this.current = null;
      this.cache = new Map();
    }

    get enabled() {
      return Boolean(this.app.state.soundEnabled);
    }

    audio(path) {
      const url = assetUrl(path, this.app.basePath);
      if (!this.cache.has(url)) {
        const audio = new Audio(url);
        audio.preload = "auto";
        this.cache.set(url, audio);
      }
      return this.cache.get(url);
    }

    stop() {
      if (this.current) {
        this.current.pause();
        this.current.currentTime = 0;
      }
      this.current = null;
    }

    play(path, options = {}) {
      if (!this.enabled || !path) return;
      const audio = this.audio(path);
      this.stop();
      audio.loop = Boolean(options.loop);
      audio.currentTime = 0;
      this.current = audio;
      audio.play().catch(() => {});
    }

    playOverlay(path) {
      if (!this.enabled || !path) return;
      const audio = this.audio(path);
      audio.loop = false;
      audio.currentTime = 0;
      audio.play().catch(() => {});
    }
  }

  class StudentSelectorApp {
    constructor(container, options = {}) {
      this.container = container;
      this.options = options;
      this.basePath = options.basePath || SCRIPT_BASE || window.location.href;
      this.classes = {};
      this.messages = {};
      this.state = loadStoredState();
      this.stage = { mode: "idle" };
      this.modal = null;
      this.timers = new Set();
      this.sound = new SoundManager(this);
      this.boundKeydown = (event) => this.handleKeydown(event);
      document.addEventListener("keydown", this.boundKeydown);
      this.loadData();
      this.render();
    }

    destroy() {
      this.clearTimers();
      this.sound.stop();
      document.removeEventListener("keydown", this.boundKeydown);
      this.container.innerHTML = "";
    }

    clearTimers() {
      for (const id of this.timers) {
        window.clearTimeout(id);
        window.clearInterval(id);
      }
      this.timers.clear();
    }

    setTimeout(fn, delay) {
      const id = window.setTimeout(() => {
        this.timers.delete(id);
        fn();
      }, delay);
      this.timers.add(id);
      return id;
    }

    setInterval(fn, delay) {
      const id = window.setInterval(fn, delay);
      this.timers.add(id);
      return id;
    }

    save() {
      const saved = {
        selectedClass: this.state.selectedClass,
        timerSeconds: this.state.timerSeconds,
        soundEnabled: this.state.soundEnabled,
        slotEffectEnabled: this.state.slotEffectEnabled,
        selectedStudentsByClass: this.state.selectedStudentsByClass,
        studentGradesByClass: this.state.studentGradesByClass,
        studentUngradedByClass: this.state.studentUngradedByClass,
        absentStudentsByClass: this.state.absentStudentsByClass
      };
      window.sessionStorage.setItem(STORAGE_KEY, JSON.stringify(saved));
    }

    async loadData() {
      try {
        const [studentText, messageText] = await Promise.all([
          fetch(assetUrl("assets/students.csv", this.basePath)).then((response) => {
            if (!response.ok) throw new Error("Missing students.csv");
            return response.text();
          }),
          fetch(assetUrl("assets/messages.csv", this.basePath)).then((response) => {
            if (!response.ok) throw new Error("Missing messages.csv");
            return response.text();
          })
        ]);
        this.classes = this.parseStudents(studentText);
        this.messages = this.parseMessages(messageText);
        if (this.state.selectedClass && !this.classes[this.state.selectedClass]) {
          this.state.selectedClass = "";
          this.save();
        }
        this.render();
      } catch (error) {
        this.stage = {
          mode: "error",
          message: error.message || "Unable to load selector data."
        };
        this.render();
      }
    }

    parseStudents(text) {
      const rows = parseCsv(text);
      const classes = {};
      rows.forEach((row, index) => {
        if (row.length < 2) return;
        const first = row[0].trim();
        const second = row[1].trim();
        if (index === 0) {
          const header = row.join(",").toLowerCase();
          if (header.includes("class") && (header.includes("student") || header.includes("name"))) return;
        }
        if (!first || !second) return;
        classes[first] ||= [];
        uniquePush(classes[first], second);
      });
      return classes;
    }

    parseMessages(text) {
      const rows = parseCsv(text);
      const messages = {};
      if (!rows.length) return messages;
      const headers = rows[0].map((value) => value.trim().toLowerCase());
      const ratingIndex = headers.indexOf("rating");
      const messageIndex = headers.indexOf("message");
      rows.slice(1).forEach((row) => {
        const rating = (row[ratingIndex >= 0 ? ratingIndex : 0] || "").trim();
        const message = (row[messageIndex >= 0 ? messageIndex : 1] || "").trim();
        if (!rating || !message) return;
        messages[rating] ||= [];
        messages[rating].push(message);
      });
      return messages;
    }

    classState(className = this.state.selectedClass) {
      this.state.selectedStudentsByClass[className] ||= [];
      this.state.studentGradesByClass[className] ||= {};
      this.state.studentUngradedByClass[className] ||= [];
      this.state.absentStudentsByClass[className] ||= [];
      return {
        selected: this.state.selectedStudentsByClass[className],
        grades: this.state.studentGradesByClass[className],
        ungraded: this.state.studentUngradedByClass[className],
        absent: this.state.absentStudentsByClass[className]
      };
    }

    chosenStudents(className = this.state.selectedClass) {
      if (!className) return [];
      const data = this.classState(className);
      const chosen = [];
      data.selected.forEach((student) => uniquePush(chosen, student));
      Object.keys(data.grades).forEach((student) => uniquePush(chosen, student));
      data.ungraded.forEach((student) => uniquePush(chosen, student));
      return chosen;
    }

    effectiveRoster(className = this.state.selectedClass) {
      const roster = this.classes[className] || [];
      const data = this.classState(className || "none");
      const blocked = new Set([...this.chosenStudents(className), ...data.absent]);
      return roster.filter((student) => !blocked.has(student));
    }

    metrics(className = this.state.selectedClass) {
      const roster = this.classes[className] || [];
      const data = className ? this.classState(className) : { grades: {}, ungraded: [], absent: [] };
      return {
        rosterTotal: roster.length,
        remaining: className ? this.effectiveRoster(className).length : 0,
        graded: Object.keys(data.grades).length,
        ungraded: data.ungraded.length,
        absent: data.absent.length
      };
    }

    setClass(className) {
      this.state.selectedClass = className;
      if (className) this.classState(className);
      this.stage = { mode: "idle" };
      this.save();
      this.render();
    }

    setTimer(seconds) {
      this.state.timerSeconds = seconds;
      this.save();
      this.render();
    }

    toggle(key) {
      this.state[key] = !this.state[key];
      if (key === "soundEnabled" && !this.state[key]) this.sound.stop();
      this.save();
      this.render();
    }

    playIntro() {
      this.sound.play(AUDIO.intro);
    }

    playClosing() {
      this.sound.play(AUDIO.closing);
    }

    startSelection() {
      const className = this.state.selectedClass;
      if (!className || !this.classes[className]) {
        this.showMessage("Select a class", "Please select a valid class before starting.");
        return;
      }
      const roster = this.effectiveRoster(className);
      if (!roster.length) {
        this.showSummary();
        return;
      }

      const finalStudent = choice(roster);
      const duration = Math.max(1, Number(this.state.timerSeconds) || 5);
      const slotEffectEnabled = Boolean(this.state.slotEffectEnabled);
      this.clearTimers();
      this.stage = {
        mode: "selecting",
        className,
        finalStudent,
        pool: roster,
        names: slotEffectEnabled ? [choice(roster), choice(roster), choice(roster)] : ["", "Get ready", ""],
        startedAt: Date.now(),
        duration,
        progress: 0
      };
      this.render();

      if (slotEffectEnabled) {
        this.sound.play(this.slotSound(duration), { loop: true });
        this.setInterval(() => this.tickSelection(), 90);
      }

      this.setTimeout(() => this.finalizeSelection(), duration * 1000);
      this.setTimeout(() => this.sound.playOverlay(AUDIO.timeup), Math.max(0, duration * 1000 - 200));
    }

    slotSound(duration) {
      if (duration <= 4.99) return AUDIO.slotShort;
      if (duration <= 19.99) return AUDIO.slotMedium;
      return AUDIO.slotLong;
    }

    tickSelection() {
      if (this.stage.mode !== "selecting") return;
      const elapsed = Date.now() - this.stage.startedAt;
      this.stage.progress = Math.min(100, (elapsed / (this.stage.duration * 1000)) * 100);
      this.stage.names = [choice(this.stage.pool), choice(this.stage.pool), choice(this.stage.pool)];
      this.updateStageOnly();
    }

    finalizeSelection() {
      if (this.stage.mode !== "selecting") return;
      this.sound.stop();
      const className = this.stage.className;
      const student = this.stage.finalStudent;
      uniquePush(this.classState(className).selected, student);
      this.stage = {
        mode: "selected",
        className,
        finalStudent: student,
        names: ["", student, ""],
        progress: 100
      };
      this.save();
      this.render();
    }

    applyRating(rating) {
      if (this.stage.mode !== "selected") return;
      const className = this.stage.className;
      const student = this.stage.finalStudent;
      const data = this.classState(className);
      removeFrom(data.absent, student);
      removeFrom(data.ungraded, student);
      data.grades[student] = rating;
      uniquePush(data.selected, student);
      this.save();
      this.sound.play(AUDIO.ratings[rating]);
      const messages = this.messages[rating] || [];
      this.showFeedback("Feedback", messages.length ? choice(messages) : "Noted.");
    }

    markNoGrade() {
      if (this.stage.mode !== "selected") return;
      const className = this.stage.className;
      const student = this.stage.finalStudent;
      const data = this.classState(className);
      delete data.grades[student];
      removeFrom(data.absent, student);
      uniquePush(data.ungraded, student);
      uniquePush(data.selected, student);
      this.save();
      this.showFeedback("No Grade", `No grade recorded for ${student} this round.`);
    }

    markAbsent() {
      if (this.stage.mode !== "selected") return;
      const className = this.stage.className;
      const student = this.stage.finalStudent;
      const data = this.classState(className);
      delete data.grades[student];
      removeFrom(data.ungraded, student);
      uniquePush(data.absent, student);
      uniquePush(data.selected, student);
      this.save();
      this.showFeedback("Absent", `${student} was marked absent and removed from today's list.`);
    }

    nextStudent() {
      this.modal = null;
      this.startSelection();
    }

    returnToDock() {
      this.modal = null;
      this.stage = { mode: "idle" };
      this.clearTimers();
      this.sound.stop();
      this.render();
    }

    startAttendance() {
      const className = this.state.selectedClass;
      if (!className || !this.classes[className]) {
        this.showMessage("Select a class", "Please select a valid class before taking attendance.");
        return;
      }
      const roster = this.classes[className] || [];
      if (!roster.length) {
        this.showMessage("No students", "This class has no students in the roster.");
        return;
      }
      this.modal = {
        type: "attendance",
        className,
        roster,
        index: 0,
        absent: []
      };
      this.render();
    }

    markAttendance(isPresent) {
      if (!this.modal || this.modal.type !== "attendance") return;
      const current = this.modal.roster[this.modal.index];
      if (!current) return;
      if (!isPresent) this.modal.absent.push(current);
      this.modal.index += 1;
      if (this.modal.index >= this.modal.roster.length) {
        this.finishAttendance();
      } else {
        this.renderModal();
      }
    }

    finishAttendance() {
      const className = this.modal.className;
      const absent = Array.from(new Set(this.modal.absent));
      this.classState(className).absent = absent;
      this.state.absentStudentsByClass[className] = absent;
      this.save();
      this.modal = {
        type: "attendance-result",
        className,
        absent
      };
      this.render();
    }

    copyAbsentList() {
      if (!this.modal || this.modal.type !== "attendance-result") return;
      const text = this.modal.absent.length ? this.modal.absent.join("\n") : "(none)";
      navigator.clipboard?.writeText(text).then(
        () => this.showMessage("Copied", "Absent list copied to clipboard."),
        () => this.showMessage("Copy failed", "Unable to copy the absent list in this browser.")
      );
    }

    showSummary() {
      this.modal = { type: "summary", className: this.state.selectedClass };
      this.render();
    }

    showFeedback(title, message) {
      this.modal = { type: "feedback", title, message };
      this.render();
    }

    showMessage(title, message) {
      this.modal = { type: "message", title, message };
      this.render();
    }

    closeModal() {
      this.modal = null;
      this.render();
    }

    resetSession() {
      const className = this.state.selectedClass;
      if (!className) return;
      this.state.selectedStudentsByClass[className] = [];
      this.state.studentGradesByClass[className] = {};
      this.state.studentUngradedByClass[className] = [];
      this.state.absentStudentsByClass[className] = [];
      this.stage = { mode: "idle" };
      this.modal = null;
      this.clearTimers();
      this.sound.stop();
      this.save();
      this.render();
    }

    handleKeydown(event) {
      if (!this.modal || this.modal.type !== "attendance") {
        if (event.key === "Escape" && this.modal) this.closeModal();
        return;
      }
      const presentKeys = ["Enter", " ", "ArrowRight", "p", "P"];
      const absentKeys = ["Backspace", "ArrowLeft", "a", "A"];
      if (presentKeys.includes(event.key)) {
        event.preventDefault();
        this.markAttendance(true);
      }
      if (absentKeys.includes(event.key)) {
        event.preventDefault();
        this.markAttendance(false);
      }
      if (event.key === "Escape") this.closeModal();
    }

    updateStageOnly() {
      const root = this.container.querySelector("[data-stage]");
      if (!root) return;
      const current = root.querySelector("[data-current-name]");
      const prev = root.querySelector("[data-prev-name]");
      const next = root.querySelector("[data-next-name]");
      const progress = root.querySelector("[data-progress]");
      if (prev) prev.textContent = this.stage.names[0] || "";
      if (current) current.textContent = this.stage.names[1] || "";
      if (next) next.textContent = this.stage.names[2] || "";
      if (progress) progress.style.width = `${Math.round(this.stage.progress || 0)}%`;
    }

    render() {
      this.container.classList.add("selector-root");
      this.container.innerHTML = `
        <div class="selector-shell">
          ${this.renderDock()}
          ${this.renderStage()}
        </div>
        <div class="selector-modal" data-modal ${this.modal ? "" : "hidden"}>
          ${this.modal ? this.renderModalContent() : ""}
        </div>
      `;
      this.bind();
    }

    renderDock() {
      const classNames = Object.keys(this.classes).sort((a, b) => a.localeCompare(b));
      const metrics = this.metrics();
      const selected = this.state.selectedClass;
      return `
        <section class="selector-dock" aria-label="Selector controls">
          <div class="selector-titlebar">
            <div class="selector-brand">
              <img src="${assetUrl("assets/icon.png", this.basePath)}" alt="">
              <div>
                <h1>Random Student Selector</h1>
                <p>${selected ? this.metricSummary(metrics) : "Select a class to start"}</p>
              </div>
            </div>
            ${this.options.onClose ? `<button class="selector-close" type="button" data-action="close-overlay" aria-label="Close selector">x</button>` : `<a class="selector-about" href="${assetUrl("about.html", this.basePath)}">About</a>`}
          </div>

          <label class="selector-section">
            <span class="selector-section-label">Class</span>
            <select class="selector-class-select" data-action="class">
              <option value="">Select a Class</option>
              ${classNames.map((name) => `<option value="${escapeHtml(name)}" ${name === selected ? "selected" : ""}>${escapeHtml(name)}</option>`).join("")}
            </select>
          </label>

          <div class="selector-metrics" aria-label="Class metrics">
            ${this.metric("Remaining", metrics.remaining)}
            ${this.metric("Graded", metrics.graded)}
            ${this.metric("No Grade", metrics.ungraded)}
            ${this.metric("Absent", metrics.absent)}
          </div>

          <div class="selector-section">
            <span class="selector-section-label">Timer</span>
            <div class="selector-segmented">
              ${[[5, "5 sec"], [30, "30 sec"], [60, "1 min"], [120, "2 min"]].map(([seconds, label]) => `
                <button class="selector-time-button ${this.state.timerSeconds === seconds ? "is-active" : ""}" type="button" data-action="timer" data-seconds="${seconds}">${label}</button>
              `).join("")}
            </div>
          </div>

          <div class="selector-toggle-row">
            <button class="selector-toggle ${this.state.soundEnabled ? "is-active" : ""}" type="button" data-action="toggle" data-key="soundEnabled">Sound</button>
            <button class="selector-toggle ${this.state.slotEffectEnabled ? "is-active" : ""}" type="button" data-action="toggle" data-key="slotEffectEnabled">Slot Effect</button>
          </div>

          <div class="selector-actions">
            <button class="selector-button" type="button" data-action="attendance">Attendance</button>
            <button class="selector-button" type="button" data-action="summary">View Summary</button>
            <button class="selector-button" type="button" data-action="intro">Play Intro</button>
            <button class="selector-button" type="button" data-action="closing">Play Closing</button>
          </div>

          <button class="selector-button selector-primary" type="button" data-action="start">START SELECTION</button>
          <button class="selector-button" type="button" data-action="reset" ${selected ? "" : "disabled"}>Reset Current Class</button>
        </section>
      `;
    }

    renderStage() {
      if (this.stage.mode === "error") {
        return `
          <section class="selector-stage" data-stage>
            <div class="selector-stage-header"><h2>Data unavailable</h2></div>
            <div class="selector-reel"><p class="selector-empty">${escapeHtml(this.stage.message)}</p></div>
            <div></div>
          </section>
        `;
      }

      const metrics = this.metrics();
      const className = this.state.selectedClass || "No class selected";
      const names = this.stage.names || ["", "Ready", ""];
      const current =
        this.stage.mode === "selected"
          ? this.stage.finalStudent
          : this.state.selectedClass
            ? "Ready"
            : "Select a class";
      const currentName = this.stage.mode === "idle" ? current : names[1];
      return `
        <section class="selector-stage" data-stage aria-label="Selection stage">
          <header class="selector-stage-header">
            <div>
              <h2>${escapeHtml(className)}</h2>
              <p class="selector-meta">${this.metricSummary(metrics)}</p>
            </div>
            <div class="selector-pill-row">
              <span class="selector-pill">${formatSeconds(this.state.timerSeconds)}</span>
              <span class="selector-pill">${this.state.slotEffectEnabled ? "Slot effect" : "No slot effect"}</span>
              <span class="selector-pill">${this.state.soundEnabled ? "Sound on" : "Muted"}</span>
            </div>
          </header>

          <div class="selector-reel">
            <div class="selector-reel-window">
              <div class="selector-name-stack">
                <div class="selector-name" data-prev-name>${escapeHtml(names[0] || "")}</div>
                <div class="selector-name is-current ${this.stage.mode === "selected" ? "is-selected" : ""}" data-current-name>${escapeHtml(currentName || "")}</div>
                <div class="selector-name" data-next-name>${escapeHtml(names[2] || "")}</div>
              </div>
            </div>
            <div class="selector-progress" aria-hidden="true"><span data-progress style="width: ${Math.round(this.stage.progress || 0)}%"></span></div>
          </div>

          <footer class="selector-stage-footer">
            ${this.stage.mode === "selected" ? this.renderOutcomeControls() : `<p class="selector-help">Ask the question, protect thinking time, then start selection.</p>`}
          </footer>
        </section>
      `;
    }

    renderOutcomeControls() {
      return `
        <div class="selector-outcomes" aria-label="Outcome controls">
          ${RATINGS.map(([rating, label]) => `
            <button class="selector-outcome" type="button" data-action="rate" data-rating="${rating}">
              ${rating}<span>${label}</span>
            </button>
          `).join("")}
          <button class="selector-outcome" type="button" data-action="no-grade" data-action-kind="no-grade">No Grade<span>Skip grading</span></button>
          <button class="selector-outcome" type="button" data-action="absent" data-action-kind="absent">Absent<span>Remove for today</span></button>
        </div>
      `;
    }

    metric(label, value) {
      return `<div class="selector-metric"><strong>${Number(value) || 0}</strong><span>${label}</span></div>`;
    }

    metricSummary(metrics) {
      const parts = [`${metrics.remaining} left`, `${metrics.graded} graded`];
      if (metrics.ungraded) parts.push(`${metrics.ungraded} no grade`);
      if (metrics.absent) parts.push(`${metrics.absent} absent`);
      return parts.join(" | ");
    }

    renderModalContent() {
      if (!this.modal) return "";
      if (this.modal.type === "attendance") return this.renderAttendance();
      if (this.modal.type === "attendance-result") return this.renderAttendanceResult();
      if (this.modal.type === "summary") return this.renderSummary();
      if (this.modal.type === "feedback") return this.renderFeedback();
      return this.renderMessage();
    }

    renderModal() {
      const modal = this.container.querySelector("[data-modal]");
      if (!modal) return;
      modal.hidden = !this.modal;
      modal.innerHTML = this.modal ? this.renderModalContent() : "";
      this.bind();
    }

    renderPanelHeader(title, meta = "") {
      return `
        <div class="selector-modal__header">
          <div>
            <h2>${escapeHtml(title)}</h2>
            ${meta ? `<p class="selector-meta">${escapeHtml(meta)}</p>` : ""}
          </div>
          <button class="selector-close" type="button" data-action="close-modal" aria-label="Close">x</button>
        </div>
      `;
    }

    renderAttendance() {
      const current = this.modal.roster[this.modal.index] || "";
      return `
        <section class="selector-modal__panel" role="dialog" aria-modal="true" aria-label="Roll call">
          ${this.renderPanelHeader(this.modal.className, `Student ${this.modal.index + 1} of ${this.modal.roster.length}`)}
          <div class="selector-rollcall-name">${escapeHtml(current)}</div>
          <p class="selector-help">Present: Enter, Space, P, Right Arrow. Absent: A, Backspace, Left Arrow.</p>
          <div class="selector-actions">
            <button class="selector-button selector-primary" type="button" data-action="attendance-present">Present</button>
            <button class="selector-button" type="button" data-action="attendance-absent">Absent</button>
          </div>
        </section>
      `;
    }

    renderAttendanceResult() {
      const absentText = this.modal.absent.length ? this.modal.absent.join("\n") : "(none)";
      return `
        <section class="selector-modal__panel" role="dialog" aria-modal="true" aria-label="Absent students">
          ${this.renderPanelHeader(`${this.modal.className} - Absent Students`, `${this.modal.absent.length} absent`)}
          <textarea rows="10" readonly>${escapeHtml(absentText)}</textarea>
          <div class="selector-actions">
            <button class="selector-button selector-primary" type="button" data-action="copy-absent">Copy Absent List</button>
            <button class="selector-button" type="button" data-action="close-modal">Close</button>
          </div>
        </section>
      `;
    }

    renderFeedback() {
      return `
        <section class="selector-modal__panel" role="dialog" aria-modal="true" aria-label="${escapeHtml(this.modal.title)}">
          ${this.renderPanelHeader(this.modal.title, this.state.selectedClass)}
          <div class="selector-rollcall-name">${escapeHtml(this.modal.message)}</div>
          <div class="selector-actions">
            <button class="selector-button selector-primary" type="button" data-action="next-student">Next Student</button>
            <button class="selector-button" type="button" data-action="return-dock">Return To Dock</button>
          </div>
        </section>
      `;
    }

    renderMessage() {
      return `
        <section class="selector-modal__panel" role="dialog" aria-modal="true" aria-label="${escapeHtml(this.modal.title)}">
          ${this.renderPanelHeader(this.modal.title)}
          <p class="selector-help">${escapeHtml(this.modal.message)}</p>
          <button class="selector-button selector-primary" type="button" data-action="close-modal">Close</button>
        </section>
      `;
    }

    renderSummary() {
      const className = this.modal.className;
      const data = className ? this.classState(className) : { grades: {}, ungraded: [], absent: [] };
      const metrics = className ? this.metrics(className) : this.metrics();
      let rowNumber = 1;
      const row = (student, tag) => `
        <div class="selector-summary-row">
          <span>${String(rowNumber++).padStart(2, "0")}</span>
          <strong>${escapeHtml(student)}</strong>
          <span class="selector-tag">${escapeHtml(tag)}</span>
        </div>
      `;
      const sections = [];
      const gradeEntries = Object.entries(data.grades);
      if (gradeEntries.length) {
        sections.push(`<div class="selector-summary-section">Grades</div>`);
        gradeEntries.forEach(([student, rating]) => sections.push(row(student, rating)));
      }
      if (data.ungraded.length) {
        sections.push(`<div class="selector-summary-section">No Grade</div>`);
        data.ungraded.forEach((student) => sections.push(row(student, "No Grade")));
      }
      if (data.absent.length) {
        sections.push(`<div class="selector-summary-section">Absent</div>`);
        data.absent.forEach((student) => sections.push(row(student, "Absent")));
      }
      return `
        <section class="selector-modal__panel" role="dialog" aria-modal="true" aria-label="Session summary">
          ${this.renderPanelHeader(className || "Session Summary", this.metricSummary(metrics))}
          <div class="selector-metrics">
            ${this.metric("Remaining", metrics.remaining)}
            ${this.metric("Graded", metrics.graded)}
            ${this.metric("No Grade", metrics.ungraded)}
            ${this.metric("Absent", metrics.absent)}
          </div>
          <div class="selector-summary-list">
            ${sections.length ? sections.join("") : `<p class="selector-empty">No session records yet.</p>`}
          </div>
          <div class="selector-actions">
            <button class="selector-button selector-primary" type="button" data-action="start">Resume Session</button>
            <button class="selector-button" type="button" data-action="close-modal">Close</button>
          </div>
        </section>
      `;
    }

    bind() {
      this.container.querySelectorAll("[data-action]").forEach((element) => {
        element.addEventListener("click", (event) => {
          const action = event.currentTarget.dataset.action;
          if (action === "timer") this.setTimer(Number(event.currentTarget.dataset.seconds));
          if (action === "toggle") this.toggle(event.currentTarget.dataset.key);
          if (action === "attendance") this.startAttendance();
          if (action === "summary") this.showSummary();
          if (action === "intro") this.playIntro();
          if (action === "closing") this.playClosing();
          if (action === "start") this.startSelection();
          if (action === "reset") this.resetSession();
          if (action === "rate") this.applyRating(event.currentTarget.dataset.rating);
          if (action === "no-grade") this.markNoGrade();
          if (action === "absent") this.markAbsent();
          if (action === "attendance-present") this.markAttendance(true);
          if (action === "attendance-absent") this.markAttendance(false);
          if (action === "copy-absent") this.copyAbsentList();
          if (action === "next-student") this.nextStudent();
          if (action === "return-dock") this.returnToDock();
          if (action === "close-modal") this.closeModal();
          if (action === "close-overlay") this.options.onClose?.();
        });
      });

      const select = this.container.querySelector("[data-action='class']");
      if (select) {
        select.addEventListener("change", (event) => this.setClass(event.currentTarget.value));
      }
    }
  }

  function ensureStylesheet(basePath) {
    const href = assetUrl("selector.css", basePath || SCRIPT_BASE);
    const stylesheets =
      typeof document.querySelectorAll === "function"
        ? Array.from(document.querySelectorAll("link[rel='stylesheet']"))
        : [];
    const alreadyLoaded = stylesheets.some((link) => {
      try {
        return new URL(link.href, document.baseURI || window.location.href).href === href;
      } catch (_error) {
        return false;
      }
    });
    if (alreadyLoaded) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    document.head.appendChild(link);
  }

  function mount(container, options = {}) {
    if (!container) throw new Error("StudentSelector.mount requires a container element.");
    ensureStylesheet(options.basePath);
    return new StudentSelectorApp(container, options);
  }

  function open(options = {}) {
    ensureStylesheet(options.basePath);
    const host = document.createElement("div");
    host.className = "selector-overlay-host";
    document.body.appendChild(host);
    let app;
    const close = () => {
      app?.destroy();
      host.remove();
      options.onClose?.();
    };
    app = mount(host, { ...options, onClose: close });
    return { app, close, element: host };
  }

  window.StudentSelector = {
    mount,
    open
  };
})();
