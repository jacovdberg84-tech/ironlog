// IRONLOG/web/app/june-avatar.js — lightweight, visual-only presence for June.
// This module deliberately knows nothing about GPT-Live, WebRTC, or Ironlog
// tools. june.js tells it which single visual state to show and, when present,
// gives it the existing remote audio stream for a subtle speaking response.

const JUNE_AVATAR_STATES = new Set(["idle", "listening", "thinking", "speaking"]);
let juneAvatarState = "idle";
let juneAvatarAudioContext = null;
let juneAvatarAnalyser = null;
let juneAvatarSource = null;
let juneAvatarAudioStream = null;
let juneAvatarData = null;
let juneAvatarFrame = null;
let juneAvatarLevel = 0;
let juneAvatarWarned = false;

function juneAvatarRoot() {
  return document.getElementById("juneAvatar");
}

function juneAvatarReducedMotion() {
  return Boolean(window.matchMedia?.("(prefers-reduced-motion: reduce)")?.matches);
}

function juneAvatarLabel(state) {
  return {
    idle: "June is ready",
    listening: "June is listening",
    thinking: "June is thinking",
    speaking: "June is speaking",
  }[state] || "June is ready";
}

function juneAvatarSetLevel(level = 0) {
  const root = juneAvatarRoot();
  if (!root) return;
  const next = Math.max(0, Math.min(1, Number(level) || 0));
  root.style.setProperty("--june-avatar-level", next.toFixed(3));
}

function juneAvatarStopFrame() {
  if (juneAvatarFrame) window.cancelAnimationFrame(juneAvatarFrame);
  juneAvatarFrame = null;
}

function juneAvatarRunFrame() {
  juneAvatarStopFrame();
  if (juneAvatarState !== "speaking" || !juneAvatarAnalyser || juneAvatarReducedMotion()) return;
  const tick = () => {
    if (juneAvatarState !== "speaking" || !juneAvatarAnalyser) return;
    try {
      juneAvatarAnalyser.getByteTimeDomainData(juneAvatarData);
      let energy = 0;
      for (const sample of juneAvatarData) {
        const value = (sample - 128) / 128;
        energy += value * value;
      }
      const target = Math.min(1, Math.sqrt(energy / juneAvatarData.length) * 5.25);
      juneAvatarLevel += (target - juneAvatarLevel) * 0.16;
      juneAvatarSetLevel(juneAvatarLevel);
    } catch {
      // The audio element still plays normally if analysis becomes unavailable.
      juneAvatarStopFrame();
      return;
    }
    juneAvatarFrame = window.requestAnimationFrame(tick);
  };
  juneAvatarFrame = window.requestAnimationFrame(tick);
}

function juneAvatarDiagnostic(error) {
  if (juneAvatarWarned) return;
  juneAvatarWarned = true;
  console.info("June portrait audio response is unavailable; using the subtle visual fallback.", error);
}

function setJuneAvatarVisualState(nextState = "idle") {
  const state = JUNE_AVATAR_STATES.has(nextState) ? nextState : "idle";
  juneAvatarState = state;
  const root = juneAvatarRoot();
  if (!root) return;
  root.dataset.state = state;
  const status = document.getElementById("juneAvatarStatus");
  if (status) status.textContent = juneAvatarLabel(state);
  if (state === "speaking") juneAvatarRunFrame();
  else {
    juneAvatarStopFrame();
    juneAvatarLevel = 0;
    juneAvatarSetLevel(0);
  }
}

function startJuneAvatarAudioAnalysis(stream) {
  if (!stream || juneAvatarReducedMotion()) return;
  if (juneAvatarAudioStream === stream && juneAvatarAnalyser) {
    juneAvatarRunFrame();
    return;
  }
  stopJuneAvatarAudioAnalysis();
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) return;
  try {
    juneAvatarAudioContext = new AudioContextClass();
    juneAvatarSource = juneAvatarAudioContext.createMediaStreamSource(stream);
    juneAvatarAnalyser = juneAvatarAudioContext.createAnalyser();
    juneAvatarAnalyser.fftSize = 256;
    juneAvatarAnalyser.smoothingTimeConstant = 0.78;
    juneAvatarData = new Uint8Array(juneAvatarAnalyser.fftSize);
    juneAvatarSource.connect(juneAvatarAnalyser);
    juneAvatarAudioStream = stream;
    Promise.resolve(juneAvatarAudioContext.resume()).catch(() => {});
    juneAvatarRunFrame();
  } catch (error) {
    juneAvatarDiagnostic(error);
    stopJuneAvatarAudioAnalysis();
  }
}

function stopJuneAvatarAudioAnalysis() {
  juneAvatarStopFrame();
  juneAvatarLevel = 0;
  juneAvatarSetLevel(0);
  try { juneAvatarSource?.disconnect(); } catch {}
  try { juneAvatarAnalyser?.disconnect(); } catch {}
  if (juneAvatarAudioContext && juneAvatarAudioContext.state !== "closed") {
    juneAvatarAudioContext.close().catch(() => {});
  }
  juneAvatarAudioContext = null;
  juneAvatarAnalyser = null;
  juneAvatarSource = null;
  juneAvatarAudioStream = null;
  juneAvatarData = null;
}

