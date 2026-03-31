const state = {
  hemisphere: "N",
  event: "sunrise",
  lastResult: null,
  guessLocation: null,
  magneticCorrectionEnabled: false,
  lastDeclination: null,
  lastCorrectedAzimuth: null,
  wmm: null,
  wmmReady: false,
};

const MAP_VIEWBOX = { width: 1000, height: 500 };
const MAP_LABEL = {
  height: 36,
  paddingX: 12,
  charWidth: 7.2,
  minWidth: 96,
  offsetX: 16,
  offsetY: 38,
  margin: 8,
};
const WMM_MAX_N = 12;
const WMM_A = 6378.137;
const WMM_B = 6356.7523142;
const WMM_RE = 6371.2;

function $(id) {
  return document.getElementById(id);
}

function setSegmented(controlId, value) {
  const control = $(controlId);
  if (!control) return;
  const buttons = control.querySelectorAll("button");
  buttons.forEach((btn) => btn.classList.toggle("active", btn.dataset.value === value));
}

function bindSegmented(controlId, stateKey) {
  const control = $(controlId);
  if (!control) return;

  control.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-value]");
    if (!btn) return;
    state[stateKey] = btn.dataset.value;
    setSegmented(controlId, state[stateKey]);
  });
}

function showError(message) {
  const box = $("errorBox");
  if (!box) return;
  box.textContent = message;
  box.classList.remove("hidden");
}

function clearError() {
  const box = $("errorBox");
  if (!box) return;
  box.textContent = "";
  box.classList.add("hidden");
}

function fmtFixed(value, digits = 6) {
  return Number(value).toFixed(digits);
}

function dmsParts(value) {
  const abs = Math.abs(value);
  const wholeDegrees = Math.floor(abs);
  const minutesFloat = (abs - wholeDegrees) * 60;
  let minutes = Math.round(minutesFloat);
  let degrees = wholeDegrees;

  if (minutes === 60) {
    degrees += 1;
    minutes = 0;
  }

  return { d: degrees, m: minutes };
}

function fmtLatDM(lat) {
  const { d, m } = dmsParts(lat);
  return `${d}° ${m}′ ${lat >= 0 ? "N" : "S"}`;
}

function fmtLonDM(lon) {
  const { d, m } = dmsParts(lon);
  return `${d}° ${m}′ ${lon >= 0 ? "E" : "W"}`;
}

function buildCoordinateLabel(lat, lon) {
  return `${fmtLatDM(lat)} · ${fmtLonDM(lon)}`;
}

function row(label, value) {
  return `<div class="row"><div class="row-label">${label}</div><div class="row-value">${value}</div></div>`;
}

function clamp(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

function deg2rad(deg) {
  return (deg * Math.PI) / 180;
}

function rad2deg(rad) {
  return (rad * 180) / Math.PI;
}

function normalize360(angleDeg) {
  let a = angleDeg % 360;
  if (a < 0) a += 360;
  return a;
}

function normalize180(angleDeg) {
  return ((angleDeg + 180) % 360 + 360) % 360 - 180;
}

function horizonDipDeg(observerAltM) {
  return 0.0293 * Math.sqrt(observerAltM);
}

function julianDay(date) {
  return date.getTime() / 86400000 + 2440587.5;
}

function decimalYearFromDate(date) {
  const start = Date.UTC(date.getUTCFullYear(), 0, 1);
  const end = Date.UTC(date.getUTCFullYear() + 1, 0, 1);
  const now = date.getTime();
  return date.getUTCFullYear() + (now - start) / (end - start);
}

function parseInputs() {
  const date = $("dateInput")?.value;
  const time = $("timeInput")?.value;
  const azimuth = parseFloat($("azimuthInput")?.value.trim() || "");
  const altitude = parseFloat($("altitudeInput")?.value.trim() || "");
  const compassCorrection = parseFloat($("compassCorrectionInput")?.value.trim() || "");

  if (!date) throw new Error("Please enter a date.");
  if (!time) throw new Error("Please enter a UTC time.");
  if (!Number.isFinite(azimuth)) throw new Error("Azimuth must be a valid number.");
  if (!Number.isFinite(altitude)) throw new Error("Altitude must be a valid number.");
  if (!Number.isFinite(compassCorrection)) throw new Error("Compass correction must be a valid number.");
  if (altitude < 0) throw new Error("Altitude must be greater than or equal to zero.");

  return {
    date,
    time,
    azimuth,
    altitude,
    compassCorrection,
    hemisphere: state.hemisphere,
    event: state.event,
  };
}

function initBitmapMap() {
  const img = $("mapImage");
  const missing = $("mapMissingGroup");
  const svg = $("mapSvg");

  if (img && missing) {
    img.onload = () => missing.classList.add("hidden");
    img.onerror = () => {
      missing.classList.remove("hidden");
      if ($("mapCaption")) $("mapCaption").textContent = "";
    };
  }

  if (svg) {
    svg.addEventListener("click", onMapClick);
  }
}

function mapPixelToLatLon(svgPointX, svgPointY) {
  const lon = (svgPointX / MAP_VIEWBOX.width) * 360 - 180;
  const lat = 90 - (svgPointY / MAP_VIEWBOX.height) * 180;
  return { lat, lon };
}

function showGuessPreview() {
  if (!$("bigCoord") || !state.guessLocation) return;
  $("bigCoord").textContent = buildCoordinateLabel(state.guessLocation.lat, state.guessLocation.lon);
  if ($("subCoord")) {
    $("subCoord").textContent = "Guessed position selected for magnetic correction";
  }
}

function onMapClick(event) {
  if (!state.magneticCorrectionEnabled) return;

  const svg = $("mapSvg");
  if (!svg) return;

  const rect = svg.getBoundingClientRect();
  const x = ((event.clientX - rect.left) / rect.width) * MAP_VIEWBOX.width;
  const y = ((event.clientY - rect.top) / rect.height) * MAP_VIEWBOX.height;
  const { lat, lon } = mapPixelToLatLon(x, y);

  state.guessLocation = { lat, lon };
  updateGuessMarker();
  updateCorrectionInfoPanel();
  showGuessPreview();
}

function setMarkerPosition(prefix, x, y) {
  $(`${prefix}Halo`)?.setAttribute("cx", x);
  $(`${prefix}Halo`)?.setAttribute("cy", y);
  $(`${prefix}Ring`)?.setAttribute("cx", x);
  $(`${prefix}Ring`)?.setAttribute("cy", y);
  $(`${prefix}Dot`)?.setAttribute("cx", x);
  $(`${prefix}Dot`)?.setAttribute("cy", y);
}

function measureMapLabelWidth(text) {
  return Math.max(MAP_LABEL.minWidth, MAP_LABEL.paddingX * 2 + text.length * MAP_LABEL.charWidth);
}

function positionEstimateLabel(x, y, text) {
  const boxWidth = measureMapLabelWidth(text);
  let labelX = x + MAP_LABEL.offsetX;
  let labelY = y - MAP_LABEL.offsetY;

  if (labelX + boxWidth > MAP_VIEWBOX.width - MAP_LABEL.margin) {
    labelX = x - boxWidth - MAP_LABEL.offsetX;
  }
  if (labelX < MAP_LABEL.margin) {
    labelX = MAP_LABEL.margin;
  }
  if (labelY < MAP_LABEL.margin) {
    labelY = y + 14;
  }
  if (labelY + MAP_LABEL.height > MAP_VIEWBOX.height - MAP_LABEL.margin) {
    labelY = MAP_VIEWBOX.height - MAP_LABEL.height - MAP_LABEL.margin;
  }

  $("mapLabelBox")?.setAttribute("x", labelX);
  $("mapLabelBox")?.setAttribute("y", labelY);
  $("mapLabelBox")?.setAttribute("width", boxWidth);
  $("mapLabelText")?.setAttribute("x", labelX + MAP_LABEL.paddingX);
  $("mapLabelText")?.setAttribute("y", labelY + 23);
  if ($("mapLabelText")) $("mapLabelText").textContent = text;
}

function latLonToMapXY(lat, lon) {
  return {
    x: ((lon + 180) / 360) * MAP_VIEWBOX.width,
    y: ((90 - lat) / 180) * MAP_VIEWBOX.height,
  };
}

function updateGuessMarker() {
  const guessMarker = $("guessMarker");
  if (!guessMarker) return;

  if (!state.magneticCorrectionEnabled || !state.guessLocation) {
    guessMarker.classList.add("hidden");
    return;
  }

  const { x, y } = latLonToMapXY(state.guessLocation.lat, state.guessLocation.lon);
  setMarkerPosition("guessMarker", x, y);
  guessMarker.classList.remove("hidden");
}

function updateMap(lat, lon) {
  const marker = $("mapMarker");
  const label = $("mapMarkerLabel");

  updateGuessMarker();

  if (!marker || !label) return;

  if (lat == null || lon == null || !Number.isFinite(lat) || !Number.isFinite(lon)) {
    marker.classList.add("hidden");
    label.classList.add("hidden");
    if ($("mapCaption")) $("mapCaption").textContent = "";
    return;
  }

  const { x, y } = latLonToMapXY(lat, lon);
  const labelText = buildCoordinateLabel(lat, lon);

  setMarkerPosition("mapMarker", x, y);
  positionEstimateLabel(x, y, labelText);

  marker.classList.remove("hidden");
  label.classList.remove("hidden");
  if ($("mapCaption")) $("mapCaption").textContent = "";
}

function updateWmmStatus(text) {
  if ($("wmmStatusLine")) $("wmmStatusLine").textContent = text;
}

async function tryLoadWmmViaFetch() {
  try {
    updateWmmStatus("WMM2025 status: loading WMM.COF...");
    const response = await fetch("./WMM.COF");
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    state.wmm = parseWmmCof(text);
    state.wmmReady = true;
    updateWmmStatus(`WMM2025 status: loaded ${state.wmm.modelName} (${state.wmm.epoch.toFixed(1)})`);
    updateCorrectionInfoPanel();
  } catch {
    state.wmmReady = false;
    state.wmm = null;
    updateWmmStatus("WMM2025 status: automatic load failed — choose WMM.COF manually");
  }
}

function initWmmFilePicker() {
  const fileButton = $("wmmFileButton");
  const fileInput = $("wmmFileInput");
  const fileName = $("wmmFileName");

  if (fileButton && fileInput) {
    fileButton.addEventListener("click", () => fileInput.click());
  }

  if (!fileInput) return;

  fileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];

    if (!file) {
      if (fileName) {
        fileName.textContent = "Uploaded file: —";
        fileName.classList.add("hidden");
      }
      return;
    }

    try {
      const text = await readFileAsText(file);
      state.wmm = parseWmmCof(text);
      state.wmmReady = true;

      if (fileName) {
        fileName.textContent = `Uploaded file: ${file.name}`;
        fileName.classList.remove("hidden");
      }

      updateWmmStatus(`WMM2025 status: loaded from selected file (${state.wmm.modelName} ${state.wmm.epoch.toFixed(1)})`);
      updateCorrectionInfoPanel();
      clearError();
    } catch (err) {
      state.wmmReady = false;
      state.wmm = null;

      if (fileName) {
        fileName.textContent = "Uploaded file: —";
        fileName.classList.remove("hidden");
      }

      updateWmmStatus("WMM2025 status: selected file could not be parsed");
      showError(err.message || String(err));
    }
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read selected file."));
    reader.readAsText(file);
  });
}

function parseWmmCof(text) {
  const lines = text.replace(/\r/g, "").split("\n").map((s) => s.trim()).filter(Boolean);
  if (lines.length < 10) throw new Error("WMM.COF is too short.");

  const header = lines[0].split(/\s+/);
  if (header.length < 2) throw new Error("Invalid WMM.COF header.");

  const epoch = parseFloat(header[0]);
  const modelName = header[1] || "WMM";
  if (!Number.isFinite(epoch)) throw new Error("Invalid WMM epoch.");

  const g = Array.from({ length: WMM_MAX_N + 1 }, () => Array(WMM_MAX_N + 1).fill(0));
  const h = Array.from({ length: WMM_MAX_N + 1 }, () => Array(WMM_MAX_N + 1).fill(0));
  const gd = Array.from({ length: WMM_MAX_N + 1 }, () => Array(WMM_MAX_N + 1).fill(0));
  const hd = Array.from({ length: WMM_MAX_N + 1 }, () => Array(WMM_MAX_N + 1).fill(0));

  let parsedRows = 0;

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (line.startsWith("999999")) break;

    const parts = line.split(/\s+/);
    if (parts.length < 6) continue;

    const n = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const gnm = parseFloat(parts[2]);
    const hnm = parseFloat(parts[3]);
    const dgnm = parseFloat(parts[4]);
    const dhnm = parseFloat(parts[5]);

    if (
      Number.isInteger(n) &&
      Number.isInteger(m) &&
      n >= 1 && n <= WMM_MAX_N &&
      m >= 0 && m <= n &&
      Number.isFinite(gnm) &&
      Number.isFinite(hnm) &&
      Number.isFinite(dgnm) &&
      Number.isFinite(dhnm)
    ) {
      g[n][m] = gnm;
      h[n][m] = hnm;
      gd[n][m] = dgnm;
      hd[n][m] = dhnm;
      parsedRows += 1;
    }
  }

  if (parsedRows < 90) {
    throw new Error("WMM.COF does not contain the expected coefficient rows.");
  }

  return { epoch, modelName, g, h, gd, hd };
}

function buildAssociatedLegendre(thetaRad, maxN) {
  const sinT = Math.sin(thetaRad);
  const cosT = Math.cos(thetaRad);

  const P = Array.from({ length: maxN + 1 }, () => Array(maxN + 1).fill(0));
  const dP = Array.from({ length: maxN + 1 }, () => Array(maxN + 1).fill(0));

  P[0][0] = 1;
  dP[0][0] = 0;

  for (let n = 1; n <= maxN; n++) {
    for (let m = 0; m <= n; m++) {
      if (n === m) {
        P[n][m] = sinT * P[n - 1][m - 1];
        dP[n][m] = sinT * dP[n - 1][m - 1] + cosT * P[n - 1][m - 1];
      } else if (n === 1 && m === 0) {
        P[n][m] = cosT * P[n - 1][m];
        dP[n][m] = cosT * dP[n - 1][m] - sinT * P[n - 1][m];
      } else if (n > 1 && n !== m) {
        const k = (((n - 1) * (n - 1)) - (m * m)) / ((2 * n - 1) * (2 * n - 3));
        P[n][m] = cosT * P[n - 1][m] - k * P[n - 2][m];
        dP[n][m] = cosT * dP[n - 1][m] - sinT * P[n - 1][m] - k * dP[n - 2][m];
      }
    }
  }

  return { P, dP };
}

function schmidtQuasiNormFactors(maxN) {
  const S = Array.from({ length: maxN + 1 }, () => Array(maxN + 1).fill(0));
  S[0][0] = 1;

  for (let n = 1; n <= maxN; n++) {
    S[n][0] = S[n - 1][0] * ((2 * n - 1) / n);
    for (let m = 1; m <= n; m++) {
      const factor = m === 1 ? 2 : 1;
      S[n][m] = S[n][m - 1] * Math.sqrt((factor * (n - m + 1)) / (n + m));
    }
  }

  return S;
}

const SCHMIDT = schmidtQuasiNormFactors(WMM_MAX_N);

function geodeticToSpherical(latDeg, lonDeg, altitudeKm) {
  const lat = deg2rad(latDeg);
  const lon = deg2rad(lonDeg);

  const a2 = WMM_A * WMM_A;
  const b2 = WMM_B * WMM_B;
  const sinLat = Math.sin(lat);
  const cosLat = Math.cos(lat);

  const e2 = (a2 - b2) / a2;
  const N = WMM_A / Math.sqrt(1 - e2 * sinLat * sinLat);

  const X = (N + altitudeKm) * cosLat;
  const Z = (N * (1 - e2) + altitudeKm) * sinLat;

  const r = Math.sqrt(X * X + Z * Z);
  const theta = Math.atan2(X, Z);
  const latGc = Math.PI / 2 - theta;
  const psi = latGc - lat;

  return { r, theta, lon, psi };
}

function estimateDeclinationWMM2025(latDeg, lonDeg, whenUtc, altitudeMeters = 0) {
  if (!state.wmmReady || !state.wmm) {
    throw new Error("WMM2025 coefficients are not loaded.");
  }

  const altKm = altitudeMeters / 1000;
  const t = decimalYearFromDate(whenUtc) - state.wmm.epoch;

  const { r, theta, lon, psi } = geodeticToSpherical(latDeg, lonDeg, altKm);
  const { P, dP } = buildAssociatedLegendre(theta, WMM_MAX_N);

  let Br = 0;
  let Bt = 0;
  let Bp = 0;

  const sinTheta = Math.sin(theta) || 1e-9;
  const aOverR = WMM_RE / r;

  for (let n = 1; n <= WMM_MAX_N; n++) {
    const ar = Math.pow(aOverR, n + 2);

    for (let m = 0; m <= n; m++) {
      const g = state.wmm.g[n][m] + t * state.wmm.gd[n][m];
      const h = state.wmm.h[n][m] + t * state.wmm.hd[n][m];
      const cosml = Math.cos(m * lon);
      const sinml = Math.sin(m * lon);

      const ghs = g * cosml + h * sinml;
      const ght = g * sinml - h * cosml;

      const Snm = SCHMIDT[n][m];
      const Pnm = P[n][m] * Snm;
      const dPnm = dP[n][m] * Snm;

      Br += ar * (n + 1) * ghs * Pnm;
      Bt -= ar * ghs * dPnm;

      if (m !== 0) {
        Bp += ar * m * ght * Pnm / sinTheta;
      }
    }
  }

  const X = -Bt * Math.cos(psi) - Br * Math.sin(psi);
  const Y = Bp;

  return rad2deg(Math.atan2(Y, X));
}

function computeCorrectedAzimuth(inputs) {
  if (!state.guessLocation) {
    throw new Error("Click the map first to set a guessed location.");
  }
  if (!state.wmmReady) {
    throw new Error("WMM2025 is not loaded. Put WMM.COF in the folder or select it manually.");
  }

  const whenUtc = new Date(`${inputs.date}T${inputs.time}Z`);
  const declination = estimateDeclinationWMM2025(
    state.guessLocation.lat,
    state.guessLocation.lon,
    whenUtc,
    inputs.altitude
  );

  const corrected = normalize360(inputs.azimuth + inputs.compassCorrection + declination);

  state.lastDeclination = declination;
  state.lastCorrectedAzimuth = corrected;

  return corrected;
}

function updateCorrectionInfoPanel() {
  if ($("guessLocationLine")) {
    $("guessLocationLine").textContent = state.guessLocation
      ? `Guessed position: ${fmtLatDM(state.guessLocation.lat)} · ${fmtLonDM(state.guessLocation.lon)}`
      : "Guessed position: not set";
  }

  if ($("declinationLine")) {
    if (state.lastDeclination == null) {
      $("declinationLine").textContent = "WMM2025 declination: not available";
    } else {
      const suffix = state.lastDeclination >= 0 ? "E" : "W";
      $("declinationLine").textContent = `WMM2025 declination: ${Math.abs(state.lastDeclination).toFixed(2)}° ${suffix}`;
    }
  }

  if ($("correctedAzimuthLine")) {
    $("correctedAzimuthLine").textContent =
      state.magneticCorrectionEnabled && state.lastCorrectedAzimuth != null
        ? `Corrected azimuth: ${state.lastCorrectedAzimuth.toFixed(2)}° true`
        : "Corrected azimuth: not applied";
  }
}

function solarPositionApprox(whenUtc, latitudeDeg, longitudeDeg) {
  const jd = julianDay(whenUtc);
  const T = (jd - 2451545.0) / 36525.0;

  let L0 = 280.46646 + T * (36000.76983 + T * 0.0003032);
  L0 = normalize360(L0);

  const M = 357.52911 + T * (35999.05029 - 0.0001537 * T + (T * T) / 24490000);
  const Mrad = deg2rad(M);

  const C =
    Math.sin(Mrad) * (1.914602 - T * (0.004817 + 0.000014 * T)) +
    Math.sin(2 * Mrad) * (0.019993 - 0.000101 * T) +
    Math.sin(3 * Mrad) * 0.000289;

  const trueLong = L0 + C;
  const omega = 125.04 - 1934.136 * T;
  const lambdaApp = trueLong - 0.00569 - 0.00478 * Math.sin(deg2rad(omega));

  const seconds = 21.448 - T * (46.815 + T * (0.00059 - T * 0.001813));
  const eps0 = 23 + (26 + seconds / 60) / 60;
  const eps = eps0 + 0.00256 * Math.cos(deg2rad(omega));

  const lambdaRad = deg2rad(lambdaApp);
  const epsRad = deg2rad(eps);

  let alpha = Math.atan2(Math.cos(epsRad) * Math.sin(lambdaRad), Math.cos(lambdaRad));
  if (alpha < 0) alpha += 2 * Math.PI;

  const delta = Math.asin(Math.sin(epsRad) * Math.sin(lambdaRad));
  const alphaDeg = normalize360(rad2deg(alpha));
  const deltaDeg = rad2deg(delta);

  const theta0 = normalize360(
    280.46061837 +
    360.98564736629 * (jd - 2451545.0) +
    0.000387933 * T * T -
    (T * T * T) / 38710000
  );

  const lst = normalize360(theta0 + longitudeDeg);
  const H = normalize180(lst - alphaDeg);

  const latRad = deg2rad(latitudeDeg);
  const HRad = deg2rad(H);

  const sinAlt =
    Math.sin(latRad) * Math.sin(delta) +
    Math.cos(latRad) * Math.cos(delta) * Math.cos(HRad);

  let altDeg = rad2deg(Math.asin(clamp(sinAlt, -1, 1)));

  if (altDeg > -2 && altDeg < 15) {
    const ref = 1.02 / Math.tan(deg2rad(altDeg + 10.3 / (altDeg + 5.11))) / 60;
    altDeg += ref;
  }

  const azRad = Math.atan2(
    Math.sin(HRad),
    Math.cos(HRad) * Math.sin(latRad) - Math.tan(delta) * Math.cos(latRad)
  );

  const azDeg = normalize360(rad2deg(azRad) + 180);

  return {
    altitudeDeg: altDeg,
    azimuthDeg: azDeg,
    declinationDeg: deltaDeg,
  };
}

function scoreResiduals(rAlt, rAz) {
  return Math.hypot(rAlt * 1.7, rAz);
}

function computeConfidence({ score, declinationDeg, measuredAzimuthDeg, warnings }) {
  let confidence = 100;
  confidence -= Math.min(58, score * 18);

  const dist = Math.min(Math.abs(measuredAzimuthDeg - 90), Math.abs(measuredAzimuthDeg - 270));
  if (dist < 5) confidence -= 35;
  else if (dist < 10) confidence -= 22;
  else if (dist < 15) confidence -= 12;

  const absDec = Math.abs(declinationDeg);
  if (absDec < 5) confidence -= 35;
  else if (absDec < 10) confidence -= 22;
  else if (absDec < 15) confidence -= 10;

  for (const warning of warnings) {
    confidence -= warning.startsWith("WARNING") ? 10 : 4;
  }

  return Math.round(clamp(confidence, 0, 100));
}

function residuals(whenUtc, latitudeDeg, longitudeDeg, altitudeM, measuredAzimuthDeg) {
  const dip = horizonDipDeg(altitudeM);
  const effectiveHorizonDeg = -0.833 + dip;
  const model = solarPositionApprox(whenUtc, latitudeDeg, longitudeDeg);

  return {
    residualAltitudeDeg: model.altitudeDeg - effectiveHorizonDeg,
    residualAzimuthDeg: normalize180(model.azimuthDeg - measuredAzimuthDeg),
    modelAltitudeDeg: model.altitudeDeg,
    modelAzimuthDeg: model.azimuthDeg,
    modelDeclinationDeg: model.declinationDeg,
    effectiveHorizonDeg,
  };
}

function eventMatchesAzimuth(event, azimuthDeg) {
  if (event === "sunrise") return azimuthDeg >= 0 && azimuthDeg <= 180;
  if (event === "sunset") return azimuthDeg >= 180 && azimuthDeg < 360;
  return true;
}

function evaluateCandidate(whenUtc, lat, lon, altitudeM, measuredAzimuthDeg, event) {
  const result = residuals(whenUtc, lat, lon, altitudeM, measuredAzimuthDeg);
  if (!eventMatchesAzimuth(event, result.modelAzimuthDeg)) return null;

  return {
    lat,
    lon,
    score: scoreResiduals(result.residualAltitudeDeg, result.residualAzimuthDeg),
    ...result,
  };
}

function buildLatitudeCandidates(hemisphere) {
  return hemisphere === "N"
    ? Array.from({ length: 71 }, (_, i) => i)
    : Array.from({ length: 71 }, (_, i) => -70 + i);
}

function buildLongitudeCandidates() {
  return Array.from({ length: 181 }, (_, i) => -180 + i * 2);
}

function refineCandidate(best, context) {
  let {
    lat,
    lon,
    residualAltitudeDeg,
    residualAzimuthDeg,
    modelAltitudeDeg,
    modelAzimuthDeg,
    modelDeclinationDeg,
    score,
  } = best;

  let latStep = 1.0;
  let lonStep = 2.0;

  for (let iter = 0; iter < 20; iter++) {
    let improved = false;
    let localBest = {
      lat, lon, residualAltitudeDeg, residualAzimuthDeg,
      modelAltitudeDeg, modelAzimuthDeg, modelDeclinationDeg, score,
    };

    for (const dlat of [-latStep, 0, latStep]) {
      for (const dlon of [-lonStep, 0, lonStep]) {
        if (dlat === 0 && dlon === 0) continue;

        const lat2 = lat + dlat;
        const lon2 = lon + dlon;

        if (context.hemisphere === "N" && !(lat2 >= 0 && lat2 <= 70)) continue;
        if (context.hemisphere === "S" && !(lat2 >= -70 && lat2 <= 0)) continue;
        if (!(lon2 >= -180 && lon2 <= 180)) continue;

        const candidate = evaluateCandidate(
          context.whenUtc,
          lat2,
          lon2,
          context.altitudeM,
          context.measuredAzimuthDeg,
          context.event
        );

        if (!candidate) continue;

        if (candidate.score + 1e-12 < localBest.score) {
          localBest = candidate;
          improved = true;
        }
      }
    }

    lat = localBest.lat;
    lon = localBest.lon;
    residualAltitudeDeg = localBest.residualAltitudeDeg;
    residualAzimuthDeg = localBest.residualAzimuthDeg;
    modelAltitudeDeg = localBest.modelAltitudeDeg;
    modelAzimuthDeg = localBest.modelAzimuthDeg;
    modelDeclinationDeg = localBest.modelDeclinationDeg;
    score = localBest.score;

    latStep *= improved ? 0.72 : 0.5;
    lonStep *= improved ? 0.72 : 0.5;

    if (latStep < 1e-5 && lonStep < 1e-5) break;
  }

  return {
    lat,
    lon,
    residualAltitudeDeg,
    residualAzimuthDeg,
    modelAltitudeDeg,
    modelAzimuthDeg,
    modelDeclinationDeg,
    score,
  };
}

function solve(inputs) {
  const warnings = [];
  const whenUtc = new Date(`${inputs.date}T${inputs.time}Z`);
  const altitudeM = inputs.altitude;

  const measuredAzimuthDeg = state.magneticCorrectionEnabled
    ? computeCorrectedAzimuth(inputs)
    : normalize360(inputs.azimuth);

  const hemisphere = inputs.hemisphere.toUpperCase();
  const event = inputs.event;

  const dip = horizonDipDeg(altitudeM);
  const effectiveHorizonDeg = -0.833 + dip;

  const eastWestDistance = Math.min(
    Math.abs(measuredAzimuthDeg - 90),
    Math.abs(measuredAzimuthDeg - 270)
  );

  if (eastWestDistance < 5) {
    warnings.push("WARNING: Azimuth is very close to due east/west; solution is highly unstable.");
  } else if (eastWestDistance < 15) {
    warnings.push("CAUTION: Azimuth is fairly close to due east/west; expect larger error.");
  }

  const latValues = buildLatitudeCandidates(hemisphere);
  const lonValues = buildLongitudeCandidates();

  let best = null;
  for (const lat of latValues) {
    for (const lon of lonValues) {
      const candidate = evaluateCandidate(whenUtc, lat, lon, altitudeM, measuredAzimuthDeg, event);
      if (!candidate) continue;
      if (!best || candidate.score < best.score) best = candidate;
    }
  }

  if (!best) {
    warnings.push("WARNING: No candidate solution found in the search domain.");
    return {
      latitude_deg: null,
      longitude_deg: null,
      azimuth_model_deg: null,
      altitude_model_deg: null,
      solar_declination_deg: null,
      horizon_dip_deg: dip,
      effective_horizon_deg: effectiveHorizonDeg,
      residual_azimuth_deg: null,
      residual_altitude_deg: null,
      score: null,
      confidence: 0,
      status: "INVALID INPUT / NO SOLUTION",
      warnings,
      used_azimuth_deg: measuredAzimuthDeg,
    };
  }

  const refined = refineCandidate(best, {
    whenUtc,
    altitudeM,
    measuredAzimuthDeg,
    hemisphere,
    event,
  });

  if (Math.abs(refined.modelDeclinationDeg) < 10) {
    warnings.push("WARNING: Too close to equinox for reliable longitude/latitude recovery from a single sunrise/sunset observation.");
  } else if (Math.abs(refined.modelDeclinationDeg) < 15) {
    warnings.push("CAUTION: Solar declination is low; result may be weak or noisy.");
  }

  if (refined.score > 1.0) {
    warnings.push("WARNING: Final residual is large; input may be inconsistent or too noisy for a reliable solution.");
  } else if (refined.score > 0.1) {
    warnings.push("CAUTION: Final residual is noticeable; treat result cautiously.");
  }

  let status = "GOOD OPERATING CONDITIONS";
  if (warnings.some((w) => w.startsWith("WARNING"))) status = "LOW RELIABILITY";
  else if (warnings.some((w) => w.startsWith("CAUTION"))) status = "USE WITH CAUTION";

  return {
    latitude_deg: refined.lat,
    longitude_deg: refined.lon,
    azimuth_model_deg: refined.modelAzimuthDeg,
    altitude_model_deg: refined.modelAltitudeDeg,
    solar_declination_deg: refined.modelDeclinationDeg,
    horizon_dip_deg: dip,
    effective_horizon_deg: effectiveHorizonDeg,
    residual_azimuth_deg: refined.residualAzimuthDeg,
    residual_altitude_deg: refined.residualAltitudeDeg,
    score: refined.score,
    confidence: computeConfidence({
      score: refined.score,
      declinationDeg: refined.modelDeclinationDeg,
      measuredAzimuthDeg,
      warnings,
    }),
    status,
    warnings,
    used_azimuth_deg: measuredAzimuthDeg,
  };
}

function renderWarnings(warnings) {
  const card = $("warningsCard");
  const list = $("warningsList");
  if (!card || !list) return;

  if (warnings?.length) {
    card.classList.remove("hidden");
    list.innerHTML = warnings
      .map((warning) =>
        `<div class="warning-item ${warning.startsWith("WARNING") ? "warning" : "caution"}">${warning}</div>`
      )
      .join("");
  } else {
    card.classList.add("hidden");
    list.innerHTML = "";
  }
}

function renderResult(result, inputs) {
  $("confidenceCard")?.classList.remove("hidden");
  $("statusBadge")?.classList.remove("hidden");

  if ($("statusBadge")) {
    $("statusBadge").textContent = result.status;
    $("statusBadge").className =
      "status-badge " +
      (result.status === "GOOD OPERATING CONDITIONS"
        ? "status-good"
        : result.status === "USE WITH CAUTION"
        ? "status-caution"
        : "status-bad");
  }

  if ($("confidenceValue")) {
    $("confidenceValue").textContent = result.confidence == null ? "—" : `${result.confidence}%`;
  }

  if ($("bigCoord")) {
    $("bigCoord").textContent =
      result.latitude_deg == null || result.longitude_deg == null
        ? "No valid solution"
        : `${fmtLatDM(result.latitude_deg)} · ${fmtLonDM(result.longitude_deg)}`;
  }

  if ($("subCoord")) $("subCoord").textContent = "";

  if ($("estimatedRows")) {
    $("estimatedRows").innerHTML = [
      row("UTC datetime", `${inputs.date}T${inputs.time}Z`),
      row("Measured azimuth", `${fmtFixed(inputs.azimuth, 10)}°`),
      row("Used azimuth in solve", `${fmtFixed(result.used_azimuth_deg, 10)}°`),
      row("Event", inputs.event),
      row("Hemisphere", inputs.hemisphere),
      row("Observer altitude above MSL", `${fmtFixed(inputs.altitude, 6)} m`),
      row("Estimated latitude", result.latitude_deg == null ? "NO VALID SOLUTION" : fmtLatDM(result.latitude_deg)),
      row("Estimated longitude", result.longitude_deg == null ? "NO VALID SOLUTION" : fmtLonDM(result.longitude_deg)),
    ].join("");
  }

  if ($("modelRows")) {
    $("modelRows").innerHTML = [
      row("Horizon dip", `${fmtFixed(result.horizon_dip_deg, 10)}°`),
      row("Effective horizon", `${fmtFixed(result.effective_horizon_deg, 10)}°`),
      row("Model azimuth at solution", result.azimuth_model_deg == null ? "—" : `${fmtFixed(result.azimuth_model_deg, 10)}°`),
      row("Model altitude at solution", result.altitude_model_deg == null ? "—" : `${fmtFixed(result.altitude_model_deg, 10)}°`),
    ].join("");
  }

  if ($("residualRows")) {
    $("residualRows").innerHTML = [
      row("Residual azimuth", result.residual_azimuth_deg == null ? "—" : `${fmtFixed(result.residual_azimuth_deg, 10)}°`),
      row("Residual altitude", result.residual_altitude_deg == null ? "—" : `${fmtFixed(result.residual_altitude_deg, 10)}°`),
    ].join("");
  }

  renderWarnings(result.warnings);
  updateMap(result.latitude_deg, result.longitude_deg);
}

function onMagneticCorrectionToggleChange() {
  state.magneticCorrectionEnabled = !!$("magCorrectionToggle")?.checked;
  state.lastDeclination = null;
  state.lastCorrectedAzimuth = null;

  const wmmField = $("wmmFileField");
  const wmmFileName = $("wmmFileName");

  if (wmmField) {
    wmmField.classList.toggle("hidden", !state.magneticCorrectionEnabled);
  }

  if (!state.magneticCorrectionEnabled && wmmFileName) {
    wmmFileName.classList.add("hidden");
  }

  updateGuessMarker();
  updateCorrectionInfoPanel();

  if (!state.magneticCorrectionEnabled) {
    if ($("bigCoord")) $("bigCoord").textContent = "Click the map to inspect and guess a position";
    if ($("subCoord")) $("subCoord").textContent = "";
  } else if (state.guessLocation) {
    showGuessPreview();
  }
}

function runEstimate() {
  clearError();
  try {
    const inputs = parseInputs();
    const result = solve(inputs);
    renderResult(result, inputs);
    state.lastResult = { result, inputs };
    updateCorrectionInfoPanel();
  } catch (err) {
    showError(err.message || String(err));
  }
}

function resetForm() {
  const now = new Date();
  const yyyy = now.getUTCFullYear();
  const mm = String(now.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(now.getUTCDate()).padStart(2, "0");

  if ($("dateInput")) $("dateInput").value = `${yyyy}-${mm}-${dd}`;
  if ($("timeInput")) $("timeInput").value = "06:00:00";
  if ($("azimuthInput")) $("azimuthInput").value = "";
  if ($("compassCorrectionInput")) $("compassCorrectionInput").value = "0";
  if ($("altitudeInput")) $("altitudeInput").value = "1";
  if ($("magCorrectionToggle")) $("magCorrectionToggle").checked = false;
  if ($("wmmFileInput")) $("wmmFileInput").value = "";

  state.hemisphere = "N";
  state.event = "sunrise";
  state.guessLocation = null;
  state.magneticCorrectionEnabled = false;
  state.lastDeclination = null;
  state.lastCorrectedAzimuth = null;
  state.lastResult = null;

  setSegmented("hemisphereControl", state.hemisphere);
  setSegmented("eventControl", state.event);

  $("warningsCard")?.classList.add("hidden");
  $("confidenceCard")?.classList.add("hidden");
  $("statusBadge")?.classList.add("hidden");
  $("wmmFileField")?.classList.add("hidden");
  $("wmmFileName")?.classList.add("hidden");

  if ($("bigCoord")) $("bigCoord").textContent = "Click the map to inspect and guess a position";
  if ($("subCoord")) $("subCoord").textContent = "";
  if ($("estimatedRows")) $("estimatedRows").innerHTML = "";
  if ($("modelRows")) $("modelRows").innerHTML = "";
  if ($("residualRows")) $("residualRows").innerHTML = "";
  if ($("wmmFileName")) $("wmmFileName").textContent = "Uploaded file: —";

  updateMap(null, null);
  updateCorrectionInfoPanel();
  clearError();
}

function initApp() {
  bindSegmented("hemisphereControl", "hemisphere");
  bindSegmented("eventControl", "event");

  $("estimateButton")?.addEventListener("click", runEstimate);
  $("resetButton")?.addEventListener("click", resetForm);
  $("magCorrectionToggle")?.addEventListener("change", onMagneticCorrectionToggleChange);

  initBitmapMap();
  initWmmFilePicker();
  resetForm();
  tryLoadWmmViaFetch();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initApp);
} else {
  initApp();
}