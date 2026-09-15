let accessKey = null;
let archivoSeleccionado = null;
let sessionTimeout = null;
let sessionInterval = null;
let sessionExpiresAt = null;

const SESSION_MINUTES = 15;

// ── Login ────────────────────────────────────────────────────────────────
function toggleKey() {
  const input = document.getElementById("accessKey");
  input.type = input.type === "password" ? "text" : "password";
}

async function doLogin() {
  const key = document.getElementById("accessKey").value.trim();
  const err = document.getElementById("loginError");
  const btnText = document.getElementById("loginBtnText");
  const spinner = document.getElementById("loginSpinner");

  err.classList.remove("show");
  if (!key) {
    err.textContent = "Ingresa una clave";
    err.classList.add("show");
    return;
  }

  btnText.textContent = "Verificando...";
  spinner.classList.add("show");

  try {
    const fd = new FormData();
    fd.append("key", key);
    const res = await fetch("/verificar-clave", { method: "POST", body: fd });
    const data = await res.json();

    if (res.ok && data.ok) {
      accessKey = key;
      document.getElementById("loginOverlay").style.display = "none";
      iniciarSesion();
    } else {
      err.textContent = "Clave de acceso incorrecta";
      err.classList.add("show");
    }
  } catch (e) {
    err.textContent = "No se pudo conectar con el servidor";
    err.classList.add("show");
  } finally {
    btnText.textContent = "Entrar";
    spinner.classList.remove("show");
  }
}

function iniciarSesion() {
  document.getElementById("sessionBadge").style.display = "flex";
  resetSessionTimer();
  ["click", "keydown", "mousemove"].forEach((ev) =>
    document.addEventListener(ev, resetSessionTimer)
  );
}

function resetSessionTimer() {
  clearTimeout(sessionTimeout);
  clearInterval(sessionInterval);
  sessionExpiresAt = Date.now() + SESSION_MINUTES * 60 * 1000;

  sessionInterval = setInterval(() => {
    const remaining = Math.max(0, sessionExpiresAt - Date.now());
    const mins = Math.floor(remaining / 60000);
    const secs = Math.floor((remaining % 60000) / 1000);
    document.getElementById("sessionTimer").textContent =
      `${mins}:${secs.toString().padStart(2, "0")}`;
  }, 1000);

  sessionTimeout = setTimeout(logout, SESSION_MINUTES * 60 * 1000);
}

function logout() {
  accessKey = null;
  clearTimeout(sessionTimeout);
  clearInterval(sessionInterval);
  document.getElementById("sessionBadge").style.display = "none";
  document.getElementById("loginOverlay").style.display = "flex";
  document.getElementById("accessKey").value = "";
  nuevo();
}

// ── Selección de archivo ─────────────────────────────────────────────────
const dropzone = document.getElementById("dropzone");

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.style.borderColor = "var(--blue)";
});
dropzone.addEventListener("dragleave", () => {
  dropzone.style.borderColor = "";
});
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.style.borderColor = "";
  if (e.dataTransfer.files.length) {
    document.getElementById("fileInput").files = e.dataTransfer.files;
    onFileChange({ target: { files: e.dataTransfer.files } });
  }
});

function onFileChange(event) {
  const file = event.target.files[0];
  ocultarAlerta();
  if (!file) return;

  const nombreOk = /\.(xlsx|xlsm|xls)$/i.test(file.name);
  if (!nombreOk) {
    mostrarAlerta("Solo se aceptan archivos .xlsx, .xlsm o .xls");
    return;
  }

  archivoSeleccionado = file;
  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileSelected").classList.add("show");

  const btn = document.getElementById("btnProcesar");
  btn.disabled = false;
  document.getElementById("btnText").textContent = "Generar SilverV2";
}

function quitarArchivo() {
  archivoSeleccionado = null;
  document.getElementById("fileInput").value = "";
  document.getElementById("fileSelected").classList.remove("show");
  document.getElementById("btnProcesar").disabled = true;
  document.getElementById("btnText").textContent = "Selecciona un archivo para continuar";
}

function mostrarAlerta(msg) {
  document.getElementById("alertaMsg").textContent = msg;
  document.getElementById("alerta").classList.add("show");
}
function ocultarAlerta() {
  document.getElementById("alerta").classList.remove("show");
}

// ── Procesar ─────────────────────────────────────────────────────────────
async function procesar() {
  if (!archivoSeleccionado || !accessKey) return;

  const btn = document.getElementById("btnProcesar");
  const btnText = document.getElementById("btnText");
  const spinner = document.getElementById("spinner");

  ocultarAlerta();
  btn.disabled = true;
  spinner.classList.add("show");
  btnText.textContent = "Procesando...";

  try {
    const fd = new FormData();
    fd.append("key", accessKey);
    fd.append("file", archivoSeleccionado);

    const res = await fetch("/procesar", { method: "POST", body: fd });
    const data = await res.json();

    if (!res.ok) {
      mostrarAlerta(data.error || "Ocurrió un error al procesar el archivo");
      btn.disabled = false;
      return;
    }

    mostrarResultado(data.resumen, archivoSeleccionado.name);
  } catch (e) {
    mostrarAlerta("No se pudo conectar con el servidor");
    btn.disabled = false;
  } finally {
    spinner.classList.remove("show");
    btnText.textContent = "Generar SilverV2";
  }
}

function mostrarResultado(resumen, nombreArchivo) {
  document.getElementById("cardUpload").style.display = "none";
  document.getElementById("resultado").style.display = "block";

  document.getElementById("resNombre").textContent = `— ${nombreArchivo}`;
  document.getElementById("mFilas").textContent = resumen.filas_totales ?? "—";
  document.getElementById("mColumnas").textContent = resumen.columnas ?? "—";
  document.getElementById("mHojas").textContent =
    (resumen.hojas_encontradas || []).length;

  const desde = resumen.rango_year_month?.desde || "—";
  const hasta = resumen.rango_year_month?.hasta || "—";
  document.getElementById("mRango").textContent = `${desde} → ${hasta}`;

  document.getElementById("mOmitidas").textContent =
    resumen.filas_omitidas_sin_fecha ?? "0";

  const tags = document.getElementById("resTags");
  tags.innerHTML = "";
  const porStatus = resumen.por_status || {};
  Object.entries(porStatus).forEach(([status, cantidad]) => {
    const tag = document.createElement("span");
    tag.className = "tag";
    tag.textContent = `${status}: ${cantidad}`;
    tags.appendChild(tag);
  });
  if (resumen.hojas_ignoradas && resumen.hojas_ignoradas.length) {
    const tag = document.createElement("span");
    tag.className = "tag";
    tag.textContent = `Hojas no encontradas: ${resumen.hojas_ignoradas.join(", ")}`;
    tags.appendChild(tag);
  }
}

async function descargar() {
  window.location.href = "/descargar";
}

function nuevo() {
  document.getElementById("resultado").style.display = "none";
  document.getElementById("cardUpload").style.display = "block";
  quitarArchivo();
  ocultarAlerta();
}