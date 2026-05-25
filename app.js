// ==========================================
// 1. CONFIGURACIÓN DE GENESYS CLOUD
// ==========================================
const clientId = '1c1e9531-bc0e-4326-a2bf-177ff46c4f25'; // ⚠️ REEMPLAZA CON EL CLIENT ID DE TU NUEVO CLIENTE
const environment = 'usw2.pure.cloud'; // 🚀 Configurado para la región US West 2 (Oregon)

const platformClient = require('platformClient');
const client = platformClient.ApiClient.instance;
client.setEnvironment(environment);
const redirectUri = window.location.href.split('?')[0];

// ==========================================
// 2. REFERENCIAS Y CACHÉ
// ==========================================
const form = document.getElementById('gamification-form');
const profileSelect = document.getElementById('profileSelect');
const metricSelect = document.getElementById('metricSelect');
const excelInput = document.getElementById('excelInput');
const alertContainer = document.getElementById('alertContainer');
const statusBadge = document.getElementById('statusBadge');
const resultsHeader = document.getElementById('resultsHeader');
const resultsBody = document.getElementById('resultsBody');
const rowCount = document.getElementById('rowCount');

// Cachés para no saturar la API
let currentProfileMetricsMap = {}; // Guardará { "conversion rate": "ID-1234" }
const emailToUserIdCache = {};     // Guardará { "correo@empresa.com": "ID-5678" }

// ==========================================
// 3. INICIALIZACIÓN
// ==========================================
document.addEventListener('DOMContentLoaded', async () => {
    try {
        setStatus('Conectando a Genesys...', 'warning');
        await client.loginImplicitGrant(clientId, redirectUri);
        setStatus('✅ Conectado', 'success');
        await loadProfiles();
    } catch (error) {
        setStatus('Error de Conexión', 'danger');
        showAlert('Error al conectar. Verifica los scopes y Client ID.', 'danger');
    }
});

profileSelect.addEventListener('change', async (e) => {
    const profileId = e.target.value;
    if (profileId) await loadExternalMetricsForProfile(profileId);
});

// ==========================================
// 4. LÓGICA DE PERFILES Y MÉTRICAS
// ==========================================
async function loadProfiles() {
    try {
        const data = await client.callApi('/api/v2/gamification/profiles', 'GET', {}, {}, {}, {}, null, ['PureCloud OAuth'], ['application/json'], ['application/json']);
        profileSelect.innerHTML = '<option value="">-- Seleccione un Perfil --</option>';
        if (data.entities) {
            data.entities.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p.id;
                opt.textContent = p.name;
                profileSelect.appendChild(opt);
            });
        }
    } catch (err) { showAlert('Error al cargar perfiles.', 'danger'); }
}

async function loadExternalMetricsForProfile(profileId) {
    try {
        metricSelect.innerHTML = '<option value="">Filtrando métricas...</option>';
        currentProfileMetricsMap = {}; // Limpiamos el mapa de métricas
        
        // A. Catálogo general
        const extCatalog = await client.callApi('/api/v2/employeeperformance/externalmetrics/definitions', 'GET', {}, {}, {}, {}, null, ['PureCloud OAuth'], ['application/json'], ['application/json']);
        const externalByName = {};
        if (extCatalog.entities) {
            extCatalog.entities.forEach(m => externalByName[m.name.toLowerCase().trim()] = m);
        }

        // B. Métricas del Perfil
        const profileMetrics = await client.callApi(`/api/v2/gamification/profiles/${profileId}/metrics`, 'GET', {}, {}, {}, {}, null, ['PureCloud OAuth'], ['application/json'], ['application/json']);
        
        metricSelect.innerHTML = '<option value="">(Informativo) Métricas válidas para este perfil:</option>';
        const metricsArray = profileMetrics.entities || profileMetrics.metrics || [];

        // C. Cruce y guardado en memoria para el Excel
        metricsArray.forEach(metric => {
            const mName = metric.name ? metric.name.toLowerCase().trim() : '';
            const matched = externalByName[mName];
            if (matched) {
                // Guardamos el nombre y el ID exacto para cuando el Excel lo pida
                currentProfileMetricsMap[mName] = matched.id; 
                
                const opt = document.createElement('option');
                opt.value = matched.id;
                opt.textContent = matched.name;
                opt.disabled = true; // Solo informativo, el Excel manda
                metricSelect.appendChild(opt);
            }
        });
    } catch (err) { showAlert('Error al filtrar métricas.', 'danger'); }
}

// ==========================================
// 5. BÚSQUEDA DE USUARIOS POR CORREO
// ==========================================
async function getUserIdByEmail(email) {
    const cleanEmail = email.toLowerCase().trim();
    if (emailToUserIdCache[cleanEmail]) return emailToUserIdCache[cleanEmail];

    const payload = {
        query: [{ fields: ["email"], value: cleanEmail, type: "EXACT" }]
    };

    const response = await client.callApi('/api/v2/users/search', 'POST', {}, {}, {}, {}, payload, ['PureCloud OAuth'], ['application/json'], ['application/json']);
    
    if (response.total > 0 && response.results[0]) {
        emailToUserIdCache[cleanEmail] = response.results[0].id;
        return response.results[0].id;
    } else {
        throw new Error(`Correo no encontrado en Genesys`);
    }
}

// ==========================================
// 6. LECTURA EXCEL Y RENDERIZADO
// ==========================================
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(worksheet, { defval: '' }));
      } catch (error) { reject(error); }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function renderTable(data) {
  resultsHeader.innerHTML = ''; resultsBody.innerHTML = '';
  if (!data.length) { rowCount.textContent = 'El archivo está vacío.'; return; }

  const headers = Object.keys(data[0]);
  headers.push('Estado API');

  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    resultsHeader.appendChild(th);
  });

  data.forEach((row, index) => {
    const tr = document.createElement('tr');
    headers.forEach((h) => {
      const td = document.createElement('td');
      if (h === 'Estado API') {
        td.id = `statusCell-${index}`;
        td.innerHTML = '<span class="badge bg-secondary">Pendiente</span>';
      } else {
        td.textContent = (row[h] instanceof Date) ? row[h].toISOString().split('T')[0] : row[h];
      }
      tr.appendChild(td);
    });
    resultsBody.appendChild(tr);
  });
  rowCount.textContent = `${data.length} fila(s) leídas.`;
}

function showAlert(msg, type = 'danger') {
  alertContainer.innerHTML = `<div class="alert alert-${type} alert-dismissible fade show">${msg}<button type="button" class="btn-close" data-bs-dismiss="alert"></button></div>`;
}
function setStatus(msg, variant = 'secondary') {
  statusBadge.textContent = msg; statusBadge.className = `badge bg-${variant}`;
}

// ==========================================
// 7. MOTOR DE SUBIDA A GENESYS
// ==========================================
async function uploadToGenesysCloud(data) {
    let successCount = 0, errorCount = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const statusCell = document.getElementById(`statusCell-${i}`);
        statusCell.innerHTML = '<span class="badge bg-info">Procesando...</span>';

        const email = row.CorreoElectronico || row.correo || row.Email;
        const metricName = row.NombreMetrica || row.Metrica || row.MetricName;
        const valor = row.Valor || row.valor || row.Value;
        let fecha = row.Fecha || row.fecha || row.Date;

        if (!email || !metricName || valor === '' || !fecha) {
            statusCell.innerHTML = '<span class="badge bg-danger">Faltan datos en Excel</span>';
            errorCount++; continue;
        }

        try {
            const metricId = currentProfileMetricsMap[metricName.toLowerCase().trim()];
            if (!metricId) throw new Error(`La métrica '${metricName}' no es válida en este perfil.`);

            const userId = await getUserIdByEmail(email);

            let formattedDate;
            if (fecha instanceof Date) {
                formattedDate = fecha.toISOString().split('T')[0] + "T12:00:00Z";
            } else {
                formattedDate = new Date(fecha).toISOString().split('T')[0] + "T12:00:00Z";
            }

            const payload = {
                items: [{
                    userId: userId,
                    metricId: metricId,
                    dateOccurred: formattedDate,
                    value: parseFloat(valor),
                    count: 1
                }]
            };

            await client.callApi('/api/v2/employeeperformance/externalmetrics/data', 'POST', {}, {}, {}, {}, payload, ['PureCloud OAuth'], ['application/json'], ['application/json']);
            
            statusCell.innerHTML = '<span class="badge bg-success">Completado ✅</span>';
            successCount++;

        } catch (error) {
            const msg = error.message || (error.status ? `API Error ${error.status}` : 'Error');
            statusCell.innerHTML = `<span class="badge bg-danger" title="${msg}">${msg}</span>`;
            errorCount++;
        }
    }

    if (errorCount === 0) showAlert('🎉 Todos los registros se enviaron correctamente.', 'success');
    else showAlert(`Proceso completado. Éxitos: ${successCount}. Errores: ${errorCount}.`, 'warning');
}

// ==========================================
// 8. EVENTO DEL FORMULARIO
// ==========================================
form.addEventListener('submit', async (event) => {
  event.preventDefault();
  alertContainer.innerHTML = '';
  
  const profileId = profileSelect.value;
  const file = excelInput.files[0];

  if (!profileId) { showAlert('⚠️ Selecciona un perfil de Gamificación.', 'warning'); return; }
  if (!file) { showAlert('⚠️ Adjunta un archivo Excel.', 'warning'); return; }

  setStatus('Leyendo Excel...', 'info');

  try {
    const data = await readExcelFile(file);
    renderTable(data);
    
    setStatus('Subiendo a Genesys...', 'primary');
    await uploadToGenesysCloud(data);
    
    setStatus('Proceso Finalizado', 'success');
  } catch (error) {
    showAlert(`Error crítico al procesar: ${error.message}`, 'danger');
    setStatus('Error', 'danger');
  }
});

// ==========================================
// 9. FUNCIONES GLOBALES DE SOPORTE
// ==========================================
window.downloadSampleExcel = function() {
    try {
        const ws_data = [
            ["CorreoElectronico", "NombreMetrica", "Valor", "Fecha"],
            ["agente.prueba@tuempresa.com", "Conversion Rate", "50", "2026-05-14"]
        ];
        const ws = XLSX.utils.aoa_to_sheet(ws_data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Gamification");
        XLSX.writeFile(wb, "Plantilla_Metricas_Actualizada.xlsx");
    } catch (err) {
        console.error("Error al descargar plantilla:", err);
        alert("No se pudo generar el archivo Excel. Asegúrate de que la librería SheetJS se cargó correctamente.");
    }
};
