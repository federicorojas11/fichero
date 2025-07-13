// ==========================================
// FUNCIÓN CONSULTAR LEGAJO 
// ==========================================
function consultarLegajo(numeroLegajo) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (!hoja) {
      throw new Error('No se pudo acceder a la hoja de cálculo activa');
    }
    
    const datos = hoja.getDataRange().getValues();

    const resultados = [];
    let ultimoRegistro = null;
    let estadoActual = 'NO REGISTRADO';

    // Buscar filas coincidentes (columna C = índice 2)
    for (let i = 3; i < datos.length; i++) { // Empieza desde fila 4
      const celdaLegajo = datos[i][2]; // Columna C - Número LPU
      if (celdaLegajo == numeroLegajo) {
        // Convertir fechas
        const fechaRetiro = datos[i][1] ? formatearFechaCliente(datos[i][1]) : '-';
        const fechaEntrada = datos[i][6] ? formatearFechaCliente(datos[i][6]) : '-';

        const registro = {
          fila: i + 1,
          fechaRetiro,
          credencialRetira: datos[i][3] || '-', // D - Credencial Salida
          division: datos[i][4] || '-',         // E - División
          credencialEntrada: datos[i][5] || '-',// F - Credencial Entrada
          fechaEntrada
        };

        resultados.push(registro);
        
        // Actualizar último registro (más reciente)
        if (!ultimoRegistro || new Date(fechaRetiro.split('/').reverse().join('-')) > 
            new Date(ultimoRegistro.fechaRetiro.split('/').reverse().join('-'))) {
          ultimoRegistro = registro;
        }
      }
    }

    if (resultados.length === 0) {
      return {
        success: true,
        estado: "NO REGISTRADO",
        numeroLegajo,
        message: `No se encontró ningún registro para el legajo ${numeroLegajo}`
      };
    }

    // Determinar estado actual
    estadoActual = ultimoRegistro.fechaEntrada !== '-' ? 'DEVUELTO' : 'EN USO';

   return {
      success: true,
      estado: estadoActual === 'DEVUELTO' ? "EN ARCHIVO" : "EN SALIDA",
      numeroLegajo,
      estadoActual,
      ultimoRegistro: {
        fechaSalida: resultados[0].fechaRetiro,
        division: resultados[0].division,
        credencialRetira: resultados[0].credencialRetira,
        fechaEntrada: resultados[0].fechaEntrada,
        credencialEntrada: resultados[0].credencialEntrada
      },
      historial: resultados
    };

  } catch (error) {
    console.log("Error en consultarLegajo:", error.message);
    return {
      success: false,
      message: "Error interno al consultar el legajo: " + error.message
    };
  }
}
// Función auxiliar para convertir cualquier formato de fecha a "dd/MM/yyyy"
function formatearFechaCliente(fecha) {
  if (!fecha) return null;

  let fechaObj;

  if (typeof fecha === 'string') {
    // Si ya es formato de texto como "23/4/2025"
    const partes = fecha.split('/');
    if (partes.length === 3) {
      const dia = partes[0];
      const mes = partes[1];
      const anio = partes[2];
      return `${dia}/${mes}/${anio}`;
    }
    return fecha; // Otro formato de string desconocido
  }

  if (fecha instanceof Date) {
    // Si es un objeto Date devuelto por Apps Script
    const dia = ('0' + fecha.getDate()).slice(-2);
    const mes = ('0' + (fecha.getMonth() + 1)).slice(-2); // getMonth() es 0-based
    const anio = fecha.getFullYear();
    return `${dia}/${mes}/${anio}`;
  }

  return null;
}

function convertirFecha(fecha) {
  if (!fecha) return null;
  
  // Si ya es cadena con formato dd/MM/yyyy
  if (typeof fecha === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(fecha)) {
    const [dia, mes, anio] = fecha.split('/');
    return new Date(anio, mes - 1, dia); // Meses en JS son 0-based
  }

  // Si es objeto Date de JavaScript
  if (fecha instanceof Date) {
    return fecha;
  }

  // Si es un valor numérico (timestamp o serial date)
  if (typeof fecha === 'number') {
    const jsDate = new Date(fecha);
    return isNaN(jsDate.getTime()) ? null : jsDate;
  }

  // Otros formatos no reconocidos
  return null;
}


/** function mostrarModalEstado(response) {
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Roboto', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
            color: #333;
          }
          .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          .header {
            padding: 20px;
            background: #3f51b5;
            color: white;
            display: flex;
            align-items: center;
          }
          .header-icon {
            font-size: 36px;
            margin-right: 15px;
          }
          .header-content h1 {
            margin: 0;
            font-size: 24px;
            font-weight: 500;
          }
          .header-content p {
            margin: 5px 0 0;
            opacity: 0.9;
          }
          .estado-section {
            padding: 25px;
          }
          .estado-card {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 3px 10px rgba(0,0,0,0.08);
            margin-bottom: 25px;
          }
          .estado-header {
            padding: 20px;
            color: white;
            display: flex;
            align-items: center;
            font-size: 20px;
          }
          .estado-icon {
            margin-right: 15px;
            font-size: 32px;
          }
          .estado-body {
            padding: 20px;
          }
          .estado-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 16px;
          }
          .estado-table th {
            text-align: left;
            padding: 12px 0;
            width: 40%;
            color: #616161;
            font-weight: 500;
          }
          .estado-table td {
            padding: 12px 0;
            font-weight: 400;
            font-size: 17px;
          }
          .historial-section {
            padding: 20px;
          }
          .historial-title {
            margin-top: 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #3f51b5;
            color: #3f51b5;
          }
          .historial-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
          }
          .historial-table th {
            background-color: #3f51b5;
            color: white;
            padding: 12px 15px;
            text-align: left;
          }
          .historial-table td {
            padding: 10px 15px;
            border-bottom: 1px solid #eee;
          }
          .historial-table tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          .btn-cerrar {
            display: block;
            margin: 20px auto;
            padding: 12px 30px;
            background: #3f51b5;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            cursor: pointer;
            transition: background 0.3s;
          }
          .btn-cerrar:hover {
            background: #303f9f;
          }
          
          /* Colores de estado *//**
          .estado-en-archivo .estado-header {
            background-color: #4caf50;
          }
          .estado-en-salida .estado-header {
            background-color: #ff9800;
          }
          .estado-no-registrado .estado-header {
            background-color: #f44336;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <div class="header-icon">📋</div>
            <div class="header-content">
              <h1>Estado del Legajo</h1>
              <p>Consulta realizada: <?= new Date().toLocaleString() ?></p>
            </div>
          </div>
          
          <div class="estado-section">
            <? const estadoClass = response.estado === "EN ARCHIVO" ? 'estado-en-archivo' : 
                                 response.estado === "EN SALIDA" ? 'estado-en-salida' : 
                                 'estado-no-registrado' ?>
            <? const estadoIcon = response.estado === "EN ARCHIVO" ? '📁' : 
                              response.estado === "EN SALIDA" ? '🚶' : '❓' ?>
            <? const estadoText = response.estado === "EN ARCHIVO" ? 'En Archivo' : 
                              response.estado === "EN SALIDA" ? 'En Salida' : 'No Registrado' ?>
            
            <div class="estado-card <?= estadoClass ?>">
              <div class="estado-header">
                <span class="estado-icon"><?= estadoIcon ?></span>
                <h3><?= estadoText ?></h3>
              </div>
              <div class="estado-body">
                <table class="estado-table">
                  <tr>
                    <th>Número de Legajo</th>
                    <td><?= response.numeroLegajo ?></td>
                  </tr>
                  <? if (response.estado !== "NO REGISTRADO") { ?>
                    <tr>
                      <th>Fecha de Salida</th>
                      <td><?= formatDate(response.ultimoRegistro.fechaSalida) ?></td>
                    </tr>
                    <tr>
                      <th>División</th>
                      <td><?= response.ultimoRegistro.division ?></td>
                    </tr>
                    <tr>
                      <th>Retirado por</th>
                      <td><?= response.ultimoRegistro.credencialRetira ?></td>
                    </tr>
                    <? if (response.estado === "EN ARCHIVO") { ?>
                      <tr>
                        <th>Fecha de Entrada</th>
                        <td><?= formatDate(response.ultimoRegistro.fechaEntrada) ?></td>
                      </tr>
                      <tr>
                        <th>Recibido por</th>
                        <td><?= response.ultimoRegistro.credencialEntrada ?></td>
                      </tr>
                    <? } ?>
                  <? } else { ?>
                    <tr>
                      <td colspan="2" style="text-align:center;padding:20px 0;">
                        El legajo no se encuentra en el sistema
                      </td>
                    </tr>
                  <? } ?>
                </table>
              </div>
            </div>
          </div>
          
          <? if (response.historial && response.historial.length > 0) { ?>
            <div class="historial-section">
              <h3 class="historial-title">Historial Completo</h3>
              <table class="historial-table">
                <thead>
                  <tr>
                    <th>Fecha Salida</th>
                    <th>División</th>
                    <th>Retiró</th>
                    <th>Fecha Entrada</th>
                    <th>Recibió</th>
                    <th>Estado</th>
                  </tr>
                </thead>
                <tbody>
                  <? for (let i = 0; i < response.historial.length; i++) { ?>
                    <? const reg = response.historial[i] ?>
                    <? const tieneEntrada = reg.fechaEntrada && reg.fechaEntrada !== '-' ?>
                    <tr>
                      <td><?= formatDate(reg.fechaSalida) ?></td>
                      <td><?= reg.division ?></td>
                      <td><?= reg.credencialRetira ?></td>
                      <td><?= tieneEntrada ? formatDate(reg.fechaEntrada) : '-' ?></td>
                      <td><?= tieneEntrada ? reg.credencialEntrada : '-' ?></td>
                      <td><?= tieneEntrada ? 'Devuelto' : 'En uso' ?></td>
                    </tr>
                  <? } ?>
                </tbody>
              </table>
            </div>
          <? } ?>
          
          <button class="btn-cerrar" onclick="google.script.host.close()">Cerrar</button>
        </div>
        
        <script>
          // Función para formatear fechas
          function formatDate(date) {
            if (!date) return '';
            if (typeof date === 'string') return date;
            try {
              const d = new Date(date);
              return d.toLocaleDateString('es-AR');
            } catch(e) {
              return date.toString();
            }
          }
          
          // Aplicar formato a todas las fechas en la tabla
          document.addEventListener('DOMContentLoaded', function() {
            const dateCells = document.querySelectorAll('td');
            dateCells.forEach(cell => {
              if (cell.textContent.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) {
                cell.textContent = formatDate(cell.textContent);
              }
            });
          });
        </script>
      </body>
    </html>
  `;

  // Procesar el template
  const html = HtmlService.createTemplate(htmlTemplate);
  html.response = response;
  html.formatDate = function(date) {
    if (!date) return '';
    try {
      if (typeof date === 'string' && date.includes('/')) return date;
      const d = new Date(date);
      return d.toLocaleDateString('es-AR');
    } catch(e) {
      return date.toString();
    }
  };
  
  return html.evaluate().getContent();
}*/

function generarModalDetalles(numeroLegajo) {
  // Primero obtener los datos nuevamente
  const response = consultarLegajo(numeroLegajo);
  
  if (!response.success) {
    return `<h2>Error</h2><p>${response.message}</p>`;
  }

  // Crear el HTML del modal
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          /* Todos los estilos CSS que estaban antes */
           <style>
          body {
            font-family: 'Roboto', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
            color: #333;
          }
          .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          .header {
            padding: 20px;
            background: #3f51b5;
            color: white;
            display: flex;
            align-items: center;
          }
          .header-icon {
            font-size: 36px;
            margin-right: 15px;
          }
          .header-content h1 {
            margin: 0;
            font-size: 24px;
            font-weight: 500;
          }
          .header-content p {
            margin: 5px 0 0;
            opacity: 0.9;
          }
          .estado-section {
            padding: 25px;
          }
          .estado-card {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 3px 10px rgba(0,0,0,0.08);
            margin-bottom: 25px;
          }
          .estado-header {
            padding: 20px;
            color: white;
            display: flex;
            align-items: center;
            font-size: 20px;
          }
          .estado-icon {
            margin-right: 15px;
            font-size: 32px;
          }
          .estado-body {
            padding: 20px;
          }
          .estado-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 16px;
          }
          .estado-table th {
            text-align: left;
            padding: 12px 0;
            width: 40%;
            color: #616161;
            font-weight: 500;
          }
          .estado-table td {
            padding: 12px 0;
            font-weight: 400;
            font-size: 17px;
          }
          .historial-section {
            padding: 20px;
          }
          .historial-title {
            margin-top: 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #3f51b5;
            color: #3f51b5;
          }
          .historial-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
          }
          .historial-table th {
            background-color: #3f51b5;
            color: white;
            padding: 12px 15px;
            text-align: left;
          }
          .historial-table td {
            padding: 10px 15px;
            border-bottom: 1px solid #eee;
          }
          .historial-table tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          .btn-cerrar {
            display: block;
            margin: 20px auto;
            padding: 12px 30px;
            background: #3f51b5;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            cursor: pointer;
            transition: background 0.3s;
          }
          .btn-cerrar:hover {
            background: #303f9f;
          }
          
          /* Colores de estado *//**
          .estado-en-archivo .estado-header {
            background-color: #4caf50;
          }
          .estado-en-salida .estado-header {
            background-color: #ff9800;
          }
          .estado-no-registrado .estado-header {
            background-color: #f44336;
          }
          /* ... (copiar todos los estilos CSS anteriores) ... */
        </style>
      </head>
      <body>
        <div class="container">
          <!-- Todo el contenido HTML que estaba antes -->
          
          <div class="header">
            <div class="header-icon">📋</div>
            <div class="header-content">
              <h1>Estado del Legajo</h1>
              <p>Consulta realizada: <?= new Date().toLocaleString() ?></p>
            </div>
          </div>
          
          <div class="estado-section">
            <? const estadoClass = response.estado === "EN ARCHIVO" ? 'estado-en-archivo' : 
                                 response.estado === "EN SALIDA" ? 'estado-en-salida' : 
                                 'estado-no-registrado' ?>
            <? const estadoIcon = response.estado === "EN ARCHIVO" ? '📁' : 
                              response.estado === "EN SALIDA" ? '🚶' : '❓' ?>
            <? const estadoText = response.estado === "EN ARCHIVO" ? 'En Archivo' : 
                              response.estado === "EN SALIDA" ? 'En Salida' : 'No Registrado' ?>
            
            <div class="estado-card <?= estadoClass ?>">
              <div class="estado-header">
                <span class="estado-icon"><?= estadoIcon ?></span>
                <h3><?= estadoText ?></h3>
              </div>
              <div class="estado-body">
                <table class="estado-table">
                  <tr>
                    <th>Número de Legajo</th>
                    <td><?= response.numeroLegajo ?></td>
                  </tr>
                  <? if (response.estado !== "NO REGISTRADO") { ?>
                    <tr>
                      <th>Fecha de Salida</th>
                      <td><?= formatDate(response.ultimoRegistro.fechaSalida) ?></td>
                    </tr>
                    <tr>
                      <th>División</th>
                      <td><?= response.ultimoRegistro.division ?></td>
                    </tr>
                    <tr>
                      <th>Retirado por</th>
                      <td><?= response.ultimoRegistro.credencialRetira ?></td>
                    </tr>
                    <? if (response.estado === "EN ARCHIVO") { ?>
                      <tr>
                        <th>Fecha de Entrada</th>
                        <td><?= formatDate(response.ultimoRegistro.fechaEntrada) ?></td>
                      </tr>
                      <tr>
                        <th>Recibido por</th>
                        <td><?= response.ultimoRegistro.credencialEntrada ?></td>
                      </tr>
                    <? } ?>
                  <? } else { ?>
                    <tr>
                      <td colspan="2" style="text-align:center;padding:20px 0;">
                        El legajo no se encuentra en el sistema
                      </td>
                    </tr>
                  <? } ?>
                </table>
              </div>
            </div>
          </div>
          
          <? if (response.historial && response.historial.length > 0) { ?>
            <div class="historial-section">
              <h3 class="historial-title">Historial Completo</h3>
              <table class="historial-table">
                <thead>
                  <tr>
                    <th>Fecha Salida</th>
                    <th>División</th>
                    <th>Retiró</th>
                    <th>Fecha Entrada</th>
                    <th>Recibió</th>
                    <th>Estado</th>
                  </tr>
                </thead>
                <tbody>
                  <? for (let i = 0; i < response.historial.length; i++) { ?>
                    <? const reg = response.historial[i] ?>
                    <? const tieneEntrada = reg.fechaEntrada && reg.fechaEntrada !== '-' ?>
                    <tr>
                      <td><?= formatDate(reg.fechaSalida) ?></td>
                      <td><?= reg.division ?></td>
                      <td><?= reg.credencialRetira ?></td>
                      <td><?= tieneEntrada ? formatDate(reg.fechaEntrada) : '-' ?></td>
                      <td><?= tieneEntrada ? reg.credencialEntrada : '-' ?></td>
                      <td><?= tieneEntrada ? 'Devuelto' : 'En uso' ?></td>
                    </tr>
                  <? } ?>
                </tbody>
              </table>
            </div>
          <? } ?>
          
          ${generarContenidoModal(response)}
          <button onclick="google.script.host.close()" class="btn-cerrar">Cerrar</button>
        </div>
      </body>
    </html>
  `;
  
  return html;
}

/*function generarContenidoModal(response) {
  // Generar el contenido dinámico del modal
  let html = `
    <div class="header">
      <div class="header-icon">📋</div>
      <div class="header-content">
        <h1>Estado del Legajo ${response.numeroLegajo}</h1>
        <p>Consulta realizada: ${new Date().toLocaleString()}</p>
      </div>
    </div>
  `;
  
  // ... (continuar con el resto del contenido como antes) ...
  
  return html;
}*/

function mostrarModalLegajo(numeroLegajo) {
  const response = consultarLegajo(numeroLegajo);
  
  if (!response.success) {
    SpreadsheetApp.getUi().alert("Error", response.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Crear el HTML del modal
  const htmlOutput = HtmlService
    .createHtmlOutput(generarHtmlModal(response))
    .setWidth(1000)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Estado del Legajo');
}

function generarHtmlModal(response) {
  const estadoClass = response.estado === "EN ARCHIVO" ? 'estado-en-archivo' : 
                     response.estado === "EN SALIDA" ? 'estado-en-salida' : 
                     'estado-no-registrado';
  const estadoIcon = response.estado === "EN ARCHIVO" ? '📁' : 
                    response.estado === "EN SALIDA" ? '🚶' : '❓';
  const estadoText = response.estado === "EN ARCHIVO" ? 'En Archivo' : 
                    response.estado === "EN SALIDA" ? 'En Salida' : 'No Registrado';

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Roboto', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
          }
          .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          /* ... (copiar todos los estilos CSS anteriores) ... */
        </style>
      </head>
      <body>
        <div class="container">
          <!-- Contenido del modal -->
          ${generarContenidoModal(response, estadoClass, estadoIcon, estadoText)}
          
          <button onclick="google.script.host.close()" 
                  style="display:block; margin:20px auto; padding:12px 30px; 
                         background:#3f51b5; color:white; border:none; 
                         border-radius:4px; font-size:16px; cursor:pointer;">
            Cerrar
          </button>
        </div>
      </body>
    </html>
  `;
}

function generarContenidoModal(response, estadoClass, estadoIcon, estadoText) {
  let html = `
    <div class="header">
      <div class="header-icon">📋</div>
      <div class="header-content">
        <h1>Estado del Legajo ${response.numeroLegajo}</h1>
        <p>Consulta realizada: ${new Date().toLocaleString()}</p>
      </div>
    </div>
    
    <div class="estado-card ${estadoClass}">
      <div class="estado-header">
        <span class="estado-icon">${estadoIcon}</span>
        <h3>${estadoText}</h3>
      </div>
      <div class="estado-body">
        <table class="estado-table">
          <tr>
            <th>Número de Legajo</th>
            <td>${response.numeroLegajo}</td>
          </tr>
  `;

  if (response.estado !== "NO REGISTRADO") {
    html += `
      <tr>
        <th>Fecha de Salida</th>
        <td>${formatDate(response.ultimoRegistro.fechaSalida)}</td>
      </tr>
      <tr>
        <th>División</th>
        <td>${response.ultimoRegistro.division}</td>
      </tr>
      <tr>
        <th>Retirado por</th>
        <td>${response.ultimoRegistro.credencialRetira}</td>
      </tr>
    `;
    
    if (response.estado === "EN ARCHIVO") {
      html += `
        <tr>
          <th>Fecha de Entrada</th>
          <td>${formatDate(response.ultimoRegistro.fechaEntrada)}</td>
        </tr>
        <tr>
          <th>Recibido por</th>
          <td>${response.ultimoRegistro.credencialEntrada}</td>
        </tr>
      `;
    }
  } else {
    html += `
      <tr>
        <td colspan="2" style="text-align:center;padding:20px 0;">
          El legajo no se encuentra en el sistema
        </td>
      </tr>
    `;
  }

  html += `
        </table>
      </div>
    </div>
  `;

  // Agregar historial si existe
  if (response.historial && response.historial.length > 0) {
    html += `
      <div class="historial-section">
        <h3 class="historial-title">Historial Completo</h3>
        <table class="historial-table">
          <thead>
            <tr>
              <th>Fecha Salida</th>
              <th>División</th>
              <th>Retiró</th>
              <th>Fecha Entrada</th>
              <th>Recibió</th>
              <th>Estado</th>
            </tr>
          </thead>
          <tbody>
    `;

    response.historial.forEach(reg => {
      const tieneEntrada = reg.fechaEntrada && reg.fechaEntrada !== '-';
      html += `
        <tr>
          <td>${formatDate(reg.fechaSalida)}</td>
          <td>${reg.division}</td>
          <td>${reg.credencialRetira}</td>
          <td>${tieneEntrada ? formatDate(reg.fechaEntrada) : '-'}</td>
          <td>${tieneEntrada ? reg.credencialEntrada : '-'}</td>
          <td>${tieneEntrada ? 'Devuelto' : 'En uso'}</td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;
  }

  return html;
}

function formatDate(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string' && date.includes('/')) return date;
    const d = new Date(date);
    return d.toLocaleDateString('es-AR');
  } catch(e) {
    return date.toString();
  }
}

// ==========================================
// FUNCIÓN GUARDAR ENTRADA
// ==========================================
function guardarEntrada(fechaEntrada, numeroLegajo, credencialEntrada) {
  try {
    const hoja = SpreadsheetApp.getActiveSheet();
    
    if (!hoja) {
      throw new Error('No se pudo acceder a la hoja de cálculo activa');
    }

    // Validar parámetros obligatorios
    if (!fechaEntrada || !numeroLegajo || !credencialEntrada) {
      throw new Error('Todos los campos son obligatorios');
    }
    
    // Validar formato de legajo (5-6 dígitos)
    if (!/^\d{5,6}$/.test(numeroLegajo.toString())) {
      throw new Error('El número de legajo debe tener entre 5 y 6 dígitos');
    }
    
    // Validar formato de credencial (5 dígitos)
    if (!/^\d{5}$/.test(credencialEntrada.toString())) {
      throw new Error('La credencial debe tener exactamente 5 dígitos');
    }
    // Verificar duplicados/modificaciones
    const resultadoVerificacion = verificarEntradaDuplicada(
      hoja, numeroLegajo, fechaEntrada, credencialEntrada
    );

    if (resultadoVerificacion) {
      if (resultadoVerificacion.tipo === "igual") {
        return {
          success: false,
          message: `Ya existe una entrada registrada para el legajo ${numeroLegajo} con estos mismos datos en la fecha ${formatearFecha(fechaEntrada)}.`,
          notification: 'info'
        };
      } else if (resultadoVerificacion.tipo === "diferente") {
        return {
          success: false,
          modal: true,
          message: "Ya existe una entrada para este legajo en la fecha seleccionada.",
          datosExistentes: resultadoVerificacion.datosExistentes,
          datosNuevos: {
            fechaEntrada: formatearFecha(fechaEntrada),
            numeroLegajo: numeroLegajo,
            credencialEntrada: credencialEntrada
          },
          filaCoincidente: resultadoVerificacion.fila
        };
      }
    }
    // Buscar TODAS las filas con este legajo
    const filasEncontradas = buscarTodasLasFilasLegajo(numeroLegajo, 3); // Columna C = 3
      
    // Filtrar filas sin entrada registrada
    const filasSinEntrada = filasEncontradas.filter(fila => {
      const credencial = hoja.getRange(fila, 6).getValue(); // Columna F
      const fecha = hoja.getRange(fila, 7).getValue();      // Columna G
      return !credencial && !fecha;
    });
    
    if (filasSinEntrada.length === 0) {
    // Obtener datos desde la fila 4 (inclusive) en adelante, columnas B-G
    const primerDato = 4; // Fila de inicio
    const numFilas = hoja.getLastRow() - primerDato + 1;
    const datos = hoja.getRange(primerDato, 2, numFilas > 0 ? numFilas : 0, 6).getValues();

    let fila = datos.findIndex(row => row.every(v => v === "")); // Índice dentro del rango
    if (fila !== -1) {
      fila = fila + primerDato; // Ajustar a número de fila real
    } else {
      fila = hoja.getLastRow() + 1; // Siguiente fila vacía después de la última con datos
    }

    editarCelda(fila, 2, '-'); // FECHA RETIRO (B)
    editarCelda(fila, 3, numeroLegajo); // NUMERO LPU (C)
    editarCelda(fila, 4, '-'); // CRED RETIRA (D)
    editarCelda(fila, 5, '-'); // DIVISION (E)
    editarCelda(fila, 6, credencialEntrada); // CRED ENTRADA (F)
    editarCelda(fila, 7, formatearFecha(fechaEntrada)); // FECHA DE ENTRADA (G)

    scrollToFila(fila);

    return {
      success: true,
      message: `No había salidas pendientes. Se registró la entrada como nueva fila (sin salida previa) para el legajo ${numeroLegajo}.`,
      filasActualizadas: 1,
    };
  }
    
    // Registrar entrada en todas las filas pendientes
    filasSinEntrada.forEach(fila => {
      editarCelda(fila, 6, credencialEntrada); // Columna F - CRED ENTRADA
      editarCelda(fila, 7, formatearFecha(fechaEntrada)); // Columna G - FECHA DE ENTRADA
    });
    
    scrollToFila(filasSinEntrada[filasSinEntrada.length - 1]);    
    return {
      success: true,
      message: `Entrada registrada para ${filasSinEntrada.length} salida(s) del legajo ${numeroLegajo}`,
      filasActualizadas: filasSinEntrada.length
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error al guardar entrada: ' + error.message
    };
  }
}

// Función para buscar todas las filas con un legajo
function buscarTodasLasFilasLegajo(legajo, columna) {
  const hoja = SpreadsheetApp.getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  const filas = [];
  
  if (ultimaFila < 4) return filas;
  
  const rango = hoja.getRange(4, columna, ultimaFila - 3, 1);
  const valores = rango.getValues();
  
  valores.forEach((fila, indice) => {
    if (fila[0] == legajo) {
      filas.push(indice + 4); // +4 porque empezamos en fila 4
    }
  });
  
  return filas;
}
//************ */
// Nueva función para encontrar fila sin entrada
function encontrarFilaSinEntrada(filas) {
  const hoja = SpreadsheetApp.getActiveSheet();
  
  // Recorrer filas de más reciente a más antigua (asumiendo que nuevas filas se añaden abajo)
  for (let i = filas.length - 1; i >= 0; i--) {
    const fila = filas[i];
    const credencial = hoja.getRange(fila, 6).getValue(); // Columna F
    const fecha = hoja.getRange(fila, 7).getValue();      // Columna G
    
    if (!credencial && !fecha) {
      return fila; // Devolver primera fila sin entrada
    }
  }
  
  return null; // Todas tienen entrada
}
// ==========================================
// FUNCIÓN GUARDAR SALIDA
// ==========================================
function guardarSalida(fechaSalida, numeroLegajo, division, credencialSalida) {
  try {
    const hoja = SpreadsheetApp.getActiveSheet();
    
    if (!hoja) {
      throw new Error('No se pudo acceder a la hoja de cálculo activa');
    }

    // Validar parámetros obligatorios
    if (!fechaSalida || !numeroLegajo || !division || !credencialSalida) {
      throw new Error('Todos los campos son obligatorios');
    }
    
    // Validar formato de legajo (5-6 dígitos)
    if (!/^\d{5,6}$/.test(numeroLegajo.toString())) {
      throw new Error('El número de legajo debe tener entre 5 y 6 dígitos');
    }
    
    // Validar formato de credencial (5 dígitos)
    if (!/^\d{5}$/.test(credencialSalida.toString())) {
      throw new Error('La credencial debe tener exactamente 5 dígitos');
    }
    
      // Verificar duplicados/modificaciones
    const resultadoVerificacion = verificarSalidaDuplicada(
      hoja, numeroLegajo, fechaSalida, credencialSalida, division
    );

    if (resultadoVerificacion) {
      if (resultadoVerificacion.tipo === "igual") {
        return {
          success: false,
          message: `Ya existe una salida registrada para el legajo ${numeroLegajo} con estos mismos datos en la fecha ${formatearFecha(fechaSalida)}.`,
          notification: 'info'
        };
      } else if (resultadoVerificacion.tipo === "diferente") {
        return {
          success: false,
          modal: true,
          message: "Ya existe una salida para este legajo en la fecha seleccionada.",
          datosExistentes: resultadoVerificacion.datosExistentes,
          datosNuevos: {
            fechaSalida: formatearFecha(fechaSalida),
            numeroLegajo: numeroLegajo,
            credencialSalida: credencialSalida,
            division: division
          },
          filaCoincidente: resultadoVerificacion.fila
        };
      }
    }

    // Marcar con '-' las entradas pendientes de este legajo (CRED ENTRADA y FECHA DE ENTRADA)
    const lastRow = hoja.getLastRow();
    if (lastRow >= 4) {
      const rangoLegajos = hoja.getRange(4, 3, lastRow - 3, 1).getValues(); // Columna C (NUMERO LPU)
      const rangoCredEntrada = hoja.getRange(4, 6, lastRow - 3, 1).getValues(); // Columna F (CRED ENTRADA)
      const rangoFechaEntrada = hoja.getRange(4, 7, lastRow - 3, 1).getValues(); // Columna G (FECHA DE ENTRADA)
      for (let i = 0; i < rangoLegajos.length; i++) {
        if (rangoLegajos[i][0].toString() === numeroLegajo.toString()) {
          // Si la celda de entrada está vacía, poner '-'
          if (!rangoCredEntrada[i][0] || rangoCredEntrada[i][0] === "") {
            editarCelda(i + 4, 6, "-"); // Col F
          }
          // Si la celda de fecha de entrada está vacía, poner '-'
          if (!rangoFechaEntrada[i][0] || rangoFechaEntrada[i][0] === "") {
            editarCelda(i + 4, 7, "-"); // Col G
          }
        }
      }
    }

    // Encontrar la próxima fila vacía desde B4 en adelante
    const fila = encontrarProximaFilaVacia();
    
    // Insertar datos en las columnas B, C, D, E
    editarCelda(fila, 2, formatearFecha(fechaSalida)); // Columna B - FECHA RETIRO
    editarCelda(fila, 3, numeroLegajo); // Columna C - NUMERO LPU
    editarCelda(fila, 4, credencialSalida); // Columna D - CRED RETIRA
    editarCelda(fila, 5, division); // Columna E - DIVISION
    
    scrollToFila(fila);    
    return {
      success: true,
      message: `Salida registrada correctamente para el legajo ${numeroLegajo} en la fila ${fila}`
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error al guardar salida: ' + error.message
    };
  }
}

// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

/** Busca si ya existe una salida para el legajo y fecha dados */
function verificarSalidaDuplicada(hoja, numeroLegajo, fechaSalida, credencialSalida, division) {
  const lastRow = hoja.getLastRow();
  if (lastRow < 4) return null;

  const rangoLegajos = hoja.getRange(4, 3, lastRow - 3, 1).getValues(); // Col C
  const rangoFechas = hoja.getRange(4, 2, lastRow - 3, 1).getValues();  // Col B
  const rangoCred = hoja.getRange(4, 4, lastRow - 3, 1).getValues();    // Col D
  const rangoDiv = hoja.getRange(4, 5, lastRow - 3, 1).getValues();     // Col E

  for (let i = 0; i < rangoLegajos.length; i++) {
    if (
      rangoLegajos[i][0].toString() === numeroLegajo.toString() &&
      formatearFecha(rangoFechas[i][0]) === formatearFecha(fechaSalida)
    ) {
      const datosExistentes = {
        fechaSalida: formatearFecha(rangoFechas[i][0]),
        numeroLegajo: rangoLegajos[i][0],
        credencialSalida: rangoCred[i][0],
        division: rangoDiv[i][0]
      };
      const iguales =
        datosExistentes.fechaSalida === formatearFecha(fechaSalida) &&
        datosExistentes.numeroLegajo.toString() === numeroLegajo.toString() &&
        datosExistentes.credencialSalida.toString() === credencialSalida.toString() &&
        datosExistentes.division.toString() === division.toString();
      return {
        tipo: iguales ? "igual" : "diferente",
        fila: i + 4,
        datosExistentes: datosExistentes
      };
    }
  }
  return null;
}


// Función para editar desde el formulario (ya la tienes)
function editarCelda(fila, columna, valor) {
  SpreadsheetApp.getActiveSheet()
    .getRange(fila, columna)
    .setValue(valor);
}

// Función para encontrar la próxima fila vacía desde B4 en adelante
function encontrarProximaFilaVacia() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const rangoB = hoja.getRange('B4:B'); // Desde B4 hasta el final
  const valores = rangoB.getValues();
  
  // Buscar la primera celda vacía
  for (let i = 0; i < valores.length; i++) {
    if (valores[i][0] === '' || valores[i][0] == null) {
      return i + 4; // +4 porque empezamos desde la fila 4
    }
  }
  
  // Si no hay celdas vacías, devolver la siguiente fila después de los datos
  return hoja.getLastRow() + 1;
}

// Función para buscar un legajo en una columna específica
function buscarLegajoEnColumna(numeroLegajo, columna) {
  const hoja = SpreadsheetApp.getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  
  if (ultimaFila < 4) return -1; // No hay datos desde la fila 4
  
  // Obtener valores desde la fila 4 hasta la última fila
  const rango = hoja.getRange(4, columna, ultimaFila - 3, 1);
  const valores = rango.getValues();
  
  // Buscar el legajo
  for (let i = 0; i < valores.length; i++) {
    if (valores[i][0] == numeroLegajo) {
      return i + 4; // +4 porque empezamos desde la fila 4
    }
  }
  
  return -1; // No encontrado
}

// Función para formatear fecha a dd/mm/yyyy
function formatearFecha(fecha) {
  console.log("FECHA: ", fecha);
  // Si viene como string del formulario HTML (yyyy-mm-dd)
  if (typeof fecha === 'string') {
    const partes = fecha.split('-');
    return `${partes[2]}/${partes[1]}/${partes[0]}`;
  }
  
  // Si viene como objeto Date
  if (fecha instanceof Date) {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  
  return fecha;
}

 /** * * * * * * * * * * * * * * * **/
  function scrollToFila(fila) {
     let hoja = SpreadsheetApp.getActiveSheet();
     hoja.setActiveSelection('B' + fila + ':G' + fila);
   }

function generarHtmlModalReemplazo(datosExistentes, datosNuevos, filaCoincidente) {
  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 24px; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 18px; }
          th, td { border: 1px solid #ccc; padding: 8px 12px; text-align: left; }
          th { background: #f7f7f7; }
          .diff { background: #ffe082; }
          .btns { text-align: right; }
          button {
            margin-left: 10px;
            padding: 8px 18px;
            border-radius: 4px;
            border: none;
            font-size: 1em;
            cursor: pointer;
          }
          .reemplazar { background: #388e3c; color: white; }
          .cancelar { background: #bdbdbd; color: #333; }
        </style>
      </head>
      <body>
        <h2>Ya existe una salida para este legajo y fecha</h2>
        <p>¿Desea reemplazar los datos existentes?</p>
        <table>
          <tr>
            <th>Campo</th>
            <th>Actual</th>
            <th>Nuevo</th>
          </tr>
          ${Object.keys(datosNuevos).map(campo => `
            <tr>
              <td>${campo}</td>
              <td${datosExistentes[campo] !== datosNuevos[campo] ? ' class="diff"' : ''}>${datosExistentes[campo]}</td>
              <td${datosExistentes[campo] !== datosNuevos[campo] ? ' class="diff"' : ''}>${datosNuevos[campo]}</td>
            </tr>
          `).join('')}
        </table>
        <div class="btns">
          <button class="reemplazar" onclick="reemplazarCarga()">Reemplazar carga</button>
          <button class="cancelar" onclick="google.script.host.close()">Cancelar</button>
        </div>
        <script>
          function reemplazarCarga() {
            google.script.run
              .withSuccessHandler(function(res) {
                if(res.success){
                  google.script.host.close();
                  google.script.run.showNotification('✅ ' + res.message, 'success');
                } else {
                  alert(res.message);
                }
              })
              .reemplazarDatosSalida(${filaCoincidente}, ${JSON.stringify(datosNuevos)});
          }
        </script>
      </body>
    </html>
  `;
}

  /** * * * * * * * * * * * * * * * **/
  function mostrarModalReemplazo(datosExistentes, datosNuevos, filaCoincidente) {
    let html = generarHtmlModalReemplazo(datosExistentes, datosNuevos, filaCoincidente);
    let modal = HtmlService.createHtmlOutput(html)
      .setWidth(600)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(modal, '¿Reemplazar carga existente?');
  }

  function reemplazarDatosSalida(fila, datosNuevos) {
    try {
      const hoja = SpreadsheetApp.getActiveSheet();
      hoja.getRange(fila, 2).setValue(datosNuevos.fechaSalida);      // Columna B
      hoja.getRange(fila, 3).setValue(datosNuevos.numeroLegajo);     // Columna C
      hoja.getRange(fila, 4).setValue(datosNuevos.credencialSalida); // Columna D
      hoja.getRange(fila, 5).setValue(datosNuevos.division);         // Columna E
      return {
        success: true,
        message: 'Datos reemplazados correctamente en la fila ' + fila
      };
    } catch (e) {
      return { success: false, message: 'Error al reemplazar datos: ' + e.message };
    }
  }

  /** * * * * * * * * * * * * * * * **/
  function verificarEntradaDuplicada(hoja, numeroLegajo, fechaEntrada, credencialEntrada) {
    const lastRow = hoja.getLastRow();
    if (lastRow < 4) return null;

    const rangoLegajos = hoja.getRange(4, 3, lastRow - 3, 1).getValues(); // Col C
    const rangoFechas = hoja.getRange(4, 7, lastRow - 3, 1).getValues();  // Col G
    const rangoCred = hoja.getRange(4, 6, lastRow - 3, 1).getValues();    // Col F

    for (let i = 0; i < rangoLegajos.length; i++) {
      if (
        rangoLegajos[i][0].toString() === numeroLegajo.toString() &&
        formatearFecha(rangoFechas[i][0]) === formatearFecha(fechaEntrada)
      ) {
        const datosExistentes = {
          fechaEntrada: formatearFecha(rangoFechas[i][0]),
          numeroLegajo: rangoLegajos[i][0],
          credencialEntrada: rangoCred[i][0]
        };
        const iguales =
          datosExistentes.fechaEntrada === formatearFecha(fechaEntrada) &&
          datosExistentes.numeroLegajo.toString() === numeroLegajo.toString() &&
          datosExistentes.credencialEntrada.toString() === credencialEntrada.toString();
        return {
          tipo: iguales ? "igual" : "diferente",
          fila: i + 4,
          datosExistentes: datosExistentes
        };
      }
    }
    return null;
  }

function generarHtmlModalReemplazoEntrada(datosExistentes, datosNuevos, filaCoincidente) {
  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 24px; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 18px; }
          th, td { border: 1px solid #ccc; padding: 8px 12px; text-align: left; }
          th { background: #f7f7f7; }
          .diff { background: #ffe082; }
          .btns { text-align: right; }
          button {
            margin-left: 10px;
            padding: 8px 18px;
            border-radius: 4px;
            border: none;
            font-size: 1em;
            cursor: pointer;
          }
          .reemplazar { background: #388e3c; color: white; }
          .cancelar { background: #bdbdbd; color: #333; }
        </style>
      </head>
      <body>
        <h2>Ya existe una entrada para este legajo y fecha</h2>
        <p>¿Desea reemplazar los datos existentes de la entrada?</p>
        <table>
          <tr>
            <th>Campo</th>
            <th>Actual</th>
            <th>Nuevo</th>
          </tr>
          ${Object.keys(datosNuevos).map(campo => `
            <tr>
              <td>${campo}</td>
              <td${datosExistentes[campo] !== datosNuevos[campo] ? ' class="diff"' : ''}>${datosExistentes[campo]}</td>
              <td${datosExistentes[campo] !== datosNuevos[campo] ? ' class="diff"' : ''}>${datosNuevos[campo]}</td>
            </tr>
          `).join('')}
        </table>
        <div class="btns">
          <button class="reemplazar" onclick="reemplazarCarga()">Reemplazar entrada</button>
          <button class="cancelar" onclick="google.script.host.close()">Cancelar</button>
        </div>
        <script>
          function reemplazarCarga() {
            google.script.run
              .withSuccessHandler(function(res) {
                if(res.success){
                  google.script.host.close();
                  google.script.run.showNotification('✅ ' + res.message, 'success');
                } else {
                  alert(res.message);
                }
              })
              .reemplazarDatosEntrada(${filaCoincidente}, ${JSON.stringify(datosNuevos)});
          }
        </script>
      </body>
    </html>
  `;
}

function mostrarModalReemplazoEntrada(datosExistentes, datosNuevos, filaCoincidente) {
  let html = generarHtmlModalReemplazoEntrada(datosExistentes, datosNuevos, filaCoincidente);
  let modal = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(modal, '¿Reemplazar entrada existente?');
}

function reemplazarDatosEntrada(fila, datosNuevos) {
  try {
    const hoja = SpreadsheetApp.getActiveSheet();
    hoja.getRange(fila, 3).setValue(datosNuevos.numeroLegajo);     // Columna C
    hoja.getRange(fila, 6).setValue(datosNuevos.credencialEntrada); // Columna F
    hoja.getRange(fila, 7).setValue(datosNuevos.fechaEntrada);     // Columna G
    return {
      success: true,
      message: 'Datos de entrada reemplazados correctamente en la fila ' + fila
    };
  } catch (e) {
    return { success: false, message: 'Error al reemplazar datos de entrada: ' + e.message };
  }
}
