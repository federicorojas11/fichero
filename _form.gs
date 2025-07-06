
// ==========================================
// FUNCIÓN CONSULTAR LEGAJO
// ==========================================
/*function consultarLegajo(numeroLegajo) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const datos = hoja.getDataRange().getValues();

    // Validación básica
    if (!numeroLegajo || !/^\d{5,6}$/.test(numeroLegajo.toString())) {
      return {
        success: false,
        message: 'Número de legajo inválido. Debe tener entre 5 y 6 dígitos.'
      };
    }

    const resultados = [];

    // Buscar filas coincidentes (columna C = índice 2)
    for (let i = 3; i < datos.length; i++) { // Empieza desde fila 4
      const celdaLegajo = datos[i][2]; // Columna C - Número LPU
      if (celdaLegajo == numeroLegajo) {
        resultados.push({
          fila: i + 1,
          fechaRetiro: datos[i][1] || '-', // B - Fecha Retiro
          credencialRetira: datos[i][3] || '-', // D - Credencial Salida
          division: datos[i][4] || '-',   // E - División
          credencialEntrada: datos[i][5] || '-', // F - Credencial Entrada
          fechaEntrada: datos[i][6] || '-' // G - Fecha Entrada
        });
      }
    }

    if (resultados.length === 0) {
      return {
        success: false,
        message: `No se encontró el legajo ${numeroLegajo}`
      };
    }

    // Ordenar por fecha más reciente
    resultados.sort((a, b) => {
      const fechaA = new Date(a.fechaRetiro.split('/').reverse().join('-'));
      const fechaB = new Date(b.fechaRetiro.split('/').reverse().join('-'));
      return fechaB - fechaA;
    });

    const ultimoRegistro = resultados[0];
    const estadoActual = ultimoRegistro.fechaEntrada !== '-' ? 'DEVUELTO' : 'EN USO';

    return {
      success: true,
      numeroLegajo,
      estadoActual,
      ultimaFechaRetiro: ultimoRegistro.fechaRetiro,
      ultimaDivision: ultimoRegistro.division,
      resultados
    };

  } catch (error) {
    console.log("Error en consultarLegajo:", error.message);
    return {
      success: false,
      message: "Error interno al consultar el legajo: " + error.message
    };
  }
}*/
function consultarLegajo(numeroLegajo) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const datos = hoja.getDataRange().getValues();

    // Validación básica del número de legajo
    if (!numeroLegajo || !/^\d{5,6}$/.test(numeroLegajo.toString())) {
      return {
        success: false,
        message: 'Número de legajo inválido. Debe tener entre 5 y 6 dígitos.'
      };
    }

    const resultados = [];

    // Buscar filas coincidentes (columna C = índice 2)
    for (let i = 3; i < datos.length; i++) { // Empieza desde fila 4
      const celdaLegajo = datos[i][2]; // Columna C - Número LPU
      if (celdaLegajo == numeroLegajo) {
        // Convertir fecha de objeto Date a cadena dd/MM/yyyy si es necesario
        const fechaRetiro = datos[i][1] ? formatearFechaCliente(datos[i][1]) : '-';
        const fechaEntrada = datos[i][6] ? formatearFechaCliente(datos[i][6]) : '-';

        resultados.push({
          fila: i + 1,
          fechaRetiro,
          credencialRetira: datos[i][3] || '-', // D - Credencial Salida
          division: datos[i][4] || '-',         // E - División
          credencialEntrada: datos[i][5] || '-',// F - Credencial Entrada
          fechaEntrada
        });
      }
    }

    if (resultados.length === 0) {
      return {
        success: false,
        message: `No se encontró ningún registro para el legajo ${numeroLegajo}`
      };
    }

    // Ordenar por fecha más reciente
    resultados.sort((a, b) => {
      const fechaA = new Date(a.fechaRetiro.split('/').reverse().join('-'));
      const fechaB = new Date(b.fechaRetiro.split('/').reverse().join('-'));
      return fechaB - fechaA;
    });

    const ultimoRegistro = resultados[0];
    const estadoActual = ultimoRegistro.fechaEntrada !== '-' ? 'DEVUELTO' : 'EN USO';

    return {
      success: true,
      numeroLegajo,
      estadoActual,
      ultimaFechaRetiro: ultimoRegistro.fechaRetiro,
      ultimaDivision: ultimoRegistro.division,
      resultados
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

// ==========================================
// FUNCIÓN GUARDAR ENTRADA
// ==========================================
function guardarEntrada(fechaEntrada, numeroLegajo, credencialEntrada) {
  try {
    const hoja = SpreadsheetApp.getActiveSheet();
    
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
    
    // Buscar TODAS las filas con este legajo
    const filasEncontradas = buscarTodasLasFilasLegajo(numeroLegajo, 3); // Columna C = 3
    
    if (filasEncontradas.length === 0) {
      throw new Error('No se encontró el legajo ' + numeroLegajo + ' en registros de salida');
    }
    
    // Filtrar filas sin entrada registrada
    const filasSinEntrada = filasEncontradas.filter(fila => {
      const credencial = hoja.getRange(fila, 6).getValue(); // Columna F
      const fecha = hoja.getRange(fila, 7).getValue();      // Columna G
      return !credencial && !fecha;
    });
    
    if (filasSinEntrada.length === 0) {
      throw new Error('Todas las salidas del legajo ' + numeroLegajo + ' ya tienen entrada registrada');
    }
    
    // Registrar entrada en todas las filas pendientes
    filasSinEntrada.forEach(fila => {
      editarCelda(fila, 6, credencialEntrada); // Columna F - CRED ENTRADA
      editarCelda(fila, 7, formatearFecha(fechaEntrada)); // Columna G - FECHA DE ENTRADA
    });
    
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
    
    // Encontrar la próxima fila vacía desde B4 en adelante
    const proximaFilaVacia = encontrarProximaFilaVacia();
    
    // Insertar datos en las columnas B, C, D, E
    editarCelda(proximaFilaVacia, 2, formatearFecha(fechaSalida)); // Columna B - FECHA RETIRO
    editarCelda(proximaFilaVacia, 3, numeroLegajo); // Columna C - NUMERO LPU
    editarCelda(proximaFilaVacia, 4, credencialSalida); // Columna D - CRED RETIRA
    editarCelda(proximaFilaVacia, 5, division); // Columna E - DIVISION
    
    return {
      success: true,
      message: `Salida registrada correctamente para el legajo ${numeroLegajo} en la fila ${proximaFilaVacia}`
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
  alert("fecha en función formatearFecha(fecha) del _form.gs: ", fecha);
  if (!fecha) return '';
  
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
