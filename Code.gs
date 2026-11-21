const SC = {
  OWNER: '',
  TZ: 'Europe/Madrid',
  SHEETS: {
    CLIENTAS: ['ID clienta', 'Nombre y apellidos', 'Teléfono', 'Email', 'Fecha de alta', 'Notas'],
    SERVICIOS: ['ID servicio', 'Categoría', 'Servicio', 'Duración (min)', 'Precio visible', 'Activo'],
    RESERVAS: ['ID reserva', 'Fecha', 'Hora inicio', 'Hora fin', 'ID servicio', 'Servicio (copia)', 'Duración reservada (min)', 'Precio visible (copia)', 'ID clienta', 'Nombre', 'Teléfono', 'Email', 'Estado', 'Origen', 'Creada el'],
    HORARIOS: ['Día', 'Abierto', 'Apertura', 'Cierre'],
    BLOQUEOS: ['ID bloqueo', 'Fecha', 'Hora inicio', 'Hora fin', 'Motivo', 'Activo'],
    CONFIGURACION: ['Clave', 'Valor', 'Descripción']
  }
};

// The workbook already in use predates this API and uses title-case sheet names
// and a slightly different set of column headers. Resolve both schemas so the
// integration reads and writes the existing workbook in place.
const SHEET_ALIASES = {
  CLIENTAS: ['CLIENTAS', 'Clientas', 'Clientes'],
  SERVICIOS: ['SERVICIOS', 'Servicios'],
  RESERVAS: ['RESERVAS', 'Reservas'],
  HORARIOS: ['HORARIOS', 'Horarios'],
  BLOQUEOS: ['BLOQUEOS', 'Bloqueos'],
  CONFIGURACION: ['CONFIGURACION', 'Configuracion', 'Configuración']
};

const SC_SERVICES = [
  ['mani-rusa', 'Manicura', 'Manicura rusa sin esmalte', 30, '20 €', 'Sí'],
  ['mani-trad', 'Manicura', 'Manicura + esmalte tradicional', 45, '25 €', 'Sí'],
  ['nude', 'Manicura', 'Manicura + Semi Nude', 75, '35 €', 'Sí'],
  ['color', 'Manicura', 'Manicura + Semi con color', 90, '40 €', 'Sí'],
  ['french', 'Diseños', 'Francesa / Pigmentos / Babyboomer', 15, '5 €', 'Sí'],
  ['design-extra', 'Diseños', 'Diseños complejos y otros', 15, 'Desde 5 €', 'Sí'],
  ['gel-s', 'Relleno de gel', 'Relleno de gel · Cortas S', 120, 'Desde 45 €', 'Sí'],
  ['gel-m', 'Relleno de gel', 'Relleno de gel · Medianas M', 120, 'Desde 50 €', 'Sí'],
  ['gel-l', 'Relleno de gel', 'Relleno de gel · Largas L', 150, 'Desde 55 €', 'Sí'],
  ['ext-s', 'Extensiones', 'Extensiones · Cortas S', 120, 'Desde 50 €', 'Sí'],
  ['ext-m', 'Extensiones', 'Extensiones · Medianas M', 150, 'Desde 60 €', 'Sí'],
  ['ext-l', 'Extensiones', 'Extensiones · Largas L', 150, 'Desde 70 €', 'Sí'],
  ['ext-xl', 'Extensiones', 'Extensiones · Extra largas XL', 180, 'Desde 80 €', 'Sí'],
  ['pedi-lam', 'Pedicura', 'Pedicura completa con láminas', 120, '60 €', 'Sí'],
  ['pedi-clean', 'Pedicura', 'Pedicura completa sin esmaltado', 60, '40 €', 'Sí'],
  ['pedi-semi', 'Pedicura', 'Pedicura completa + semipermanente', 90, '50 €', 'Sí'],
  ['pedi-toes', 'Pedicura', 'Pedicura solo dedos + semipermanente', 60, '40 €', 'Sí'],
  ['pedi-clean-toes', 'Pedicura', 'Solo limpieza de dedos', 45, '30 €', 'Sí'],
  ['pedi-heel', 'Pedicura', 'Solo limpieza de talón', 30, '20 €', 'Sí'],
  ['remove-semi', 'Otros servicios', 'Retirada de semi completa sin volver a esmaltar', 30, '10 €', 'Sí'],
  ['remove-gel', 'Otros servicios', 'Retirada de gel completa sin volver a esmaltar', 40, '15 €', 'Sí']
];

function setupSabinaCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abre Apps Script desde la hoja de cálculo de Sabina.');
  ss.setSpreadsheetTimeZone(SC.TZ);
  Object.keys(SC.SHEETS).forEach(name => {
    const sheet = sheet_(name) || ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, SC.SHEETS[name].length).setValues([SC.SHEETS[name]]);
      sheet.setFrozenRows(1);
    }
  });
  const services = sheet_('SERVICIOS');
  if (services.getLastRow() < 2) services.getRange(2, 1, SC_SERVICES.length, 6).setValues(SC_SERVICES);
  const hours = sheet_('HORARIOS');
  if (hours.getLastRow() < 2) {
    hours.getRange(2, 1, 7, 4).setValues([
      ['Lunes', 'Sí', '09:30', '21:30'],
      ['Martes', 'Sí', '09:30', '21:30'],
      ['Miércoles', 'Sí', '09:30', '21:30'],
      ['Jueves', 'Sí', '09:30', '21:30'],
      ['Viernes', 'Sí', '09:30', '21:30'],
      ['Sábado', 'No', '', ''],
      ['Domingo', 'No', '', '']
    ]);
  }
  const config = sheet_('CONFIGURACION');
  if (config.getLastRow() < 2) {
    config.getRange(2, 1, 5, 3).setValues([
      ['email_propietario', ownerEmail_(), 'Cuenta propietaria'],
      ['permitir_nuevas_clientas', 'Sí', 'Permitir primeras reservas online'],
      ['margen_entre_citas_min', '0', 'Margen entre citas'],
      ['optimizacion_agenda', 'Destacar mejores huecos', 'Ordenar horarios recomendados'],
      ['zona_horaria', SC.TZ, 'Zona horaria']
    ]);
  }
  PropertiesService.getScriptProperties().setProperty('SABINA_SHEET_ID', ss.getId());
  return {ok: true, spreadsheetId: ss.getId(), sheets: Object.keys(SC.SHEETS)};
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  try {
    let result;
    if (params.action === 'catalog') result = {ok: true, services: catalog_()};
    else if (params.action === 'availability') result = availability_(params.date, params.serviceId);
    else if (params.action === 'health') {
      ss_();
      result = {ok: true, service: 'Sabina Calendar API', state: 'READY', timezone: SC.TZ};
    } else {
      result = {ok: true, service: 'Sabina Calendar API', state: 'READY'};
    }
    return output_(result, params.callback);
  } catch (error) {
    return output_({ok: false, error: errorMessage_(error)}, params.callback);
  }
}

function ownerEmail_() {
  return String(Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '').trim();
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (error) {
    return output_({ok: false, error: 'Solicitud no válida.'});
  }
  try {
    const action = String(body.action || '');
    if (action === 'createBooking') return output_(createBooking_(body));
    if (action === 'findBookings') return output_(findBookings_(body));
    if (action === 'updateBooking') return output_(updateBooking_(body));
    if (action === 'cancelBooking') return output_(cancelBooking_(body));
    return output_({ok: false, error: 'Acción no reconocida.'});
  } catch (error) {
    return output_({ok: false, error: errorMessage_(error)});
  }
}

function createBooking_(request) {
  return runLocked_(function() {
    const input = validateBooking_(request);
    const client = findClient_(input.phone, input.email);
    const config = config_();
    if (!client && !booleanish_(configValue_(config, ['permitir_nuevas_clientas', 'permitirNuevas'], 'Sí'), true)) {
      throw new Error('Las primeras citas deben solicitarse directamente a Sabina.');
    }
    assertSlotFree_(input.date, input.start, input.end, '');
    const clientId = client ? String(client.id) : Utilities.getUuid();
    const ss = ss_();
    if (!client) {
      const clientSheet = sheet_('CLIENTAS');
      const headers = clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).getValues()[0];
      clientSheet.appendRow(clientRowValues_({id: clientId, name: input.name, phone: input.phone, email: input.email, createdAt: new Date()}, headers));
    }
    const booking = {
      id: Utilities.getUuid(),
      date: input.date,
      start: fromMin_(input.start),
      end: fromMin_(input.end),
      serviceId: input.service.id,
      service: input.service.name,
      duration: input.service.duration,
      price: input.service.price,
      clientId: clientId,
      name: input.name,
      phone: input.phone,
      email: input.email,
      status: 'CONFIRMADA',
      origin: String(request.origin || 'WEB'),
      createdAt: new Date()
    };
    writeBooking_(booking);
    return {ok: true, booking: bookingPublic_(booking)};
  });
}

function findBookings_(request) {
  const phone = normalizePhone_(request.phone);
  const email = normalizeEmail_(request.email);
  if (!phone || !isEmail_(email)) throw new Error('Introduce el teléfono y el correo usados en la reserva.');
  const bookings = bookingRecords_()
    .filter(record => isActiveStatus_(record.status) && record.phoneKey === phone && record.email === email)
    .sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start))
    .map(bookingPublic_);
  return {ok: true, bookings: bookings};
}

function updateBooking_(request) {
  return runLocked_(function() {
    const current = findManagedBooking_(request);
    assertManageable_(current);
    const input = validateBooking_(request);
    assertSlotFree_(input.date, input.start, input.end, current.id);
    const updated = {
      id: current.id,
      date: input.date,
      start: fromMin_(input.start),
      end: fromMin_(input.end),
      serviceId: input.service.id,
      service: input.service.name,
      duration: input.service.duration,
      price: input.service.price,
      clientId: current.clientId,
      name: input.name,
      phone: input.phone,
      email: input.email,
      status: 'CONFIRMADA',
      origin: current.origin || 'WEB',
      createdAt: current.createdAt || new Date()
    };
    writeBooking_(updated, current.row);
    return {ok: true, booking: bookingPublic_(updated)};
  });
}

function cancelBooking_(request) {
  return runLocked_(function() {
    const current = findManagedBooking_(request);
    assertManageable_(current);
    const sheet = sheet_('RESERVAS');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColumn = fieldIndex_(headers, ['estado', 'status'], 12) + 1;
    sheet.getRange(current.row, statusColumn).setValue('CANCELADA');
    current.status = 'CANCELADA';
    return {ok: true, booking: bookingPublic_(current)};
  });
}

function availability_(dateValue, serviceId) {
  const date = String(dateValue || '');
  validateDate_(date);
  const context = requestContext_();
  const service = serviceById_(serviceId, context);
  const hours = opening_(date, context);
  if (!hours) return {ok: true, date: date, service: servicePublic_(service), slots: []};
  // Load the live Sheets data once per availability request. The old code
  // re-read RESERVAS and BLOQUEOS for every 15-minute candidate slot.
  const bookings = bookingRecords_(context);
  const blocks = table_('BLOQUEOS', context);
  const slots = [];
  for (let start = hours.start; start + service.duration <= hours.end; start += 15) {
    if (!slotIsInPast_(date, start) && !hasConflict_(date, start, start + service.duration, '', context, bookings, blocks)) {
      slots.push(fromMin_(start));
    }
  }
  return {
    ok: true,
    date: date,
    service: servicePublic_(service),
    opening: {start: fromMin_(hours.start), end: fromMin_(hours.end)},
    slots: slots
  };
}

function validateBooking_(request) {
  const date = String(request.date || '');
  const start = toMin_(request.start);
  const service = serviceById_(request.serviceId);
  const name = String(request.name || '').trim();
  const phone = String(request.phone || '').trim();
  const email = normalizeEmail_(request.email);
  validateDate_(date);
  if (!Number.isInteger(start) || start % 15 !== 0) throw new Error('El horario seleccionado no es válido.');
  if (!name || !phone || !isEmail_(email)) throw new Error('Completa nombre, teléfono y correo electrónico.');
  if (slotIsInPast_(date, start)) throw new Error('Ese horario ya ha pasado. Elige otro.');
  const hours = opening_(date);
  const end = start + service.duration;
  if (!hours || start < hours.start || end > hours.end) throw new Error('Horario fuera de apertura.');
  return {date: date, start: start, end: end, service: service, name: name, phone: phone, email: email};
}

function findManagedBooking_(request) {
  const bookingId = String(request.bookingId || '');
  const phone = normalizePhone_(request.authPhone || request.phone);
  const email = normalizeEmail_(request.authEmail || request.email);
  if (!bookingId || !phone || !isEmail_(email)) throw new Error('No se han podido validar los datos de la reserva.');
  const booking = bookingRecords_().find(record => record.id === bookingId && record.phoneKey === phone && record.email === email && isActiveStatus_(record.status));
  if (!booking) throw new Error('No encontramos una reserva con esos datos.');
  return booking;
}

function assertManageable_(booking) {
  if (appointmentDate_(booking.date, booking.start).getTime() - new Date().getTime() < 12 * 60 * 60 * 1000) {
    throw new Error('El plazo para modificar o cancelar esta cita es de 12 horas antes.');
  }
}

function assertSlotFree_(date, start, end, excludedId) {
  if (hasConflict_(date, start, end, excludedId)) throw new Error('Ese horario acaba de ocuparse. Elige otro.');
}

function hasConflict_(date, start, end, excludedId, context, bookings, blocks) {
  const margin = Number(configValue_(config_(context), ['margen_entre_citas_min', 'margen', 'intervalo'], 0)) || 0;
  const bookingConflict = (bookings || bookingRecords_(context)).some(record => {
    if (record.id === excludedId || record.date !== date || !isActiveStatus_(record.status)) return false;
    return start < record.endMin + margin && end > record.startMin - margin;
  });
  if (bookingConflict) return true;
  return (blocks || table_('BLOQUEOS', context)).some(record => {
    if (dateKey_(recordField_(record, ['fecha', 'date'])) !== date) return false;
    const active = recordField_(record, ['activo', 'active'], true);
    if (!booleanish_(active, true)) return false;
    return start < toMin_(recordField_(record, ['hora fin', 'fin', 'end'])) && end > toMin_(recordField_(record, ['hora inicio', 'inicio', 'start']));
  });
}

function bookingRecords_(context) {
  return table_('RESERVAS', context).map(record => {
    const start = timeText_(recordField_(record, ['hora inicio', 'inicio', 'start']));
    const end = timeText_(recordField_(record, ['hora fin', 'fin', 'end']));
    return {
      row: record.row,
      values: record.values,
      headers: record.headers,
      id: String(recordField_(record, ['id reserva', 'id']) || ''),
      date: dateKey_(recordField_(record, ['fecha', 'date'])),
      start: start,
      end: end,
      startMin: toMin_(start),
      endMin: toMin_(end),
      serviceId: String(recordField_(record, ['id servicio', 'servicioid', 'serviceid']) || ''),
      service: String(recordField_(record, ['servicio copia', 'servicio']) || ''),
      duration: Number(recordField_(record, ['duracion reservada min', 'duracion']) || 0),
      price: String(recordField_(record, ['precio visible copia', 'precio']) || ''),
      clientId: String(recordField_(record, ['id clienta', 'clientaid']) || ''),
      name: String(recordField_(record, ['nombre y apellidos', 'nombre', 'name']) || ''),
      phone: String(recordField_(record, ['telefono', 'phone']) || ''),
      phoneKey: normalizePhone_(recordField_(record, ['telefono', 'phone'])),
      email: normalizeEmail_(recordField_(record, ['email', 'correo'])),
      status: String(recordField_(record, ['estado', 'status']) || '').toUpperCase(),
      origin: String(recordField_(record, ['origen', 'origin']) || ''),
      createdAt: recordField_(record, ['creada el', 'creada', 'created'])
    };
  }).filter(record => record.id && record.date && record.start && record.end);
}

function writeBooking_(booking, row) {
  const sheet = sheet_('RESERVAS');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = [bookingRowValues_(booking, headers)];
  if (row) sheet.getRange(row, 1, 1, values[0].length).setValues(values);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, values[0].length).setValues(values);
}

function bookingPublic_(booking) {
  return {
    bookingId: booking.id,
    name: booking.name,
    date: booking.date,
    start: booking.start,
    end: booking.end,
    serviceId: booking.serviceId,
    service: booking.service,
    duration: booking.duration,
    price: booking.price,
    status: booking.status,
    canManage: isActiveStatus_(booking.status) && appointmentDate_(booking.date, booking.start).getTime() - new Date().getTime() >= 12 * 60 * 60 * 1000
  };
}

function catalog_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('sabina.catalog.v1');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      cache.remove('sabina.catalog.v1');
    }
  }
  const services = table_('SERVICIOS').map(serviceFromRecord_).filter(service => service.active);
  cache.put('sabina.catalog.v1', JSON.stringify(services), 60);
  return services;
}

function serviceById_(serviceId, context) {
  const id = String(serviceId || '');
  const service = catalog_(context).find(item => item.id === id);
  if (!service) throw new Error('El servicio ya no está disponible.');
  return service;
}

function serviceFromRecord_(record) {
  return {
    id: String(recordField_(record, ['id servicio', 'id']) || ''),
    category: String(recordField_(record, ['categoria', 'categoría']) || ''),
    name: String(recordField_(record, ['servicio', 'nombre']) || ''),
    duration: Number(recordField_(record, ['duracion min', 'duracion', 'duración']) || 0),
    price: displayPrice_(recordField_(record, ['precio visible', 'precio'])),
    active: booleanish_(recordField_(record, ['activo', 'active']), false)
  };
}

function servicePublic_(service) {
  return {id: service.id, category: service.category, name: service.name, duration: service.duration, price: service.price};
}

function displayPrice_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  return /(?:€|eur)/i.test(text) ? text : text + ' €';
}

function findClient_(phone, email) {
  return table_('CLIENTAS').find(record => {
    return normalizePhone_(recordField_(record, ['telefono', 'phone'])) === normalizePhone_(phone) || normalizeEmail_(recordField_(record, ['email', 'correo'])) === normalizeEmail_(email);
  }) || null;
}

function opening_(date, context) {
  const dayIndex = new Date(date + 'T12:00:00').getDay();
  const weekday = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'][dayIndex];
  const record = table_('HORARIOS', context).find(item => {
    const day = recordField_(item, ['dia', 'día', 'day']);
    const numericDay = Number(day);
    const matchesDay = Number.isFinite(numericDay) ? (numericDay === (dayIndex === 0 ? 7 : dayIndex)) : normalizeText_(day) === normalizeText_(weekday);
    return matchesDay && booleanish_(recordField_(item, ['activo', 'active', 'abierto', 'open']), false);
  });
  if (!record) return null;
  return {start: toMin_(recordField_(record, ['horainicio', 'hora inicio', 'apertura', 'start'])), end: toMin_(recordField_(record, ['horafin', 'hora fin', 'cierre', 'end']))};
}

function config_(context) {
  const result = {};
  table_('CONFIGURACION', context).forEach(record => {
    const key = String(recordField_(record, ['clave', 'key']) || '');
    result[key] = String(recordField_(record, ['valor', 'value']) || '');
  });
  return result;
}

function requestContext_() {
  return {tables: {}};
}

function table_(name, context) {
  if (context && Object.prototype.hasOwnProperty.call(context.tables, name)) return context.tables[name];
  const sheet = sheet_(name);
  if (!sheet || sheet.getLastRow() < 2) {
    if (context) context.tables[name] = [];
    return [];
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const records = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().map((values, index) => ({values: values, headers: headers, row: index + 2}));
  if (context) context.tables[name] = records;
  return records;
}

function sheet_(name) {
  const ss = ss_unsafe_();
  const candidates = SHEET_ALIASES[name] || [name];
  for (let i = 0; i < candidates.length; i++) {
    const sheet = ss.getSheetByName(candidates[i]);
    if (sheet) return sheet;
  }
  return null;
}

function ss_unsafe_() {
  const id = PropertiesService.getScriptProperties().getProperty('SABINA_SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function normalizeHeader_(value) {
  return normalizeText_(value).replace(/[^a-z0-9]+/g, '');
}

function normalizeText_(value) {
  return String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function fieldIndex_(headers, names, fallback) {
  const wanted = names.map(normalizeHeader_);
  for (let i = 0; i < headers.length; i++) if (wanted.includes(normalizeHeader_(headers[i]))) return i;
  return fallback;
}

function recordField_(record, names, defaultValue) {
  const index = fieldIndex_(record.headers || [], names, -1);
  return index >= 0 ? record.values[index] : defaultValue;
}

function booleanish_(value, defaultValue) {
  if (value === true) return true;
  if (value === false) return false;
  const text = normalizeText_(value);
  if (['si', 'true', '1', 'activo', 'abierto', 'yes'].includes(text)) return true;
  if (['no', 'false', '0', 'inactivo', 'cerrado'].includes(text)) return false;
  return defaultValue;
}

function configValue_(config, names, defaultValue) {
  const wanted = names.map(normalizeHeader_);
  const key = Object.keys(config).find(item => wanted.includes(normalizeHeader_(item)));
  return key === undefined ? defaultValue : config[key];
}

function clientRowValues_(client, headers) {
  return headers.map(header => {
    const key = normalizeHeader_(header);
    if (key === 'idclienta' || key === 'id') return client.id;
    if (key === 'nombreyapellidos' || key === 'nombre' || key === 'name') return client.name;
    if (key === 'telefono' || key === 'phone') return client.phone;
    if (key === 'email' || key === 'correo') return client.email;
    if (key === 'fechadealta' || key === 'fechaalta' || key === 'created') return client.createdAt;
    return '';
  });
}

function bookingRowValues_(booking, headers) {
  return headers.map(header => {
    const key = normalizeHeader_(header);
    if (key === 'idreserva' || key === 'id') return booking.id;
    if (key === 'fecha' || key === 'date') return booking.date;
    if (key === 'horainicio' || key === 'inicio' || key === 'start') return booking.start;
    if (key === 'horafin' || key === 'fin' || key === 'end') return booking.end;
    if (key === 'idservicio' || key === 'servicioid' || key === 'serviceid') return booking.serviceId;
    if (key === 'serviciocopia' || key === 'servicio' || key === 'service') return booking.service;
    if (key === 'duracionreservadamin' || key === 'duracion' || key === 'duration') return booking.duration;
    if (key === 'preciovisiblecopia' || key === 'precio' || key === 'price') return booking.price;
    if (key === 'idclienta' || key === 'clientaid') return booking.clientId;
    if (key === 'nombreyapellidos' || key === 'nombre' || key === 'name') return booking.name;
    if (key === 'telefono' || key === 'phone') return booking.phone;
    if (key === 'email' || key === 'correo') return booking.email;
    if (key === 'estado' || key === 'status') return booking.status;
    if (key === 'origen' || key === 'origin') return booking.origin;
    if (key === 'creadael' || key === 'creada' || key === 'created') return booking.createdAt;
    return '';
  });
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SABINA_SHEET_ID');
  if (!id) throw new Error('Ejecuta primero setupSabinaCalendar().');
  return SpreadsheetApp.openById(id);
}

function runLocked_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function validateDate_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Fecha no válida.');
  const parts = date.split('-').map(Number);
  const value = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (value.getUTCFullYear() !== parts[0] || value.getUTCMonth() !== parts[1] - 1 || value.getUTCDate() !== parts[2]) throw new Error('Fecha no válida.');
}

function slotIsInPast_(date, start) {
  const today = Utilities.formatDate(new Date(), SC.TZ, 'yyyy-MM-dd');
  if (date < today) return true;
  if (date > today) return false;
  const now = toMin_(Utilities.formatDate(new Date(), SC.TZ, 'HH:mm'));
  return start <= now;
}

function appointmentDate_(date, time) {
  return Utilities.parseDate(date + ' ' + time, SC.TZ, 'yyyy-MM-dd HH:mm');
}

function dateKey_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, SC.TZ, 'yyyy-MM-dd');
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? '' : Utilities.formatDate(parsed, SC.TZ, 'yyyy-MM-dd');
}

function timeText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, SC.TZ, 'HH:mm');
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return match ? fromMin_(Number(match[1]) * 60 + Number(match[2])) : '';
}

function toMin_(value) {
  if (value instanceof Date) return Number(Utilities.formatDate(value, SC.TZ, 'H')) * 60 + Number(Utilities.formatDate(value, SC.TZ, 'm'));
  if (typeof value === 'number' && isFinite(value)) return value >= 0 && value < 1 ? Math.round(value * 24 * 60) : Math.round(value);
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : NaN;
}

function fromMin_(minutes) {
  return String(Math.floor(minutes / 60)).padStart(2, '0') + ':' + String(minutes % 60).padStart(2, '0');
}

function normalizePhone_(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function isEmail_(value) {
  return /^\S+@\S+\.\S+$/.test(value);
}

function isActiveStatus_(status) {
  return !['CANCELADA', 'ANULADA'].includes(String(status || '').toUpperCase());
}

function errorMessage_(error) {
  return String(error && error.message ? error.message : error);
}

function output_(value, callback) {
  const json = JSON.stringify(value);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}


