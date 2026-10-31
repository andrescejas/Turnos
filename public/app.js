let dashboardData = null;
let currentMonth = 8;

const NOMBRES_MESES = {
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
};

async function loadData(force = false) {
    try {
        const btn = document.getElementById('btn-refresh');
        if (force && btn) btn.classList.add('loading');

        const url = force ? '/api/data?force=true' : '/api/data';
        const res = await fetch(url);
        const json = await res.json();

        if (json.ok && json.data) {
            dashboardData = json.data;
            renderDashboard();
        }
    } catch (e) {
        console.error("Error al cargar datos:", e);
    } finally {
        const btn = document.getElementById('btn-refresh');
        if (btn) btn.classList.remove('loading');
    }
}

function forceRefresh() {
    loadData(true);
}

function renderDashboard() {
    if (!dashboardData) return;

    const data = dashboardData;

    // Header & Stats
    document.getElementById('stat-total-turnos').textContent = data.totalTurnos;
    document.getElementById('stat-modalidad-ratio').textContent = `${data.totalPresencial} presenciales • ${data.totalVirtual} virtuales`;
    
    const estadosEl = document.getElementById('stat-estados-ratio');
    if (estadosEl) {
        estadosEl.textContent = `${data.totalAtendidosGlobal || 0} atendidos • ${data.totalAusentesGlobal || 0} ausentes • ${data.totalCanceladosGlobal || 0} cancelados`;
    }

    document.getElementById('stat-total-pacientes').textContent = data.totalControlActivos;
    document.getElementById('stat-inconsistencias-count').textContent = `${data.inconsistencias.length} con huecos de control`;

    const totalAlertas = data.inactivosConTurnos.length + data.unmatchedAlerts.length + (data.sinControlConTurnos ? data.sinControlConTurnos.length : 0);
    document.getElementById('stat-total-alertas').textContent = totalAlertas;
    document.getElementById('stat-alertas-detalle').textContent = `${data.inactivosConTurnos.length} inactivos • ${data.sinControlConTurnos ? data.sinControlConTurnos.length : 0} sin control • ${data.unmatchedAlerts.length} no coincidentes`;

    // Recaudación stat card
    const statRecaudadoEl = document.getElementById('stat-total-recaudado');
    const statCobranzasEl = document.getElementById('stat-cobranzas-detalle');
    if (statRecaudadoEl) {
        statRecaudadoEl.textContent = `$${(data.totalRecaudado || 0).toLocaleString('es-AR')}`;
    }
    if (statCobranzasEl) {
        statCobranzasEl.textContent = `${data.totalAtenciones || 0} atenciones • ${data.totalPendientesCobro || 0} pendientes (deben)`;
    }

    // Counts on tabs
    const countPacEl = document.getElementById('count-pacientes');
    if (countPacEl) countPacEl.textContent = data.pacientesHistorial ? data.pacientesHistorial.length : 0;

    const countSinCtrlEl = document.getElementById('count-sin-control');
    if (countSinCtrlEl) countSinCtrlEl.textContent = data.sinControlConTurnos ? data.sinControlConTurnos.length : 0;

    document.getElementById('count-huecos').textContent = data.inconsistencias.length;
    document.getElementById('count-inactivos').textContent = data.inactivosConTurnos.length;
    document.getElementById('count-unmatched').textContent = data.unmatchedAlerts.length;

    // Critical Banner
    const banner = document.getElementById('critical-banner');
    const sobrecupos = data.capacitySummary.filter(c => c.status === 'overflow');
    if (sobrecupos.length > 0 || data.inactivosConTurnos.length > 0) {
        banner.style.display = 'flex';
        document.getElementById('stat-meses-sobrecupo').textContent = `${sobrecupos.length} meses con sobrecupo`;
        banner.querySelector('.alert-banner-content').innerHTML = `
            <strong>Atención requerida:</strong> Se detectaron <strong>${sobrecupos.length} meses con sobrecupo</strong> de turnos 
            y <strong>${data.inactivosConTurnos.length} pacientes inactivos</strong> con turnos asignados a futuro.
        `;
    } else {
        banner.style.display = 'none';
        document.getElementById('stat-meses-sobrecupo').textContent = `Capacidad en regla`;
    }

    // Capacity Grid
    renderCapacityGrid(data.capacitySummary);

    // Weekly Matrix & Intermonthly Comparison
    if (data.weeklyAnalysis) {
        renderWeeklyMatrix(data.weeklyAnalysis);
    }

    // Tables
    populatePacientesDropdown();
    filterHistorialPacientes();
    renderSinControlTable(data.sinControlConTurnos || []);
    filterInconsistencias();
    renderInactivosTable(data.inactivosConTurnos);
    renderUnmatchedTable(data.unmatchedAlerts);

    // Footer timestamp
    const dateStr = new Date(data.timestamp).toLocaleTimeString();
    document.getElementById('footer-last-sync').textContent = `Última lectura: ${dateStr}`;
}

function renderWeeklyMatrix(weeklyAnalysis) {
    const container = document.getElementById('weekly-matrix-grid');
    if (!container) return;
    container.innerHTML = '';

    if (!weeklyAnalysis || weeklyAnalysis.length === 0) {
        container.innerHTML = '<p style="color: var(--text-muted); padding: 1.5rem;">No se encontraron datos semanales disponibles.</p>';
        return;
    }

    weeklyAnalysis.forEach(monthItem => {
        const card = document.createElement('div');
        card.className = 'weekly-month-card';

        // Trend badge
        let trendBadgeHtml = '';
        const comp = monthItem.comparacionMesAnterior;
        if (comp && comp.mesAnteriorNombre) {
            if (comp.tendencia === 'sube') {
                trendBadgeHtml = `<span class="trend-badge up" title="Comparado con ${comp.mesAnteriorNombre} (${comp.mesAnteriorTotal} turnos)">▲ +${comp.deltaTurnos} vs ${comp.mesAnteriorNombre} (+${comp.deltaPorcentual}%)</span>`;
            } else if (comp.tendencia === 'baja') {
                trendBadgeHtml = `<span class="trend-badge down" title="Comparado con ${comp.mesAnteriorNombre} (${comp.mesAnteriorTotal} turnos)">▼ ${comp.deltaTurnos} vs ${comp.mesAnteriorNombre} (${comp.deltaPorcentual}%)</span>`;
            } else {
                trendBadgeHtml = `<span class="trend-badge equal" title="Comparado con ${comp.mesAnteriorNombre}">▬ = vs ${comp.mesAnteriorNombre} (0%)</span>`;
            }
        }

        // Peak week info
        let peakHtml = '';
        if (monthItem.semanaPico) {
            peakHtml = `
                <div class="weekly-meta-item peak-indicator">
                    <span class="peak-icon">🔥</span>
                    <div>
                        <span class="meta-label">Semana Pico (Mayor Carga):</span>
                        <strong class="text-rose">${monthItem.semanaPico.semana} (${monthItem.semanaPico.turnos} turnos)</strong>
                        <span class="meta-sub">${monthItem.semanaPico.rango}</span>
                    </div>
                </div>
            `;
        }

        // Build list of weeks
        let weeksRowsHtml = '';
        monthItem.semanas.forEach(w => {
            const isPeak = monthItem.semanaPico && monthItem.semanaPico.semana === w.label;
            const pct = Math.min(100, Math.round((w.count / 30) * 100));

            let paymentBadge = '';
            if (w.pagados > 0 || w.deben > 0) {
                paymentBadge = `
                    <span class="week-pill-payment">
                        <span class="text-success">✅ ${w.pagados} pagados</span>
                        ${w.deben > 0 ? `<span class="text-danger">⚠️ ${w.deben} deben</span>` : ''}
                    </span>
                `;
            }

            let statusBreakdownHtml = `
                <div class="week-status-breakdown">
                    <span class="week-pill-status atendido">✓ ${w.atendidos} atendidos</span>
                    ${w.ausentes > 0 ? `<span class="week-pill-status ausente">✕ ${w.ausentes} ausente${w.ausentes > 1 ? 's' : ''}</span>` : ''}
                    ${w.cancelados > 0 ? `<span class="week-pill-status cancelado">🚫 ${w.cancelados} canc.</span>` : ''}
                    ${w.agendados > 0 ? `<span class="week-pill-status agendado">🗓️ ${w.agendados} agend.</span>` : ''}
                </div>
            `;

            weeksRowsHtml += `
                <div class="week-row ${isPeak ? 'is-peak-week' : ''}">
                    <div class="week-left">
                        <div class="week-label-group">
                            <span class="week-name">${w.label}</span>
                            ${isPeak ? '<span class="badge-peak-pill">MÁXIMA CARGA</span>' : ''}
                        </div>
                        <span class="week-dates">${w.dateRange}</span>
                    </div>
                    <div class="week-center">
                        <div class="week-bar-track">
                            <div class="week-bar-fill ${isPeak ? 'peak-fill' : ''}" style="width: ${pct}%;"></div>
                        </div>
                        <div class="week-modalities">
                            <span>Presenciales: <strong>${w.presencial}</strong></span>
                            <span>Virtuales: <strong>${w.virtual}</strong></span>
                            ${statusBreakdownHtml}
                            ${paymentBadge}
                        </div>
                    </div>
                    <div class="week-right">
                        <span class="week-count-badge ${isPeak ? 'peak-count' : ''}">${w.count}</span>
                    </div>
                </div>
            `;
        });

        card.innerHTML = `
            <div class="weekly-card-header">
                <div class="weekly-month-title">
                    <h3>${monthItem.mesNombre} 2026</h3>
                    <div style="display: flex; gap: 6px; align-items: center; flex-wrap: wrap;">
                        <span class="weekly-total-pill">${monthItem.totalTurnos} turnos totales</span>
                        <span class="weekly-state-pill" style="font-size: 0.73rem; color: #a5f3fc; background: rgba(56, 189, 248, 0.1); padding: 0.15rem 0.5rem; border-radius: 4px; border: 1px solid rgba(56, 189, 248, 0.2);">
                            ${monthItem.totalAtendidos} atendidos • ${monthItem.totalAusentes} ausentes${monthItem.totalCancelados > 0 ? ` • ${monthItem.totalCancelados} canc.` : ''}${monthItem.totalAgendados > 0 ? ` • ${monthItem.totalAgendados} agend.` : ''}
                        </span>
                    </div>
                </div>
                <div class="weekly-trend-box">
                    ${trendBadgeHtml}
                </div>
            </div>

            <div class="weekly-meta-bar">
                <div class="weekly-meta-item">
                    <span class="meta-label">Promedio semanal:</span>
                    <strong class="text-cyan">${monthItem.promedioSemanal} turnos/sem</strong>
                </div>
                ${peakHtml}
            </div>

            <div class="weekly-rows-container">
                ${weeksRowsHtml}
            </div>
        `;

        container.appendChild(card);
    });
}

function renderCapacityGrid(capacitySummary) {
    const grid = document.getElementById('capacity-grid');
    grid.innerHTML = '';

    capacitySummary.forEach(c => {
        const isOverflow = c.status === 'overflow';
        const card = document.createElement('div');
        card.className = `cap-card ${c.status}`;

        let statusBadge = '';
        if (isOverflow) {
            statusBadge = `<span class="cap-status-pill red">Sobrecupo: ${c.margin}</span>`;
        } else if (c.status === 'full') {
            statusBadge = `<span class="cap-status-pill yellow">Capacidad Completa</span>`;
        } else {
            statusBadge = `<span class="cap-status-pill green">+${c.margin} Lugares Libres</span>`;
        }

        const barWidth = Math.min(100, c.pctOcupacion);
        const barClass = isOverflow ? 'red' : 'green';

        card.innerHTML = `
            <div class="cap-month">
                <span>${c.monthName} ${c.year}</span>
                ${statusBadge}
            </div>
            <div class="cap-numbers">
                <span class="cap-count ${isOverflow ? 'text-danger' : 'text-success'}">${c.count}</span>
                <span class="cap-limit">/ ${c.limit} efectivos</span>
            </div>
            <div class="cap-progress-track">
                <div class="cap-progress-fill ${barClass}" style="width: ${barWidth}%"></div>
            </div>
            <div class="cap-modalities">
                <span>🏢 Presenciales: <strong>${c.presencial}</strong></span>
                <span>💻 Virtuales: <strong>${c.virtual}</strong></span>
            </div>
            <div class="cap-states" style="margin-top: 6px; font-size: 0.72rem; display: flex; gap: 8px; flex-wrap: wrap; border-top: 1px solid rgba(255, 255, 255, 0.05); padding-top: 4px;">
                <span style="color: #34d399; font-weight: 600;">✓ ${c.atendidos} Atendidos</span>
                ${c.ausentes > 0 ? `<span style="color: #fda4af; font-weight: 600;">✕ ${c.ausentes} Ausentes</span>` : ''}
                ${c.cancelados > 0 ? `<span style="color: #94a3b8; font-weight: 600;">🚫 ${c.cancelados} Canc.</span>` : ''}
                ${c.agendados > 0 ? `<span style="color: #7dd3fc; font-weight: 600;">🗓️ ${c.agendados} Agendados</span>` : ''}
            </div>
        `;
        grid.appendChild(card);
    });
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));

    document.getElementById(`tab-btn-${tabId}`).classList.add('active');
    document.getElementById(`panel-${tabId}`).classList.add('active');
}

function filterInconsistencias() {
    if (!dashboardData) return;

    const query = document.getElementById('input-search-paciente').value.toLowerCase().trim();
    const monthFilter = document.getElementById('filter-month').value;
    const tbody = document.getElementById('tbody-inconsistencias');
    tbody.innerHTML = '';

    const list = dashboardData.inconsistencias.filter(item => {
        const matchesQuery = !query || item.Paciente.toLowerCase().includes(query);
        const matchesMonth = monthFilter === 'all' || item.mesesFaltantesArr.includes(parseInt(monthFilter));
        return matchesQuery && matchesMonth;
    });

    if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 2rem;">No se encontraron inconsistencias con los filtros aplicados.</td></tr>`;
        return;
    }

    list.forEach(item => {
        const tr = document.createElement('tr');
        const mesesChips = item.mesesFaltantesArr.map(m => `<span class="chip red">${NOMBRES_MESES[m] || m}</span>`).join(' ');

        tr.innerHTML = `
            <td><strong>${item.Paciente}</strong></td>
            <td><span class="chip blue">${item.Frecuencia || 'Control'}</span></td>
            <td>${mesesChips}</td>
            <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem;">${item['Arranque (Si es nuevo)'] || '-'}</td>
            <td style="color: var(--text-secondary); font-size: 0.82rem;">
                ${item.Comentario ? `<div><strong>Nota:</strong> ${item.Comentario}</div>` : ''}
                ${item.Seguimiento ? `<div><strong>Seguimiento:</strong> ${item.Seguimiento}</div>` : ''}
                ${!item.Comentario && !item.Seguimiento ? '<span style="color: var(--text-muted);">-</span>' : ''}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderInactivosTable(inactivos) {
    const tbody = document.getElementById('tbody-inactivos');
    tbody.innerHTML = '';

    if (inactivos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem;">✓ No hay pacientes inactivos con turnos a futuro.</td></tr>`;
        return;
    }

    inactivos.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong class="text-danger">${item.Paciente}</strong></td>
            <td style="font-family: 'JetBrains Mono', monospace;">${item['Turnos Futuros']}</td>
            <td><span class="chip red">Verificar o Liberar Turno</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderUnmatchedTable(unmatched) {
    const tbody = document.getElementById('tbody-unmatched');
    tbody.innerHTML = '';

    if (unmatched.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: var(--text-muted); padding: 2rem;">✓ Todos los turnos coinciden con pacientes registrados.</td></tr>`;
        return;
    }

    unmatched.forEach(item => {
        const tr = document.createElement('tr');
        const isBlock = item.Turno.toLowerCase().includes('no poner') || item.Turno.toLowerCase().includes('bloqueo');

        tr.innerHTML = `
            <td style="font-family: 'JetBrains Mono', monospace;">${item.Fecha}</td>
            <td><strong>"${item.Turno}"</strong></td>
            <td><span class="chip blue">${item.Modalidad}</span></td>
            <td>${isBlock ? '<span class="chip yellow">Bloqueo de Agenda</span>' : '<span class="chip red">Paciente Nuevo / Revisar</span>'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function populatePacientesDropdown() {
    const select = document.getElementById('select-paciente-especifico');
    if (!select || !dashboardData || !dashboardData.pacientesHistorial) return;

    const currentVal = select.value;
    select.innerHTML = '<option value="">-- Elige un paciente de la lista... --</option>';

    const sorted = [...dashboardData.pacientesHistorial].sort((a, b) => a.fullName.localeCompare(b.fullName));
    sorted.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = `${p.fullName} (${p.frecuenciaTexto})`;
        select.appendChild(opt);
    });

    if (currentVal) select.value = currentVal;
}

function onPacienteSelectChange() {
    const select = document.getElementById('select-paciente-especifico');
    const filterSelect = document.getElementById('filter-estado-paciente');
    const searchInput = document.getElementById('input-search-historial');

    if (select && select.value) {
        if (filterSelect) filterSelect.value = '';
        if (searchInput) searchInput.value = '';
    }
    filterHistorialPacientes();
}

function onFiltroGrupoChange() {
    const select = document.getElementById('select-paciente-especifico');
    const filterSelect = document.getElementById('filter-estado-paciente');
    const searchInput = document.getElementById('input-search-historial');

    if (filterSelect && filterSelect.value) {
        if (select) select.value = '';
        if (searchInput) searchInput.value = '';
    }
    filterHistorialPacientes();
}

function onSearchHistorialInput() {
    const select = document.getElementById('select-paciente-especifico');
    const filterSelect = document.getElementById('filter-estado-paciente');
    const searchInput = document.getElementById('input-search-historial');

    if (searchInput && searchInput.value.trim()) {
        if (select) select.value = '';
        if (filterSelect) filterSelect.value = '';
    }
    filterHistorialPacientes();
}

function onFiltroMesChange() {
    filterHistorialPacientes();
}

function filterHistorialPacientes() {
    if (!dashboardData || !dashboardData.pacientesHistorial) return;

    const select = document.getElementById('select-paciente-especifico');
    const selectedPacienteId = select ? select.value : '';

    const filterSelect = document.getElementById('filter-estado-paciente');
    const filterValue = filterSelect ? filterSelect.value : '';

    const monthSelect = document.getElementById('filter-mes-paciente');
    const monthVal = monthSelect ? monthSelect.value : 'all';
    const selectedMonth = monthVal === 'all' ? null : parseInt(monthVal, 10);

    const searchInput = document.getElementById('input-search-historial');
    const query = searchInput ? searchInput.value.toLowerCase().trim() : '';

    const placeholder = document.getElementById('placeholder-sin-seleccion');
    const summaryStrip = document.getElementById('historial-summary-strip');
    const tableContainer = document.getElementById('table-container-pacientes');
    const tbody = document.getElementById('tbody-pacientes-historial');

    // Dynamic table header labels based on month
    const thAtenciones = document.getElementById('th-col-atenciones');
    const thPagos = document.getElementById('th-col-pagos');
    if (thAtenciones) {
        thAtenciones.textContent = selectedMonth !== null 
            ? `Turnos & Atenciones (${NOMBRES_MESES[selectedMonth]})` 
            : 'Días que se Atendió (desde Agosto)';
    }
    if (thPagos) {
        thPagos.textContent = selectedMonth !== null 
            ? `Estado de Pago (${NOMBRES_MESES[selectedMonth]})` 
            : 'Estado de Pago (Pagó o Debe)';
    }

    const hasSelection = Boolean(selectedPacienteId || filterValue || query || selectedMonth !== null);

    if (!hasSelection) {
        if (placeholder) placeholder.style.display = 'flex';
        if (summaryStrip) summaryStrip.style.display = 'none';
        if (tableContainer) tableContainer.style.display = 'none';
        if (tbody) tbody.innerHTML = '';
        return;
    }

    if (placeholder) placeholder.style.display = 'none';
    if (summaryStrip) summaryStrip.style.display = 'flex';
    if (tableContainer) tableContainer.style.display = 'block';
    if (tbody) tbody.innerHTML = '';

    let list = dashboardData.pacientesHistorial;

    if (selectedPacienteId) {
        list = list.filter(item => String(item.id) === String(selectedPacienteId));
    } else if (filterValue) {
        list = list.filter(item => {
            if (selectedMonth !== null) {
                switch (filterValue) {
                    case 'deben':
                        return item.controlesAtendidos.some(t => t.m === selectedMonth && t.estadoPago === 'debe');
                    case 'pagados':
                        return item.controlesAtendidos.some(t => t.m === selectedMonth && t.estadoPago === 'pagado');
                    case 'con_ausencias':
                        return item.turnosAusentes && item.turnosAusentes.some(t => t.m === selectedMonth);
                    case 'con_futuros':
                        return item.turnosProximos.some(t => t.m === selectedMonth);
                    case 'frec_30':
                        return item.frecuenciaDias === 30 && (item.controlesAtendidos.some(t => t.m === selectedMonth) || item.turnosProximos.some(t => t.m === selectedMonth));
                    case 'frec_60':
                        return item.frecuenciaDias === 60 && (item.controlesAtendidos.some(t => t.m === selectedMonth) || item.turnosProximos.some(t => t.m === selectedMonth));
                    case 'frec_15':
                        return item.frecuenciaDias === 15 && (item.controlesAtendidos.some(t => t.m === selectedMonth) || item.turnosProximos.some(t => t.m === selectedMonth));
                    case 'con_atenciones':
                    default:
                        return item.controlesAtendidos.some(t => t.m === selectedMonth);
                }
            } else {
                switch (filterValue) {
                    case 'deben':
                        return item.cantPendientesPago > 0;
                    case 'pagados':
                        return item.cantPagas > 0;
                    case 'con_ausencias':
                        return (item.cantAusencias || 0) > 0;
                    case 'con_futuros':
                        return item.turnosProximos && item.turnosProximos.length > 0;
                    case 'frec_30':
                        return item.frecuenciaDias === 30;
                    case 'frec_60':
                        return item.frecuenciaDias === 60;
                    case 'frec_15':
                        return item.frecuenciaDias === 15;
                    case 'con_atenciones':
                    default:
                        return item.cantAtendidos > 0;
                }
            }
        });
    } else if (query) {
        list = list.filter(item => 
            item.fullName.toLowerCase().includes(query) || 
            item.obraSocial.toLowerCase().includes(query) || 
            (item.comentario && item.comentario.toLowerCase().includes(query))
        );
        if (selectedMonth !== null) {
            list = list.filter(item => 
                item.controlesAtendidos.some(t => t.m === selectedMonth) || 
                (item.turnosAusentes && item.turnosAusentes.some(t => t.m === selectedMonth)) ||
                item.turnosProximos.some(t => t.m === selectedMonth)
            );
        }
    } else if (selectedMonth !== null) {
        // Only month filter selected
        list = list.filter(item => 
            item.controlesAtendidos.some(t => t.m === selectedMonth) || 
            (item.turnosAusentes && item.turnosAusentes.some(t => t.m === selectedMonth)) ||
            item.turnosProximos.some(t => t.m === selectedMonth)
        );
    }

    // Update summary pills
    let totalRecaudadoFiltrado = 0;
    let totalAtencionesFiltrado = 0;
    let totalAusenciasFiltrado = 0;
    let totalDebenFiltrado = 0;
    let totalMontoDebenFiltrado = 0;

    if (selectedMonth !== null) {
        list.forEach(item => {
            item.controlesAtendidos.forEach(at => {
                if (at.m === selectedMonth) {
                    totalAtencionesFiltrado++;
                    if (at.estadoPago === 'pagado') totalRecaudadoFiltrado += (at.importe || 0);
                    if (at.estadoPago === 'debe') {
                        totalDebenFiltrado++;
                        totalMontoDebenFiltrado += (at.importe || 0);
                    }
                }
            });
            (item.turnosAusentes || []).forEach(aus => {
                if (aus.m === selectedMonth) totalAusenciasFiltrado++;
            });
            // También contabilizar turnos del mes que ya tengan pago registrado (adelantado) o deuda
            item.turnosProximos.forEach(fut => {
                if (fut.m === selectedMonth) {
                    if (fut.estadoPago === 'pagado') totalRecaudadoFiltrado += (fut.importe || 0);
                    if (fut.estadoPago === 'debe') {
                        totalDebenFiltrado++;
                        totalMontoDebenFiltrado += (fut.importe || 0);
                    }
                }
            });
        });
    } else {
        list.forEach(item => {
            totalRecaudadoFiltrado += (item.totalAbonado || 0);
            totalAtencionesFiltrado += item.cantAtendidos;
            totalAusenciasFiltrado += (item.cantAusencias || 0);
            totalDebenFiltrado += item.cantPendientesPago;
            item.controlesAtendidos.forEach(at => {
                if (at.estadoPago === 'debe') totalMontoDebenFiltrado += (at.importe || 0);
            });
        });
    }

    const sumCount = document.getElementById('sum-pacientes-count');
    const sumRec = document.getElementById('sum-total-recaudado');
    const sumAt = document.getElementById('sum-total-atenciones');
    const sumAus = document.getElementById('sum-total-ausentes');
    const sumDeb = document.getElementById('sum-total-deben');

    if (sumCount) sumCount.textContent = list.length;
    if (sumRec) sumRec.textContent = `$${totalRecaudadoFiltrado.toLocaleString('es-AR')}`;
    if (sumAt) sumAt.textContent = totalAtencionesFiltrado;
    if (sumAus) sumAus.textContent = totalAusenciasFiltrado;
    if (sumDeb) {
        sumDeb.textContent = totalMontoDebenFiltrado > 0 
            ? `${totalDebenFiltrado} ($${totalMontoDebenFiltrado.toLocaleString('es-AR')})` 
            : totalDebenFiltrado;
    }

    if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 2.5rem;">No se encontraron pacientes con los criterios de búsqueda y filtros seleccionados.</td></tr>`;
        return;
    }

    list.forEach(item => {
        const tr = document.createElement('tr');

        // 1. Paciente & Frecuencia
        const col1 = `
            <td>
                <div><strong style="font-size: 0.95rem; color: #ffffff;">${item.fullName}</strong></div>
                <div style="font-size: 0.75rem; color: var(--text-secondary); margin-top: 2px;">
                    ${item.obraSocial ? `🏥 ${item.obraSocial}` : 'Particular'}
                </div>
                <div style="margin-top: 5px; display: flex; align-items: center; gap: 4px; flex-wrap: wrap;">
                    <span class="frecuencia-pill">⏱️ ${item.frecuenciaTexto}</span>
                    <span class="chip ${item.isActivo ? 'blue' : 'red'}" style="font-size: 0.68rem;">${item.isActivo ? 'Activo' : 'Inactivo'}</span>
                    ${item.cantAusencias > 0 ? `<span class="chip red" style="font-size: 0.68rem;">✕ ${item.cantAusencias} Ausente${item.cantAusencias > 1 ? 's' : ''}</span>` : ''}
                </div>
                ${item.comentario ? `<div style="font-size: 0.74rem; color: #fbbf24; margin-top: 4px;">📝 ${item.comentario}</div>` : ''}
            </td>
        `;

        // 2. Días que se Atendió / Turnos
        let col2Content = '';
        if (selectedMonth === null) {
            const hasTurnos = item.controlesAtendidos.length > 0 || (item.turnosAusentes && item.turnosAusentes.length > 0);
            if (!hasTurnos) {
                col2Content = '<span style="color: var(--text-muted); font-size: 0.8rem;">Sin atenciones ni registros desde agosto</span>';
            } else {
                col2Content = `
                    <div class="atenciones-list">
                        ${item.controlesAtendidos.map(at => `
                            <div class="atencion-item">
                                <span class="atencion-fecha">📅 ${at.fechaStr}</span>
                                <span style="font-size: 0.74rem; color: var(--text-secondary);">${at.dayName.slice(0,3)} ${at.hora}</span>
                                <span class="atencion-mod">${at.mod === 'presencial' ? '🏢 Presencial' : (at.mod === 'virtual' ? '💻 Virtual' : 'Sin def.')}</span>
                                <span class="chip-pago ${at.estadoPago}">${at.pagoLabel}</span>
                                ${at.comentario ? `<span class="comentario-pago" title="${at.comentario}">💬 ${at.comentario}</span>` : ''}
                            </div>
                        `).join('')}
                        ${(item.turnosAusentes || []).map(aus => `
                            <div class="atencion-item" style="border-left: 2px solid var(--rose); background: rgba(244, 63, 94, 0.08);">
                                <span class="atencion-fecha" style="color: #fda4af;">📅 ${aus.fechaStr}</span>
                                <span style="font-size: 0.74rem; color: var(--text-secondary);">${aus.dayName.slice(0,3)} ${aus.hora}</span>
                                <span class="atencion-mod">${aus.mod === 'presencial' ? '🏢 Presencial' : (aus.mod === 'virtual' ? '💻 Virtual' : 'Sin def.')}</span>
                                <span class="chip-pago debe" style="background: rgba(244, 63, 94, 0.25); color: #fda4af; border: 1px solid rgba(244, 63, 94, 0.4);">✕ Inasistencia (Ausente)</span>
                                ${aus.comentario ? `<span class="comentario-pago" title="${aus.comentario}">💬 ${aus.comentario}</span>` : ''}
                            </div>
                        `).join('')}
                    </div>
                `;
            }
        } else {
            const monthAts = item.controlesAtendidos.filter(t => t.m === selectedMonth);
            const monthAus = (item.turnosAusentes || []).filter(t => t.m === selectedMonth);
            const monthFuts = item.turnosProximos.filter(t => t.m === selectedMonth);

            if (monthAts.length === 0 && monthAus.length === 0 && monthFuts.length === 0) {
                col2Content = `<span style="color: var(--text-muted); font-size: 0.8rem;">Sin turnos en ${NOMBRES_MESES[selectedMonth]}</span>`;
            } else {
                col2Content = `
                    <div class="atenciones-list">
                        ${monthAts.map(at => `
                            <div class="atencion-item">
                                <span class="atencion-fecha">📅 ${at.fechaStr}</span>
                                <span style="font-size: 0.74rem; color: var(--text-secondary);">${at.dayName.slice(0,3)} ${at.hora}</span>
                                <span class="atencion-mod">${at.mod === 'presencial' ? '🏢 Presencial' : (at.mod === 'virtual' ? '💻 Virtual' : 'Sin def.')}</span>
                                <span class="chip-pago ${at.estadoPago}">${at.pagoLabel}</span>
                                ${at.comentario ? `<span class="comentario-pago" title="${at.comentario}">💬 ${at.comentario}</span>` : ''}
                            </div>
                        `).join('')}
                        ${monthAus.map(aus => `
                            <div class="atencion-item" style="border-left: 2px solid var(--rose); background: rgba(244, 63, 94, 0.08);">
                                <span class="atencion-fecha" style="color: #fda4af;">📅 ${aus.fechaStr}</span>
                                <span style="font-size: 0.74rem; color: var(--text-secondary);">${aus.dayName.slice(0,3)} ${aus.hora}</span>
                                <span class="atencion-mod">${aus.mod === 'presencial' ? '🏢 Presencial' : (aus.mod === 'virtual' ? '💻 Virtual' : 'Sin def.')}</span>
                                <span class="chip-pago debe" style="background: rgba(244, 63, 94, 0.25); color: #fda4af; border: 1px solid rgba(244, 63, 94, 0.4);">✕ Inasistencia (Ausente)</span>
                                ${aus.comentario ? `<span class="comentario-pago" title="${aus.comentario}">💬 ${aus.comentario}</span>` : ''}
                            </div>
                        `).join('')}
                        ${monthFuts.map(fut => `
                            <div class="atencion-item" style="border-left: 2px solid var(--cyan);">
                                <span class="atencion-fecha">📅 ${fut.fechaStr}</span>
                                <span style="font-size: 0.74rem; color: var(--text-secondary);">${fut.dayName.slice(0,3)} ${fut.hora}</span>
                                <span class="atencion-mod">${fut.mod === 'presencial' ? '🏢 Presencial' : (fut.mod === 'virtual' ? '💻 Virtual' : 'Sin def.')}</span>
                                <span class="chip-pago ${fut.estadoPago}">${fut.pagoLabel || 'Agendado a futuro'}</span>
                                ${fut.comentario ? `<span class="comentario-pago" title="${fut.comentario}">💬 ${fut.comentario}</span>` : ''}
                            </div>
                        `).join('')}
                    </div>
                `;
            }
        }
        const col2 = `<td>${col2Content}</td>`;

        // 3. Estado de Pago (Pagó o Debe)
        let estadoCobroHtml = '';
        if (selectedMonth === null) {
            if (item.cantAtendidos === 0) {
                estadoCobroHtml = '<span style="color: var(--text-muted); font-size: 0.8rem;">-</span>';
            } else {
                estadoCobroHtml = `
                    <div class="financial-summary">
                        <div>Total Pagado: <strong class="total-pago">${item.totalAbonadoFormatted}</strong></div>
                        <div style="font-size: 0.74rem; color: var(--text-secondary);">${item.cantPagas} consulta(s) abonada(s)</div>
                        ${item.cantPendientesPago > 0 
                            ? `<div class="debe-warning">⚠️ DEBE ${item.cantPendientesPago} consulta(s)</div>` 
                            : `<div style="color: var(--emerald); font-size: 0.74rem; font-weight: 600;">✓ Al día</div>`
                        }
                    </div>
                `;
            }
        } else {
            const monthAts = item.controlesAtendidos.filter(t => t.m === selectedMonth);
            let mPaid = 0;
            let mCountPaid = 0;
            let mCountDebe = 0;

            monthAts.forEach(at => {
                if (at.estadoPago === 'pagado') {
                    mPaid += (at.importe || 0);
                    mCountPaid++;
                } else if (at.estadoPago === 'debe') {
                    mCountDebe++;
                }
            });

            if (monthAts.length > 0) {
                estadoCobroHtml = `
                    <div class="financial-summary">
                        <div>Pagado en ${NOMBRES_MESES[selectedMonth]}: <strong class="total-pago">$${mPaid.toLocaleString('es-AR')}</strong></div>
                        <div style="font-size: 0.74rem; color: var(--text-secondary);">${mCountPaid} consulta(s) abonada(s)</div>
                        ${mCountDebe > 0 
                            ? `<div class="debe-warning">⚠️ DEBE ${mCountDebe} en ${NOMBRES_MESES[selectedMonth]}</div>` 
                            : `<div style="color: var(--emerald); font-size: 0.74rem; font-weight: 600;">✓ Al día en ${NOMBRES_MESES[selectedMonth]}</div>`
                        }
                    </div>
                `;
            } else if (item.turnosProximos.some(t => t.m === selectedMonth)) {
                const futMatch = item.turnosProximos.find(t => t.m === selectedMonth);
                const comTxt = futMatch && futMatch.comentario ? ` • 💬 ${futMatch.comentario}` : '';
                estadoCobroHtml = `<div style="color: var(--text-secondary); font-size: 0.74rem;">🗓️ Turno agendado (${futMatch ? futMatch.pagoLabel : 'Cobro al atender'})${comTxt}</div>`;
            } else {
                estadoCobroHtml = '<span style="color: var(--text-muted); font-size: 0.8rem;">-</span>';
            }
        }
        const col3 = `<td>${estadoCobroHtml}</td>`;

        // 4. Meses Venideros (Cuándo Debería Venir)
        const col4 = `
            <td>
                <div class="proyeccion-list">
                    ${item.proyeccionMeses.map(p => {
                        const isSelectedMonth = selectedMonth !== null && p.mes === selectedMonth;
                        const highlightStyle = isSelectedMonth ? 'style="border-color: var(--cyan); box-shadow: 0 0 8px rgba(56, 189, 248, 0.4);"' : '';
                        return `
                            <div class="proy-item" ${highlightStyle}>
                                <span class="proy-month-badge">${p.mesNombre.slice(0,3)}</span>
                                <span class="proy-tag ${p.badge}">${p.detalle}</span>
                            </div>
                        `;
                    }).join('')}
                </div>
            </td>
        `;

        // 5. Próximo Control Sugerido
        let col5Content = '';
        if (item.proximoTurno) {
            col5Content = `
                <div class="sugerido-badge" style="border-color: rgba(56, 189, 248, 0.4); background: rgba(56, 189, 248, 0.1);">
                    <span style="font-size: 0.7rem; color: #7dd3fc; font-weight: 700;">🗓️ AGENDADO</span>
                    <span class="sug-date">${item.proximoTurno.fechaStr}</span>
                    <span class="sug-frec">${item.proximoTurno.hora} (${item.proximoTurno.mod})</span>
                </div>
            `;
        } else if (item.proximaFechaSugerida && item.isActivo) {
            col5Content = `
                <div class="sugerido-badge">
                    <span style="font-size: 0.7rem; color: #c7d2fe; font-weight: 700;">💡 SUGERIDO (${item.frecuenciaDias}d)</span>
                    <span class="sug-date">${item.proximaFechaSugerida}</span>
                    <span class="sug-frec">Según frecuencia de control</span>
                </div>
            `;
        } else if (!item.isActivo) {
            col5Content = '<span class="chip red" style="font-size: 0.75rem;">Inactivo</span>';
        } else {
            col5Content = `<span style="color: var(--text-muted); font-size: 0.78rem;">${item.frecuenciaTexto || 'A demanda'}</span>`;
        }
        const col5 = `<td>${col5Content}</td>`;

        tr.innerHTML = col1 + col2 + col3 + col4 + col5;
        tbody.appendChild(tr);
    });
}

function renderSinControlTable(sinControlList) {
    const tbody = document.getElementById('tbody-sin-control');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!sinControlList || sinControlList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 2rem;">✓ Todos los pacientes con turnos tienen frecuencia de control asignada.</td></tr>`;
        return;
    }

    sinControlList.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong style="color: #ffffff;">${item.Paciente}</strong></td>
            <td><span class="chip ${item.Activo === 'Sí' ? 'blue' : 'red'}">${item.Activo === 'Sí' ? 'Activo' : 'Inactivo'}</span></td>
            <td><span class="frecuencia-pill" style="background: rgba(245, 158, 11, 0.15); color: #fbbf24;">${item.Motivo}</span></td>
            <td style="font-family: 'JetBrains Mono', monospace; font-size: 0.82rem; color: #93c5fd;">${item.TurnosDetalle}</td>
            <td><span class="chip red">Definir Frecuencia en Ficha</span></td>
        `;
        tbody.appendChild(tr);
    });
}

// Initial fetch and poll every 8 seconds for real-time live sync
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setInterval(() => loadData(false), 8000);
});
