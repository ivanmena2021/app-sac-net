// sacmap.js — Leaflet map for SAC Mapa de Calor
let sacMap = null;
let sacMarkers = [];

function sacMapInit(containerId, data, metric) {
    const el = document.getElementById(containerId);
    if (!el) return;

    // Clean previous
    if (sacMap) { sacMap.remove(); sacMap = null; }
    sacMarkers = [];

    sacMap = L.map(containerId).setView([-9.19, -75.0], 5);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; OpenStreetMap',
        maxZoom: 12,
    }).addTo(sacMap);

    sacMapUpdate(data, metric);
}

function sacMapUpdate(data, metric) {
    if (!sacMap) return;

    // Remove old markers
    sacMarkers.forEach(m => sacMap.removeLayer(m));
    sacMarkers = [];

    if (!data || data.length === 0) return;

    const maxVal = Math.max(...data.map(d => d.val || 0), 1);

    data.forEach(d => {
        const ratio = (d.val || 0) / maxVal;
        const radius = Math.max(8, Math.min(40, 8 + ratio * 32));
        const color = getColor(ratio);

        const circle = L.circleMarker([d.lat, d.lon], {
            radius: radius,
            fillColor: color,
            color: '#fff',
            weight: 1.5,
            opacity: 0.9,
            fillOpacity: 0.7,
        }).addTo(sacMap);

        const fmt = (v) => typeof v === 'number' ? v.toLocaleString('es-PE', {maximumFractionDigits: 0}) : v;
        circle.bindPopup(
            `<strong>${titleCase(d.nombre)}</strong><br>` +
            `Avisos: ${fmt(d.avisos)}<br>` +
            `Indemniz.: S/ ${fmt(d.indemnizacion)}<br>` +
            `Desembolso: S/ ${fmt(d.desembolso)}<br>` +
            `Ha Indemn.: ${(d.haIndemnizadas || 0).toLocaleString('es-PE', {maximumFractionDigits: 2})}`
        );
        circle.bindTooltip(`${titleCase(d.nombre)}: ${fmt(d.val)}`, { direction: 'top', offset: [0, -radius] });

        sacMarkers.push(circle);
    });
}

function getColor(ratio) {
    // Green → Yellow → Red gradient
    if (ratio < 0.33) return '#27ae60';
    if (ratio < 0.66) return '#f39c12';
    return '#e74c3c';
}

function titleCase(s) {
    return s.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
}
