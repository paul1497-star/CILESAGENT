const https = require("https");
const http = require("http");

const DRIVE_VENTAS_ID = "1qAvJW_RfFqaANy5COV8QQzuRNoxMNBMG";
const DRIVE_PRESUP_ID = "16lPBwVWkMT_wDTOK_mvfKcWWQ96CxxG6";
const DRIVE_ZONAS_ID  = "1CRBxDt6tfHMdfdXIz5uryfFB6amK7iea";

function downloadFile(id) {
  return new Promise((resolve, reject) => {
    const url = `https://drive.google.com/uc?export=download&id=${id}&confirm=t`;
    
    function doRequest(currentUrl, redirectCount = 0) {
      if (redirectCount > 10) return reject(new Error("Demasiadas redirecciones"));
      const lib = currentUrl.startsWith("https") ? https : http;
      lib.get(currentUrl, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          return doRequest(res.headers.location, redirectCount + 1);
        }
        if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode}`));
        const chunks = [];
        res.on("data", chunk => chunks.push(chunk));
        res.on("end", () => resolve(Buffer.concat(chunks)));
        res.on("error", reject);
      }).on("error", reject);
    }
    doRequest(url);
  });
}

const XLSX = require("xlsx");
const MESES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

exports.handler = async function(event, context) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json"
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  try {
    console.log("Descargando archivos de Drive...");
    const [bufV, bufP, bufZ] = await Promise.all([
      downloadFile(DRIVE_VENTAS_ID),
      downloadFile(DRIVE_PRESUP_ID),
      downloadFile(DRIVE_ZONAS_ID),
    ]);

    console.log(`Tamaños: ventas=${bufV.length}, presup=${bufP.length}, zonas=${bufZ.length}`);

    const wbV = XLSX.read(bufV, { type: "buffer", cellDates: true });
    const wbP = XLSX.read(bufP, { type: "buffer", cellDates: true });
    const wbZ = XLSX.read(bufZ, { type: "buffer", cellDates: true });

    // Zonas
    const zonaMap = {};
    XLSX.utils.sheet_to_json(wbZ.Sheets[wbZ.SheetNames[0]]).forEach(r => {
      if (r.ASESOR) zonaMap[String(r.ASESOR).trim().toUpperCase()] = String(r.REGIONAL || "").trim().toUpperCase();
    });

    // Presupuesto
    const presup = XLSX.utils.sheet_to_json(wbP.Sheets[wbP.SheetNames[0]], { raw: false }).map(r => ({
      asesor: String(r.ASESOR || "").trim().toUpperCase(),
      ano: new Date(r.FECHA || "").getFullYear(),
      mes: new Date(r.FECHA || "").getMonth() + 1,
      valor: parseFloat(r.PRESUPUESTO || 0) || 0
    })).filter(r => r.asesor && r.valor);

    // Ventas
    const sheetV = wbV.SheetNames.find(n => n === "Hoja2") || wbV.SheetNames[0];
    const ventas = XLSX.utils.sheet_to_json(wbV.Sheets[sheetV], { raw: false }).map(r => ({
      vendedor: String(r.NOMBRE_VENDEDOR || r.VENDEDOR || "").trim().toUpperCase(),
      cliente: String(r.CLIENTE || "").trim(),
      nomCliente: String(r.NOMBRE_CLIENTE || "").trim(),
      total: parseFloat(r.TOTAL || r.VALOR || 0) || 0,
      fecha: new Date(r.FECHA || "2025-01-01"),
      ciudad: String(r.NOMBRE_CIUDAD || r.CIUDAD || "").trim(),
      linea: String(r.NOMBRE_LINEA || "").trim(),
      ano: parseInt(r.ANO) || new Date(r.FECHA || "").getFullYear() || 2025,
      mes: parseInt(r.MES) || (new Date(r.FECHA || "").getMonth() + 1) || 1,
    })).filter(r => r.vendedor && r.total > 0);

    const hoy = new Date();
    const mesAct = hoy.getMonth() + 1;
    const anoAct = hoy.getFullYear();

    const conMeta = [...new Set(presup.filter(p => p.ano === anoAct && p.mes === mesAct).map(p => p.asesor))];

    const resultado = {};
    for (const vendedor of conMeta) {
      const vTodo = ventas.filter(v => v.vendedor === vendedor);
      if (!vTodo.length) continue;

      const vMes = vTodo.filter(v => v.ano === anoAct && v.mes === mesAct);
      const metaMes = presup.filter(p => p.asesor === vendedor && p.ano === anoAct && p.mes === mesAct).reduce((s, p) => s + p.valor, 0);
      const ventasMes = vMes.reduce((s, v) => s + v.total, 0);
      const ventas2025 = vTodo.filter(v => v.ano === 2025).reduce((s, v) => s + v.total, 0);

      // Clientes
      const cMap = {};
      vTodo.forEach(v => {
        if (!cMap[v.cliente]) cMap[v.cliente] = { nombre: v.nomCliente, ventas: 0, ultima: v.fecha, ciudad: v.ciudad, peds: 0 };
        cMap[v.cliente].ventas += v.total;
        cMap[v.cliente].peds++;
        if (v.fecha > cMap[v.cliente].ultima) cMap[v.cliente].ultima = v.fecha;
      });

      const clientes_list = Object.entries(cMap).map(([, c]) => {
        const dias = Math.floor((hoy - new Date(c.ultima)) / 86400000);
        return { n: c.nombre.slice(0, 42), v: Math.round(c.ventas), d: dias, r: dias > 90 ? "Alto" : dias > 45 ? "Medio" : "Bajo", c: c.ciudad };
      }).sort((a, b) => b.v - a.v).slice(0, 30);

      // Evolución 12 meses
      const evolucion = [];
      for (let i = 11; i >= 0; i--) {
        const d = new Date(anoAct, mesAct - 1 - i, 1);
        const a = d.getFullYear(), m = d.getMonth() + 1;
        const meta = presup.filter(p => p.asesor === vendedor && p.ano === a && p.mes === m).reduce((s, p) => s + p.valor, 0);
        const venta = vTodo.filter(v => v.ano === a && v.mes === m).reduce((s, v) => s + v.total, 0);
        if (meta > 0 || venta > 0) evolucion.push({
          label: `${MESES[m]} ${String(a).slice(2)}`, meta, ventas: venta,
          pct: meta > 0 ? Math.round(venta / meta * 10) / 10 : 0
        });
      }

      resultado[vendedor] = {
        regional: zonaMap[vendedor] || "SIN ZONA",
        meta: Math.round(metaMes), ventas: Math.round(ventasMes),
        pct: metaMes > 0 ? Math.round(ventasMes / metaMes * 1000) / 10 : 0,
        ventas2025: Math.round(ventas2025),
        clientes: clientes_list.length,
        riesgoAlto: clientes_list.filter(c => c.r === "Alto").length,
        riesgoMedio: clientes_list.filter(c => c.r === "Medio").length,
        evolucion, clientes_list,
        actualizado: hoy.toISOString()
      };
    }

    return {
      statusCode: 200, headers,
      body: JSON.stringify({ ok: true, data: resultado, fecha: hoy.toISOString() })
    };
  } catch (err) {
    console.error("Error:", err);
    return {
      statusCode: 500, headers,
      body: JSON.stringify({ ok: false, error: err.message })
    };
  }
};
