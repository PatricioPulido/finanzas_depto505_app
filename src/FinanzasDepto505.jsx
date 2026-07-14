import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Plus, Trash2, TrendingUp, TrendingDown, Wallet, ChevronLeft, ChevronRight,
  Pencil, X, Check, Download, Upload, FileSpreadsheet, ListChecks, LayoutGrid,
  ClipboardPaste, AlertTriangle,
} from "lucide-react";

const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("es-CL", { month: "long", year: "numeric" });
};
const fmtCLP = (n) =>
  new Intl.NumberFormat("es-CL", { style: "currency", currency: "CLP", maximumFractionDigits: 0 }).format(n || 0);
const fmtPct = (n) => (n === null ? "—" : `${Math.round(n * 100)}%`);

// Convierte "01/07/2026" -> "2026-07-01". Devuelve null si no calza el formato.
const parsearFechaDDMMYYYY = (raw) => {
  const m = raw.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const [, d, mo, y] = m;
  return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
};

// Convierte "15,000" o "38.556" o "15000" -> 15000. Devuelve null si queda vacío.
const parsearMonto = (raw) => {
  const limpio = raw.replace(/[.,\s]/g, "").trim();
  if (!limpio) return null;
  const val = parseInt(limpio, 10);
  return isNaN(val) ? null : val;
};

const K_CATS = "depto505:categorias";
const K_MOVS = "depto505:movimientos";

const SEED_CATEGORIAS = [
  { id: "c1", tipo: "gasto", nombre: "Peaje (ida y vuelta)", presupuesto: 0 },
  { id: "c2", tipo: "gasto", nombre: "Bencina", presupuesto: 0 },
  { id: "c3", tipo: "gasto", nombre: "Aseos externos - Socio", presupuesto: 0 },
  { id: "c4", tipo: "gasto", nombre: "Aseos externos - Peti", presupuesto: 228000 },
  { id: "c5", tipo: "gasto", nombre: "Sábanas", presupuesto: 0 },
  { id: "c6", tipo: "gasto", nombre: "Útiles de aseo", presupuesto: 20000 },
  { id: "c7", tipo: "gasto", nombre: "Lavandería", presupuesto: 10000 },
  { id: "c8", tipo: "gasto", nombre: "Dividendo", presupuesto: 360000 },
  { id: "c9", tipo: "gasto", nombre: "Gastos comunes (GGCC)", presupuesto: 77000 },
  { id: "c10", tipo: "gasto", nombre: "Agua caliente", presupuesto: 20000 },
  { id: "c11", tipo: "gasto", nombre: "Agua", presupuesto: 5000 },
  { id: "c12", tipo: "gasto", nombre: "Luz", presupuesto: 25000 },
  { id: "c13", tipo: "gasto", nombre: "Transporte", presupuesto: 0 },
  { id: "c14", tipo: "gasto", nombre: "Gastos extra", presupuesto: 0 },
  { id: "i1", tipo: "ingreso", nombre: "Airbnb", presupuesto: 604391 },
  { id: "i2", tipo: "ingreso", nombre: "Booking", presupuesto: 0 },
  { id: "i3", tipo: "ingreso", nombre: "Por fuera", presupuesto: 0 },
];

export default function FinanzasDepto505() {
  const [categorias, setCategorias] = useState(null);
  const [movimientos, setMovimientos] = useState(null);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState("registro");
  const [cursor, setCursor] = useState(() => monthKey(new Date()));
  const fileInputRef = useRef(null);

  const [nuevo, setNuevo] = useState({
    tipo: "gasto",
    categoriaId: "",
    monto: "",
    descripcion: "",
    fecha: new Date().toISOString().slice(0, 10),
  });

  const [nuevaCat, setNuevaCat] = useState({ tipo: "gasto", nombre: "", presupuesto: "" });
  const [editPresId, setEditPresId] = useState(null);
  const [editPresVal, setEditPresVal] = useState("");
  const [editNombreId, setEditNombreId] = useState(null);
  const [editNombreVal, setEditNombreVal] = useState("");

  const [mostrarPegar, setMostrarPegar] = useState(false);
  const [textoPegado, setTextoPegado] = useState("");
  const [pegarResultado, setPegarResultado] = useState(null); // { filasValidas, filasInvalidas, categoriasNuevas }
  const [tiposCategoriasNuevas, setTiposCategoriasNuevas] = useState({}); // { nombreCategoria: "gasto" | "ingreso" }

  const [editMovId, setEditMovId] = useState(null);
  const [editMovDraft, setEditMovDraft] = useState({ fecha: "", categoriaId: "", descripcion: "", monto: "" });
  const [confirmAccion, setConfirmAccion] = useState(null); // { tipo: "mes" | "todo", mensaje }

  // Carga inicial: se ejecuta UNA sola vez, al montar el componente ([] al final).
  // localStorage es síncrono (a diferencia de window.storage), así que no
  // necesitamos async/await ni Promise.all aquí.
  useEffect(() => {
    try {
      // getItem devuelve un string si existe la key, o null si nunca se guardó nada.
      const catGuardadas = localStorage.getItem(K_CATS);
      const movGuardados = localStorage.getItem(K_MOVS);

      // Si hay algo guardado (no es null), lo convertimos de texto a objeto/array
      // con JSON.parse. Si no hay nada (primera vez que se abre la app en este
      // navegador), usamos las categorías semilla y una lista vacía de movimientos.
      setCategorias(catGuardadas ? JSON.parse(catGuardadas) : SEED_CATEGORIAS);
      setMovimientos(movGuardados ? JSON.parse(movGuardados) : []);
    } catch (err) {
      // Solo debería fallar si el JSON guardado quedó corrupto por alguna razón.
      console.error("Error leyendo localStorage:", err);
      setCategorias(SEED_CATEGORIAS);
      setMovimientos([]);
    }
  }, []);

  // Estas dos funciones son el "punto único de guardado" de toda la app.
  // Cada vez que agregas un movimiento, editas una categoría, importas un
  // backup, etc., el código llama a una de estas dos — nunca escribe a
  // localStorage directamente desde otro lado. Eso es bueno, mantenlo así:
  // si el día de mañana cambias de localStorage a una API real, solo tocas
  // estas dos funciones y todo el resto del componente sigue funcionando igual.

  const guardarCategorias = (next) => {
    // 1) Actualiza el estado de React -> la UI se re-renderiza al toque
    setCategorias(next);
    try {
      // 2) Persiste en localStorage -> sobrevive a recargas y cierres de pestaña
      // JSON.stringify convierte el array a texto, porque localStorage
      // SOLO puede guardar strings, nunca objetos/arrays directamente.
      localStorage.setItem(K_CATS, JSON.stringify(next));
      setError(null);
    } catch (err) {
      // Puede fallar en modo incógnito con storage bloqueado, o si se llena
      // el límite de ~5MB de localStorage (poco probable con estos datos).
      console.error("Error guardando categorías:", err);
      setError("No se pudo guardar. Intenta de nuevo.");
    }
  };

  const guardarMovimientos = (next) => {
    setMovimientos(next);
    try {
      localStorage.setItem(K_MOVS, JSON.stringify(next));
      setError(null);
    } catch (err) {
      console.error("Error guardando movimientos:", err);
      setError("No se pudo guardar. Intenta de nuevo.");
    }
  };

  useEffect(() => {
    if (categorias && categorias.length && !nuevo.categoriaId) {
      const primera = categorias.find((c) => c.tipo === nuevo.tipo);
      if (primera) setNuevo((f) => ({ ...f, categoriaId: primera.id }));
    }
  }, [categorias, nuevo.tipo, nuevo.categoriaId]);

  const movsDelMes = useMemo(
    () => (movimientos || []).filter((m) => m.fecha.slice(0, 7) === cursor),
    [movimientos, cursor]
  );

  const catMap = useMemo(() => {
    const map = {};
    (categorias || []).forEach((c) => (map[c.id] = c));
    return map;
  }, [categorias]);

  const agregarMovimiento = (e) => {
    e.preventDefault();
    const monto = parsearMonto(nuevo.monto);
    if (!monto || monto <= 0 || !nuevo.categoriaId) return;
    const cat = catMap[nuevo.categoriaId];
    if (!cat) return;
    const mov = {
      id: `mov_${Date.now()}`,
      categoriaId: cat.id,
      tipo: cat.tipo,
      descripcion: nuevo.descripcion.trim(),
      monto,
      fecha: nuevo.fecha,
    };
    guardarMovimientos([mov, ...(movimientos || [])]);
    setNuevo((f) => ({ ...f, monto: "", descripcion: "" }));
  };

  const eliminarMovimiento = (id) => {
    guardarMovimientos((movimientos || []).filter((m) => m.id !== id));
  };

  const vaciarMesActual = () => {
    const cantidad = movsDelMes.length;
    if (cantidad === 0) return;
    setConfirmAccion({
      tipo: "mes",
      mensaje: `¿Borrar los ${cantidad} movimiento(s) de ${monthLabel(cursor)}? Esto no se puede deshacer.`,
    });
  };

  const borrarTodosLosMovimientos = () => {
    const cantidad = (movimientos || []).length;
    if (cantidad === 0) return;
    setConfirmAccion({
      tipo: "todo",
      mensaje: `¿Borrar TODOS los movimientos (${cantidad} en total, de todos los meses)? Las categorías no se tocan. Esto no se puede deshacer.`,
    });
  };

  const ejecutarConfirmAccion = () => {
    if (!confirmAccion) return;
    if (confirmAccion.tipo === "mes") {
      guardarMovimientos((movimientos || []).filter((m) => m.fecha.slice(0, 7) !== cursor));
    } else if (confirmAccion.tipo === "todo") {
      guardarMovimientos([]);
    }
    setConfirmAccion(null);
  };

  const empezarEdicionMovimiento = (m) => {
    setEditMovId(m.id);
    setEditMovDraft({ fecha: m.fecha, categoriaId: m.categoriaId, descripcion: m.descripcion || "", monto: String(m.monto) });
  };

  const cancelarEdicionMovimiento = () => {
    setEditMovId(null);
    setEditMovDraft({ fecha: "", categoriaId: "", descripcion: "", monto: "" });
  };

  const guardarEdicionMovimiento = () => {
    const monto = parsearMonto(editMovDraft.monto);
    const cat = catMap[editMovDraft.categoriaId];
    if (!monto || monto <= 0 || !cat || !editMovDraft.fecha) return;
    guardarMovimientos(
      (movimientos || []).map((m) =>
        m.id === editMovId
          ? { ...m, fecha: editMovDraft.fecha, categoriaId: cat.id, tipo: cat.tipo, descripcion: editMovDraft.descripcion.trim(), monto }
          : m
      )
    );
    cancelarEdicionMovimiento();
  };

  // Toma el texto pegado (filas copiadas de Excel, separadas por tabulaciones) y arma
  // una previsualización: filas válidas, filas inválidas (sin fecha o monto reconocible),
  // y las categorías mencionadas que no existen todavía en la app.
  const procesarTextoPegado = () => {
    const nombresConocidos = new Map((categorias || []).map((c) => [c.nombre.trim().toLowerCase(), c]));
    const lineas = textoPegado.split("\n").map((l) => l.replace(/\t+$/, "")).filter((l) => l.trim() !== "");

    const filasValidas = [];
    const filasInvalidas = [];
    const nuevasVistas = new Map(); // nombre original -> true

    lineas.forEach((linea, i) => {
      const partes = linea.split("\t").map((p) => p.trim());
      const [fechaRaw, categoriaRaw, descripcionRaw, montoRaw] = partes;
      const fecha = fechaRaw ? parsearFechaDDMMYYYY(fechaRaw) : null;
      const monto = montoRaw ? parsearMonto(montoRaw) : null;

      if (!fecha || !categoriaRaw || !monto) {
        filasInvalidas.push({ n: i + 1, texto: linea, motivo: !fecha ? "fecha no reconocida" : !categoriaRaw ? "sin categoría" : "sin monto válido" });
        return;
      }

      const catExistente = nombresConocidos.get(categoriaRaw.trim().toLowerCase());
      if (!catExistente) nuevasVistas.set(categoriaRaw.trim(), true);

      filasValidas.push({
        fecha,
        categoriaNombre: categoriaRaw.trim(),
        descripcion: (descripcionRaw || "").trim(),
        monto,
        categoriaId: catExistente ? catExistente.id : null,
        tipoExistente: catExistente ? catExistente.tipo : null,
      });
    });

    const categoriasNuevas = Array.from(nuevasVistas.keys());
    setTiposCategoriasNuevas(Object.fromEntries(categoriasNuevas.map((n) => [n, "gasto"])));
    setPegarResultado({ filasValidas, filasInvalidas, categoriasNuevas });
  };

  const confirmarPegado = () => {
    if (!pegarResultado) return;
    const { filasValidas, categoriasNuevas } = pegarResultado;

    // 1) Crea las categorías nuevas que hicieron falta, con el tipo elegido en el panel.
    const categoriasCreadas = categoriasNuevas.map((nombre, i) => ({
      id: `c_pegado_${Date.now()}_${i}`,
      tipo: tiposCategoriasNuevas[nombre] || "gasto",
      nombre,
      presupuesto: 0,
    }));
    const categoriasFinal = [...(categorias || []), ...categoriasCreadas];
    const idPorNombre = new Map(categoriasFinal.map((c) => [c.nombre.trim().toLowerCase(), c]));

    // 2) Arma los movimientos, resolviendo el id/tipo de categoría (existente o recién creada).
    const nuevosMovimientos = filasValidas.map((f, i) => {
      const cat = idPorNombre.get(f.categoriaNombre.toLowerCase());
      return {
        id: `mov_pegado_${Date.now()}_${i}`,
        categoriaId: cat.id,
        tipo: cat.tipo,
        descripcion: f.descripcion,
        monto: f.monto,
        fecha: f.fecha,
      };
    });

    guardarCategorias(categoriasFinal);
    guardarMovimientos([...nuevosMovimientos, ...(movimientos || [])]);

    setTextoPegado("");
    setPegarResultado(null);
    setTiposCategoriasNuevas({});
    setMostrarPegar(false);
  };

  const cancelarPegado = () => {
    setTextoPegado("");
    setPegarResultado(null);
    setTiposCategoriasNuevas({});
    setMostrarPegar(false);
  };

  const crearCategoria = (e) => {
    e.preventDefault();
    if (!nuevaCat.nombre.trim()) return;
    const nueva = {
      id: `c_${Date.now()}`,
      tipo: nuevaCat.tipo,
      nombre: nuevaCat.nombre.trim(),
      presupuesto: parsearMonto(nuevaCat.presupuesto) || 0,
    };
    guardarCategorias([...(categorias || []), nueva]);
    setNuevaCat({ tipo: "gasto", nombre: "", presupuesto: "" });
  };

  const eliminarCategoria = (id) => {
    guardarCategorias((categorias || []).filter((c) => c.id !== id));
  };

  const renombrarCategoria = (id, nombre) => {
    if (!nombre.trim()) return;
    guardarCategorias((categorias || []).map((c) => (c.id === id ? { ...c, nombre: nombre.trim() } : c)));
    setEditNombreId(null);
  };

  const actualizarPresupuesto = (id, valor) => {
    const val = parsearMonto(valor);
    guardarCategorias((categorias || []).map((c) => (c.id === id ? { ...c, presupuesto: val === null ? 0 : val } : c)));
    setEditPresId(null);
  };

  // ---- Agregación mensual estilo Excel ----
  const resumen = useMemo(() => {
    const acumulado = {};
    movsDelMes.forEach((m) => {
      acumulado[m.categoriaId] = (acumulado[m.categoriaId] || 0) + m.monto;
    });
    const filas = (tipo) =>
      (categorias || [])
        .filter((c) => c.tipo === tipo)
        .map((c, idx) => {
          const acum = acumulado[c.id] || 0;
          const presu = c.presupuesto || 0;
          const diferencia = tipo === "gasto" ? presu - acum : acum - presu;
          const pct = presu > 0 ? acum / presu : null;
          return { n: idx + 1, id: c.id, nombre: c.nombre, acum, presu, diferencia, pct };
        });
    const egresos = filas("gasto");
    const ingresos = filas("ingreso");
    const totalEgresosAcum = egresos.reduce((s, f) => s + f.acum, 0);
    const totalEgresosPresu = egresos.reduce((s, f) => s + f.presu, 0);
    const totalIngresosAcum = ingresos.reduce((s, f) => s + f.acum, 0);
    const totalIngresosPresu = ingresos.reduce((s, f) => s + f.presu, 0);
    return {
      egresos,
      ingresos,
      totalEgresosAcum,
      totalEgresosPresu,
      totalEgresosDif: totalEgresosPresu - totalEgresosAcum,
      totalIngresosAcum,
      totalIngresosPresu,
      totalIngresosDif: totalIngresosAcum - totalIngresosPresu,
      resultadoNeto: totalIngresosAcum - totalEgresosAcum,
    };
  }, [movsDelMes, categorias]);

  const moverMes = (delta) => {
    const [y, m] = cursor.split("-").map(Number);
    setCursor(monthKey(new Date(y, m - 1 + delta, 1)));
  };

  const descargarArchivo = (contenido, nombreArchivo, tipoMime) => {
    const blob = contenido instanceof Blob ? contenido : new Blob([contenido], { type: tipoMime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nombreArchivo;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportarExcel = () => {
    const [y, mNum] = cursor.split("-").map(Number);

    // Hoja 1: Registro Diario (todos los movimientos, ordenados por fecha)
    const filasRegistro = (movimientos || [])
      .slice()
      .sort((a, b) => a.fecha.localeCompare(b.fecha))
      .map((m) => {
        const c = catMap[m.categoriaId];
        const [yy, mm] = m.fecha.split("-").map(Number);
        return {
          Fecha: m.fecha,
          Categoría: c ? c.nombre : "—",
          "Descripción / Boleta": m.descripcion || "",
          "Monto ($CLP)": m.monto,
          Mes: mm,
          Año: yy,
        };
      });
    const hojaRegistro = XLSX.utils.json_to_sheet(filasRegistro);
    hojaRegistro["!cols"] = [{ wch: 12 }, { wch: 24 }, { wch: 36 }, { wch: 14 }, { wch: 6 }, { wch: 6 }];

    // Hoja 2: Resumen Mensual del mes actualmente seleccionado
    const filasResumen = [];
    filasResumen.push([`Resumen financiero mensual — Depto 505`]);
    filasResumen.push(["Mes (número):", mNum, "Año:", y, monthLabel(cursor)]);
    filasResumen.push([]);
    filasResumen.push(["EGRESOS"]);
    filasResumen.push(["#", "Descripción", "Monto Acumulado ($CLP)", "Presupuesto ($CLP)", "Diferencia ($CLP)", "% Usado"]);
    resumen.egresos.forEach((f) => {
      filasResumen.push([f.n, f.nombre, f.acum, f.presu, f.diferencia, f.pct === null ? "—" : f.pct]);
    });
    filasResumen.push(["TOTAL EGRESOS", "", resumen.totalEgresosAcum, resumen.totalEgresosPresu, resumen.totalEgresosDif]);
    filasResumen.push([]);
    filasResumen.push(["INGRESOS"]);
    filasResumen.push(["#", "Descripción", "Monto Acumulado ($CLP)", "Presupuesto ($CLP)", "Diferencia ($CLP)", "% Usado"]);
    resumen.ingresos.forEach((f) => {
      filasResumen.push([f.n, f.nombre, f.acum, f.presu, f.diferencia, f.pct === null ? "—" : f.pct]);
    });
    filasResumen.push(["TOTAL INGRESOS", "", resumen.totalIngresosAcum, resumen.totalIngresosPresu, resumen.totalIngresosDif]);
    filasResumen.push([]);
    filasResumen.push(["RESULTADO NETO", "", resumen.resultadoNeto]);
    const hojaResumen = XLSX.utils.aoa_to_sheet(filasResumen);
    hojaResumen["!cols"] = [{ wch: 22 }, { wch: 26 }, { wch: 20 }, { wch: 18 }, { wch: 16 }, { wch: 10 }];

    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hojaRegistro, "Registro Diario");
    XLSX.utils.book_append_sheet(libro, hojaResumen, "Resumen Mensual");
    const buffer = XLSX.write(libro, { bookType: "xlsx", type: "array" });
    descargarArchivo(new Blob([buffer]), `finanzas_depto505_${monthLabel(cursor).replace(" ", "_")}.xlsx`, "application/octet-stream");
  };

  const exportarBackup = () => {
    const payload = JSON.stringify({ categorias, movimientos, exportadoEn: new Date().toISOString() }, null, 2);
    descargarArchivo(payload, `backup-depto505-${new Date().toISOString().slice(0, 10)}.json`, "application/json");
  };

  const importarBackup = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (!Array.isArray(data.categorias) || !Array.isArray(data.movimientos)) throw new Error("Formato inválido");
        guardarCategorias(data.categorias);
        guardarMovimientos(data.movimientos);
        setError(null);
      } catch {
        setError("El archivo de backup no tiene el formato esperado.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  if (categorias === null || movimientos === null) {
    return (
      <div style={{ ...styles.app, alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: "#8A94A8", fontFamily: "'IBM Plex Mono', monospace" }}>Cargando…</div>
      </div>
    );
  }

  const categoriasDelTipo = categorias.filter((c) => c.tipo === nuevo.tipo);

  return (
    <div style={styles.app}>
      <style>{FONT_IMPORT}</style>

      <header style={styles.header}>
        <div style={styles.headerTop}>
          <div style={styles.brand}>
            <Wallet size={18} color={styles.mint} strokeWidth={2.25} />
            <span style={styles.brandText}>Finanzas · Depto 505</span>
          </div>
          <div style={styles.balanceWrap}>
            <span style={styles.balanceLabel}>Resultado neto — {monthLabel(cursor)}</span>
            <span style={{ ...styles.balanceNum, color: resumen.resultadoNeto >= 0 ? styles.mint : styles.coral }}>
              {fmtCLP(resumen.resultadoNeto)}
            </span>
          </div>
        </div>
        <div style={styles.ledgerRule}>
          {Array.from({ length: 40 }).map((_, i) => <span key={i} style={styles.tick} />)}
        </div>
      </header>

      <nav style={styles.tabs}>
        {[
          { id: "registro", label: "Registro diario", icon: ListChecks },
          { id: "resumen", label: "Resumen mensual", icon: LayoutGrid },
          { id: "categorias", label: "Categorías" },
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            style={{ ...styles.tabBtn, ...(tab === t.id ? styles.tabBtnActive : {}) }}
          >
            {t.icon ? <t.icon size={13} style={{ marginRight: 5, verticalAlign: -2 }} /> : null}
            {t.label}
          </button>
        ))}
        <div style={styles.monthNav}>
          <button
            style={{ ...styles.iconBtn, ...(mostrarPegar ? { background: "#1C2740", color: "#E7ECF5" } : {}) }}
            onClick={() => setMostrarPegar((v) => !v)}
            title="Pegar filas copiadas desde Excel"
            aria-label="Pegar desde Excel"
          >
            <ClipboardPaste size={15} />
          </button>
          <button
            style={{ ...styles.iconBtn, color: styles.coral }}
            onClick={vaciarMesActual}
            title={`Borrar todos los movimientos de ${monthLabel(cursor)}`}
            aria-label="Borrar movimientos del mes"
          >
            <Trash2 size={15} />
          </button>
          <button style={styles.iconBtn} onClick={exportarExcel} title="Exportar a Excel (estilo original)" aria-label="Exportar a Excel"><FileSpreadsheet size={15} /></button>
          <button style={styles.iconBtn} onClick={exportarBackup} title="Descargar backup (JSON)" aria-label="Descargar backup"><Download size={15} /></button>
          <button style={styles.iconBtn} onClick={() => fileInputRef.current?.click()} title="Restaurar backup (JSON)" aria-label="Restaurar backup"><Upload size={15} /></button>
          <input ref={fileInputRef} type="file" accept="application/json" onChange={importarBackup} style={{ display: "none" }} />
          <span style={{ width: 8 }} />
          <button style={styles.iconBtn} onClick={() => moverMes(-1)} aria-label="Mes anterior"><ChevronLeft size={16} /></button>
          <span style={styles.monthLabel}>{monthLabel(cursor)}</span>
          <button style={styles.iconBtn} onClick={() => moverMes(1)} aria-label="Mes siguiente"><ChevronRight size={16} /></button>
        </div>
      </nav>

      {error && <div style={styles.errorBanner}>{error}</div>}

      {mostrarPegar && (
        <section style={styles.pegarPanel}>
          <h3 style={styles.sectionTitle}>Pegar filas desde Excel</h3>
          <p style={styles.hint}>
            Copia las filas de tu planilla (fecha, categoría, descripción, monto — separadas por tabulación,
            tal como las copia Excel) y pégalas acá abajo.
          </p>
          <textarea
            value={textoPegado}
            onChange={(e) => setTextoPegado(e.target.value)}
            placeholder={"01/07/2026\tAseos externos - Peti\taseo\t15,000\n02/07/2026\tIngreso 1\tairbnb sur villa\t72571"}
            style={styles.textarea}
            rows={6}
          />
          <div style={{ display: "flex", gap: 8 }}>
            <button style={styles.addBtn} onClick={procesarTextoPegado} disabled={!textoPegado.trim()}>Revisar filas</button>
            <button style={styles.secondaryBtn} onClick={cancelarPegado}>Cancelar</button>
          </div>

          {pegarResultado && (
            <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 4 }}>
              <div style={styles.pegarResumen}>
                <span style={{ color: styles.mint }}>{pegarResultado.filasValidas.length} movimiento(s) listos para agregar</span>
                {pegarResultado.filasInvalidas.length > 0 && (
                  <span style={{ color: styles.coral }}>{pegarResultado.filasInvalidas.length} fila(s) omitida(s)</span>
                )}
              </div>

              {pegarResultado.categoriasNuevas.length > 0 && (
                <div>
                  <p style={styles.hint}>Estas categorías no existían — elige si son gasto o ingreso:</p>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {pegarResultado.categoriasNuevas.map((nombre) => (
                      <div key={nombre} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <span style={{ fontSize: 13, minWidth: 180 }}>{nombre}</span>
                        <div style={styles.segmented}>
                          {["gasto", "ingreso"].map((t) => (
                            <button
                              type="button"
                              key={t}
                              onClick={() => setTiposCategoriasNuevas((f) => ({ ...f, [nombre]: t }))}
                              style={{ ...styles.segmentBtn, ...(tiposCategoriasNuevas[nombre] === t ? { background: t === "gasto" ? styles.coral : styles.mint, color: "#0B1120" } : {}) }}
                            >
                              {t === "gasto" ? "Gasto" : "Ingreso"}
                            </button>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {pegarResultado.filasInvalidas.length > 0 && (
                <div>
                  <p style={styles.hint}>Filas omitidas (revisa formato de fecha o monto):</p>
                  <ul style={{ margin: 0, paddingLeft: 18, fontSize: 12, color: "#8A94A8" }}>
                    {pegarResultado.filasInvalidas.map((f) => (
                      <li key={f.n}>Línea {f.n} — {f.motivo}: <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{f.texto}</span></li>
                    ))}
                  </ul>
                </div>
              )}

              <div style={{ display: "flex", gap: 8 }}>
                <button style={styles.addBtn} onClick={confirmarPegado} disabled={pegarResultado.filasValidas.length === 0}>
                  <Check size={16} /> Confirmar e importar
                </button>
                <button style={styles.secondaryBtn} onClick={cancelarPegado}>Cancelar</button>
              </div>
            </div>
          )}
        </section>
      )}

      {tab === "registro" && (
        <main style={styles.main}>
          <section style={styles.cardsRow}>
            <div style={styles.card}>
              <div style={styles.cardIconRow}><TrendingUp size={16} color={styles.mint} /><span style={styles.cardLabel}>Ingresos del mes</span></div>
              <div style={{ ...styles.cardNum, color: styles.mint }}>{fmtCLP(resumen.totalIngresosAcum)}</div>
            </div>
            <div style={styles.card}>
              <div style={styles.cardIconRow}><TrendingDown size={16} color={styles.coral} /><span style={styles.cardLabel}>Egresos del mes</span></div>
              <div style={{ ...styles.cardNum, color: styles.coral }}>{fmtCLP(resumen.totalEgresosAcum)}</div>
            </div>
            <div style={styles.card}>
              <div style={styles.cardIconRow}><Wallet size={16} color={styles.textPrimary.color} /><span style={styles.cardLabel}>Neto del mes</span></div>
              <div style={{ ...styles.cardNum, color: resumen.resultadoNeto >= 0 ? styles.mint : styles.coral }}>{fmtCLP(resumen.resultadoNeto)}</div>
            </div>
          </section>

          <section>
            <h3 style={styles.sectionTitle}>Nuevo movimiento</h3>
            <form onSubmit={agregarMovimiento} style={styles.form}>
              <div style={styles.formRow}>
                <div style={styles.segmented}>
                  {["gasto", "ingreso"].map((t) => (
                    <button
                      type="button"
                      key={t}
                      onClick={() => {
                        const primera = categorias.find((c) => c.tipo === t);
                        setNuevo((f) => ({ ...f, tipo: t, categoriaId: primera ? primera.id : "" }));
                      }}
                      style={{ ...styles.segmentBtn, ...(nuevo.tipo === t ? { background: t === "gasto" ? styles.coral : styles.mint, color: "#0B1120" } : {}) }}
                    >
                      {t === "gasto" ? "Gasto" : "Ingreso"}
                    </button>
                  ))}
                </div>
                <select
                  value={nuevo.categoriaId}
                  onChange={(e) => setNuevo((f) => ({ ...f, categoriaId: e.target.value }))}
                  style={{ ...styles.inputDesc, minWidth: 190, fontFamily: "'Inter', sans-serif" }}
                >
                  {categoriasDelTipo.length === 0 && <option value="">Sin categorías — crea una</option>}
                  {categoriasDelTipo.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                </select>
                <input
                  type="text" inputMode="decimal" placeholder="Monto (ej: 15.000)"
                  value={nuevo.monto}
                  onChange={(e) => setNuevo((f) => ({ ...f, monto: e.target.value }))}
                  style={styles.inputNum} required
                />
                <input
                  type="date"
                  value={nuevo.fecha}
                  onChange={(e) => setNuevo((f) => ({ ...f, fecha: e.target.value }))}
                  style={styles.inputDate} required
                />
                <input
                  type="text" placeholder="Descripción / boleta (opcional)"
                  value={nuevo.descripcion}
                  onChange={(e) => setNuevo((f) => ({ ...f, descripcion: e.target.value }))}
                  style={styles.inputDesc}
                />
                <button type="submit" style={styles.addBtn}><Plus size={16} /> Agregar</button>
              </div>
            </form>
          </section>

          <section style={styles.tableWrap}>
            {movsDelMes.length === 0 ? (
              <EmptyState texto="Aún no hay movimientos este mes." />
            ) : (
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Fecha</th><th style={styles.th}>Categoría</th>
                    <th style={styles.th}>Descripción / boleta</th><th style={{ ...styles.th, textAlign: "right" }}>Monto</th><th style={styles.th}></th>
                  </tr>
                </thead>
                <tbody>
                  {movsDelMes.slice().sort((a, b) => b.fecha.localeCompare(a.fecha)).map((m) => {
                    const c = catMap[m.categoriaId];
                    const accent = m.tipo === "ingreso" ? styles.mint : styles.coral;
                    const enEdicion = editMovId === m.id;

                    if (enEdicion) {
                      return (
                        <tr key={m.id}>
                          <td style={styles.td}>
                            <input type="date" value={editMovDraft.fecha} onChange={(e) => setEditMovDraft((f) => ({ ...f, fecha: e.target.value }))} style={styles.inputMini} />
                          </td>
                          <td style={styles.td}>
                            <select
                              value={editMovDraft.categoriaId}
                              onChange={(e) => setEditMovDraft((f) => ({ ...f, categoriaId: e.target.value }))}
                              style={{ ...styles.inputMini, fontFamily: "'Inter', sans-serif" }}
                            >
                              {categorias.map((cc) => <option key={cc.id} value={cc.id}>{cc.nombre}</option>)}
                            </select>
                          </td>
                          <td style={styles.td}>
                            <input type="text" value={editMovDraft.descripcion} onChange={(e) => setEditMovDraft((f) => ({ ...f, descripcion: e.target.value }))} style={styles.inputMini} />
                          </td>
                          <td style={{ ...styles.td, textAlign: "right" }}>
                            <input
                              type="text" inputMode="decimal"
                              value={editMovDraft.monto}
                              onChange={(e) => setEditMovDraft((f) => ({ ...f, monto: e.target.value }))}
                              style={{ ...styles.inputMini, textAlign: "right" }}
                              autoFocus
                              onKeyDown={(e) => e.key === "Enter" && guardarEdicionMovimiento()}
                            />
                          </td>
                          <td style={{ ...styles.td, textAlign: "right" }}>
                            <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                              <button onClick={guardarEdicionMovimiento} style={styles.miniIconBtn} aria-label="Guardar"><Check size={12} /></button>
                              <button onClick={cancelarEdicionMovimiento} style={styles.miniIconBtn} aria-label="Cancelar"><X size={12} /></button>
                            </div>
                          </td>
                        </tr>
                      );
                    }

                    return (
                      <tr key={m.id}>
                        <td style={styles.td}>{m.fecha}</td>
                        <td style={styles.td}>
                          <span style={{ ...styles.badge, color: accent, borderColor: accent }}>{c ? c.nombre : "—"}</span>
                        </td>
                        <td style={{ ...styles.td, color: "#8A94A8" }}>{m.descripcion || "—"}</td>
                        <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: accent }}>
                          {m.tipo === "ingreso" ? "+" : "−"}{fmtCLP(m.monto)}
                        </td>
                        <td style={{ ...styles.td, textAlign: "right" }}>
                          <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                            <button onClick={() => empezarEdicionMovimiento(m)} style={styles.deleteBtn} aria-label="Editar"><Pencil size={14} /></button>
                            <button onClick={() => eliminarMovimiento(m.id)} style={styles.deleteBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </section>
        </main>
      )}

      {tab === "resumen" && (
        <main style={styles.main}>
          <ResumenTabla
            titulo="▼  Egresos"
            filas={resumen.egresos}
            tipo="gasto"
            accent={styles.coral}
            totalAcum={resumen.totalEgresosAcum}
            totalPresu={resumen.totalEgresosPresu}
            totalDif={resumen.totalEgresosDif}
            editPresId={editPresId}
            editPresVal={editPresVal}
            setEditPresId={setEditPresId}
            setEditPresVal={setEditPresVal}
            actualizarPresupuesto={actualizarPresupuesto}
          />
          <ResumenTabla
            titulo="▲  Ingresos"
            filas={resumen.ingresos}
            tipo="ingreso"
            accent={styles.mint}
            totalAcum={resumen.totalIngresosAcum}
            totalPresu={resumen.totalIngresosPresu}
            totalDif={resumen.totalIngresosDif}
            editPresId={editPresId}
            editPresVal={editPresVal}
            setEditPresId={setEditPresId}
            setEditPresVal={setEditPresVal}
            actualizarPresupuesto={actualizarPresupuesto}
          />
          <div style={{ ...styles.card, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ ...styles.sectionTitle }}>✅  Resultado neto — {monthLabel(cursor)}</span>
            <span style={{ ...styles.cardNum, fontSize: 24, color: resumen.resultadoNeto >= 0 ? styles.mint : styles.coral }}>
              {fmtCLP(resumen.resultadoNeto)}
            </span>
          </div>
          <p style={styles.hint}>El presupuesto de cada categoría se edita haciendo clic en el ícono de lápiz. Los montos acumulados se calculan automáticamente desde el Registro diario de {monthLabel(cursor)}.</p>
        </main>
      )}

      {tab === "categorias" && (
        <main style={styles.main}>
          <h3 style={styles.sectionTitle}>Crear nueva categoría</h3>
          <form onSubmit={crearCategoria} style={styles.form}>
            <div style={styles.formRow}>
              <div style={styles.segmented}>
                {["gasto", "ingreso"].map((t) => (
                  <button type="button" key={t} onClick={() => setNuevaCat((f) => ({ ...f, tipo: t }))}
                    style={{ ...styles.segmentBtn, ...(nuevaCat.tipo === t ? { background: t === "gasto" ? styles.coral : styles.mint, color: "#0B1120" } : {}) }}>
                    {t === "gasto" ? "Gasto" : "Ingreso"}
                  </button>
                ))}
              </div>
              <input
                type="text" placeholder="Nombre (ej: Gas, Mantención, Airbnb sur...)"
                value={nuevaCat.nombre}
                onChange={(e) => setNuevaCat((f) => ({ ...f, nombre: e.target.value }))}
                style={{ ...styles.inputDesc, minWidth: 200 }} required
              />
              <input
                type="text" inputMode="decimal" placeholder="Presupuesto (ej: 25.000, opcional)"
                value={nuevaCat.presupuesto}
                onChange={(e) => setNuevaCat((f) => ({ ...f, presupuesto: e.target.value }))}
                style={{ ...styles.inputNum, width: 170 }}
              />
              <button type="submit" style={styles.addBtn}><Plus size={16} /> Crear</button>
            </div>
          </form>

          <h3 style={{ ...styles.sectionTitle, marginTop: 20 }}>Categorías actuales</h3>
          <p style={styles.hint}>Edita el nombre o el presupuesto de cada categoría. Los cambios aplican desde este momento hacia adelante.</p>
          <div style={styles.tableWrap}>
            <table style={styles.table}>
              <thead>
                <tr><th style={styles.th}>Nombre</th><th style={styles.th}>Tipo</th><th style={{ ...styles.th, textAlign: "right" }}>Presupuesto</th><th style={styles.th}></th></tr>
              </thead>
              <tbody>
                {categorias.map((c) => {
                  const accent = c.tipo === "ingreso" ? styles.mint : styles.coral;
                  return (
                    <tr key={c.id}>
                      <td style={styles.td}>
                        {editNombreId === c.id ? (
                          <div style={{ display: "flex", gap: 6 }}>
                            <input
                              type="text"
                              value={editNombreVal}
                              onChange={(e) => setEditNombreVal(e.target.value)}
                              style={styles.inputMini}
                              autoFocus
                              onKeyDown={(e) => e.key === "Enter" && renombrarCategoria(c.id, editNombreVal)}
                            />
                            <button style={styles.miniIconBtn} onClick={() => renombrarCategoria(c.id, editNombreVal)} aria-label="Guardar nombre"><Check size={12} /></button>
                            <button style={styles.miniIconBtn} onClick={() => setEditNombreId(null)} aria-label="Cancelar"><X size={12} /></button>
                          </div>
                        ) : (
                          c.nombre
                        )}
                      </td>
                      <td style={styles.td}>
                        <span style={{ ...styles.badge, color: accent, borderColor: accent }}>
                          {c.tipo === "ingreso" ? "Ingreso" : "Gasto"}
                        </span>
                      </td>
                      <td style={{ ...styles.td, textAlign: "right" }}>
                        {editPresId === c.id ? (
                          <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                            <input
                              type="text" inputMode="decimal"
                              value={editPresVal}
                              onChange={(e) => setEditPresVal(e.target.value)}
                              style={{ ...styles.inputMini, width: 100 }}
                              autoFocus
                              onKeyDown={(e) => e.key === "Enter" && actualizarPresupuesto(c.id, editPresVal)}
                            />
                            <button style={styles.miniIconBtn} onClick={() => actualizarPresupuesto(c.id, editPresVal)} aria-label="Guardar presupuesto"><Check size={12} /></button>
                          </div>
                        ) : (
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtCLP(c.presupuesto)}</span>
                        )}
                      </td>
                      <td style={{ ...styles.td, textAlign: "right" }}>
                        <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                          <button onClick={() => { setEditNombreId(c.id); setEditNombreVal(c.nombre); }} style={styles.deleteBtn} aria-label="Editar nombre"><Pencil size={14} /></button>
                          <button onClick={() => { setEditPresId(c.id); setEditPresVal(String(c.presupuesto || "")); }} style={styles.deleteBtn} aria-label="Editar presupuesto"><FileSpreadsheet size={14} /></button>
                          <button onClick={() => eliminarCategoria(c.id)} style={styles.deleteBtn} aria-label="Eliminar"><Trash2 size={14} /></button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div style={styles.zonaPeligro}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <AlertTriangle size={16} color={styles.coral} />
              <span style={{ ...styles.sectionTitle, color: styles.coral }}>Zona de peligro</span>
            </div>
            <p style={styles.hint}>
              Borra todos los movimientos registrados, de todos los meses (útil si pegaste el mismo bloque de Excel
              varias veces por error). Las categorías y sus presupuestos no se tocan.
            </p>
            <button style={styles.dangerBtn} onClick={borrarTodosLosMovimientos}>
              <Trash2 size={14} /> Borrar todos los movimientos ({(movimientos || []).length})
            </button>
          </div>
        </main>
      )}

      {confirmAccion && (
        <div style={styles.modalOverlay} onClick={() => setConfirmAccion(null)}>
          <div style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <AlertTriangle size={18} color={styles.coral} />
              <span style={{ ...styles.sectionTitle, fontSize: 15 }}>Confirmar borrado</span>
            </div>
            <p style={{ fontSize: 13, color: "#E7ECF5", margin: 0, lineHeight: 1.5 }}>{confirmAccion.mensaje}</p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button style={styles.secondaryBtn} onClick={() => setConfirmAccion(null)}>Cancelar</button>
              <button style={styles.dangerBtn} onClick={ejecutarConfirmAccion}>
                <Trash2 size={14} /> Sí, borrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ResumenTabla({ titulo, filas, accent, totalAcum, totalPresu, totalDif, editPresId, editPresVal, setEditPresId, setEditPresVal, actualizarPresupuesto }) {
  return (
    <section style={styles.tableWrap}>
      <div style={{ padding: "12px 14px", borderBottom: "1px solid #232E47" }}>
        <span style={{ ...styles.sectionTitle, color: accent }}>{titulo}</span>
      </div>
      <table style={styles.table}>
        <thead>
          <tr>
            <th style={{ ...styles.th, width: 34 }}>#</th>
            <th style={styles.th}>Descripción</th>
            <th style={{ ...styles.th, textAlign: "right" }}>Acumulado</th>
            <th style={{ ...styles.th, textAlign: "right" }}>Presupuesto</th>
            <th style={{ ...styles.th, textAlign: "right" }}>Diferencia</th>
            <th style={{ ...styles.th, textAlign: "right" }}>% usado</th>
            <th style={styles.th}></th>
          </tr>
        </thead>
        <tbody>
          {filas.map((f) => (
            <tr key={f.id}>
              <td style={{ ...styles.td, color: "#8A94A8" }}>{f.n}</td>
              <td style={styles.td}>{f.nombre}</td>
              <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{fmtCLP(f.acum)}</td>
              <td style={{ ...styles.td, textAlign: "right" }}>
                {editPresId === f.id ? (
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <input
                      type="text" inputMode="decimal"
                      value={editPresVal}
                      onChange={(e) => setEditPresVal(e.target.value)}
                      style={{ ...styles.inputMini, width: 100 }}
                      autoFocus
                      onKeyDown={(e) => e.key === "Enter" && actualizarPresupuesto(f.id, editPresVal)}
                    />
                    <button style={styles.miniIconBtn} onClick={() => actualizarPresupuesto(f.id, editPresVal)} aria-label="Guardar"><Check size={12} /></button>
                  </div>
                ) : (
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtCLP(f.presu)}</span>
                )}
              </td>
              <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: f.diferencia >= 0 ? styles.mint : styles.coral }}>
                {f.diferencia >= 0 ? "+" : "−"}{fmtCLP(Math.abs(f.diferencia))}
              </td>
              <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{fmtPct(f.pct)}</td>
              <td style={{ ...styles.td, textAlign: "right" }}>
                {editPresId === f.id ? null : (
                  <button style={styles.deleteBtn} onClick={() => { setEditPresId(f.id); setEditPresVal(String(f.presu || "")); }} aria-label="Editar presupuesto"><Pencil size={13} /></button>
                )}
              </td>
            </tr>
          ))}
          <tr>
            <td style={styles.td}></td>
            <td style={{ ...styles.td, fontWeight: 600 }}>TOTAL</td>
            <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{fmtCLP(totalAcum)}</td>
            <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{fmtCLP(totalPresu)}</td>
            <td style={{ ...styles.td, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: totalDif >= 0 ? styles.mint : styles.coral }}>
              {totalDif >= 0 ? "+" : "−"}{fmtCLP(Math.abs(totalDif))}
            </td>
            <td style={styles.td}></td>
            <td style={styles.td}></td>
          </tr>
        </tbody>
      </table>
    </section>
  );
}

function EmptyState({ texto }) {
  return <div style={styles.empty}>{texto}</div>;
}

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');`;

const mint = "#34D9B4";
const coral = "#FF8A5B";

const styles = {
  mint, coral,
  textPrimary: { color: "#E7ECF5" },
  app: { display: "flex", flexDirection: "column", minHeight: "100vh", background: "#0B1120", color: "#E7ECF5", fontFamily: "'Inter', sans-serif" },
  header: { padding: "20px 20px 0 20px" },
  headerTop: { display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 12 },
  brand: { display: "flex", alignItems: "center", gap: 8 },
  brandText: { fontFamily: "'Space Grotesk', sans-serif", fontSize: 14, fontWeight: 600, letterSpacing: 0.2, color: "#8A94A8" },
  balanceWrap: { display: "flex", flexDirection: "column", alignItems: "flex-end" },
  balanceLabel: { fontSize: 11, color: "#8A94A8", textTransform: "uppercase", letterSpacing: 1 },
  balanceNum: { fontFamily: "'Space Grotesk', sans-serif", fontSize: 30, fontWeight: 700, lineHeight: 1.1, textTransform: "capitalize" },
  ledgerRule: { display: "flex", justifyContent: "space-between", marginTop: 14, borderTop: "1px solid #232E47", paddingTop: 6 },
  tick: { width: 1, height: 5, background: "#232E47" },
  tabs: { display: "flex", alignItems: "center", gap: 4, padding: "14px 20px", borderBottom: "1px solid #131B2E", flexWrap: "wrap" },
  tabBtn: { background: "transparent", border: "none", color: "#8A94A8", fontFamily: "'Space Grotesk', sans-serif", fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, cursor: "pointer" },
  tabBtnActive: { background: "#1C2740", color: "#E7ECF5" },
  monthNav: { marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 },
  monthLabel: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: "#E7ECF5", minWidth: 90, textAlign: "center", textTransform: "capitalize" },
  iconBtn: { background: "#131B2E", border: "1px solid #232E47", color: "#8A94A8", borderRadius: 6, width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  errorBanner: { margin: "0 20px", marginTop: 12, padding: "8px 12px", background: "#2A1620", border: "1px solid #FF8A5B", color: "#FF8A5B", borderRadius: 8, fontSize: 12 },
  pegarPanel: { margin: "12px 20px 0 20px", padding: 16, background: "#131B2E", border: "1px solid #232E47", borderRadius: 12, display: "flex", flexDirection: "column", gap: 10 },
  textarea: { background: "#0B1120", color: "#E7ECF5", border: "1px solid #232E47", borderRadius: 8, padding: "10px 12px", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace", width: "100%", resize: "vertical" },
  secondaryBtn: { background: "transparent", color: "#8A94A8", border: "1px solid #232E47", borderRadius: 8, padding: "9px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "'Space Grotesk', sans-serif" },
  pegarResumen: { display: "flex", gap: 16, fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" },
  zonaPeligro: { border: "1px solid #4A2A2A", background: "#1A1315", borderRadius: 12, padding: 16, display: "flex", flexDirection: "column", gap: 8 },
  dangerBtn: { display: "flex", alignItems: "center", gap: 6, alignSelf: "flex-start", background: "transparent", color: "#FF8A5B", border: "1px solid #FF8A5B", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "'Space Grotesk', sans-serif" },
  modalOverlay: { position: "fixed", inset: 0, background: "rgba(11,17,32,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 },
  modalCard: { background: "#131B2E", border: "1px solid #232E47", borderRadius: 12, padding: 20, display: "flex", flexDirection: "column", gap: 14, maxWidth: 360, width: "100%" },
  main: { padding: 20, display: "flex", flexDirection: "column", gap: 20 },
  cardsRow: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12 },
  card: { background: "#131B2E", border: "1px solid #232E47", borderRadius: 12, padding: 16 },
  cardIconRow: { display: "flex", alignItems: "center", gap: 6, marginBottom: 10 },
  cardLabel: { fontSize: 12, color: "#8A94A8" },
  cardNum: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 500 },
  sectionTitle: { fontFamily: "'Space Grotesk', sans-serif", fontSize: 14, fontWeight: 600, margin: 0 },
  hint: { fontSize: 12, color: "#8A94A8", margin: 0 },
  form: { background: "#131B2E", border: "1px solid #232E47", borderRadius: 12, padding: 14 },
  formRow: { display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" },
  segmented: { display: "flex", background: "#0B1120", borderRadius: 8, padding: 2, border: "1px solid #232E47" },
  segmentBtn: { border: "none", background: "transparent", color: "#8A94A8", fontSize: 12, fontWeight: 600, padding: "7px 12px", borderRadius: 6, cursor: "pointer", fontFamily: "'Space Grotesk', sans-serif" },
  inputNum: { background: "#0B1120", color: "#E7ECF5", border: "1px solid #232E47", borderRadius: 8, padding: "8px 10px", fontSize: 12, width: 110, fontFamily: "'IBM Plex Mono', monospace" },
  inputDate: { background: "#0B1120", color: "#E7ECF5", border: "1px solid #232E47", borderRadius: 8, padding: "8px 10px", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" },
  inputDesc: { background: "#0B1120", color: "#E7ECF5", border: "1px solid #232E47", borderRadius: 8, padding: "8px 10px", fontSize: 12, flex: 1, minWidth: 140, fontFamily: "'Inter', sans-serif" },
  inputMini: { background: "#0B1120", color: "#E7ECF5", border: "1px solid #232E47", borderRadius: 6, padding: "6px 8px", fontSize: 12, width: "100%", fontFamily: "'IBM Plex Mono', monospace" },
  miniIconBtn: { background: "transparent", border: "1px solid #232E47", borderRadius: 6, color: "#8A94A8", width: 22, height: 22, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  addBtn: { display: "flex", alignItems: "center", gap: 6, background: "#34D9B4", color: "#0B1120", border: "none", borderRadius: 8, padding: "9px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "'Space Grotesk', sans-serif" },
  tableWrap: { background: "#131B2E", border: "1px solid #232E47", borderRadius: 12, overflow: "hidden" },
  table: { width: "100%", borderCollapse: "collapse" },
  th: { textAlign: "left", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: "#8A94A8", padding: "10px 14px", borderBottom: "1px solid #232E47" },
  td: { padding: "10px 14px", fontSize: 13, color: "#E7ECF5", borderBottom: "1px solid #1C2740" },
  badge: { fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 999, border: "1px solid", fontFamily: "'IBM Plex Mono', monospace" },
  deleteBtn: { background: "transparent", border: "none", color: "#8A94A8", cursor: "pointer", padding: 4 },
  empty: { color: "#8A94A8", fontSize: 13, padding: "40px 0", textAlign: "center" },
};