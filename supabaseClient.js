import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  Calendar, Plus, Phone, MapPin, DollarSign, Check, Clock,
  Users, FileSpreadsheet, Trash2, Edit2, X, Upload, ChevronLeft, ChevronRight,
  Sparkles, CircleDot, CheckCircle2
} from "lucide-react";

/* ---------- estilos / tokens ---------- */
const Styles = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400..700&family=Inter:wght@400;500;600;700&display=swap');
    :root{
      --teal:#1B4B43;
      --mint:#8FE3C8;
      --mint-light:#E4F8F0;
      --cream:#FAF6F0;
      --charcoal:#26312E;
      --coral:#FF6B5C;
      --gold:#F4B860;
    }
    .sp-root{font-family:'Inter',sans-serif;background:var(--cream);color:var(--charcoal);min-height:100vh;}
    .sp-display{font-family:'Fraunces',serif;}
    .sp-swoosh{display:block;width:100%;height:34px;}
    .sp-card{background:#fff;border-radius:16px;box-shadow:0 1px 3px rgba(27,75,67,0.08);}
    .sp-btn{font-family:'Inter',sans-serif;font-weight:600;border-radius:12px;border:none;cursor:pointer;transition:transform .08s ease, opacity .15s ease;}
    .sp-btn:active{transform:scale(0.97);}
    .sp-btn-primary{background:var(--teal);color:#fff;}
    .sp-btn-primary:hover{opacity:.9;}
    .sp-btn-ghost{background:var(--mint-light);color:var(--teal);}
    .sp-btn-coral{background:var(--coral);color:#fff;}
    .sp-input{font-family:'Inter',sans-serif;border:1.5px solid #E2DCD3;border-radius:10px;padding:10px 12px;width:100%;background:#fff;font-size:14px;color:var(--charcoal);}
    .sp-input:focus{outline:none;border-color:var(--mint);}
    .sp-label{font-size:12px;font-weight:600;color:#7A8B85;text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px;display:block;}
    .sp-tag{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:600;}
    .sp-tag-pendiente{background:#FFE9E5;color:#C84A3B;}
    .sp-tag-completado{background:#E2F6EC;color:#1B4B43;}
    .sp-navbtn{display:flex;flex-direction:column;align-items:center;gap:3px;font-size:11px;font-weight:600;color:#A6B5AF;background:none;border:none;cursor:pointer;padding:6px 4px;flex:1;}
    .sp-navbtn.active{color:var(--teal);}
    .sp-fab{position:fixed;bottom:78px;right:18px;width:56px;height:56px;border-radius:50%;background:var(--coral);color:#fff;display:flex;align-items:center;justify-content:center;box-shadow:0 6px 16px rgba(255,107,92,0.4);border:none;cursor:pointer;z-index:30;}
    .sp-bottomnav{position:fixed;bottom:0;left:0;right:0;background:#fff;border-top:1px solid #EFE8DE;display:flex;padding:6px 4px 10px;z-index:30;}
    .sp-scroll::-webkit-scrollbar{display:none;}
    .sp-modalbg{position:fixed;inset:0;background:rgba(38,49,46,0.45);display:flex;align-items:flex-end;justify-content:center;z-index:50;}
    @media (min-width:640px){.sp-modalbg{align-items:center;}}
    .sp-modal{background:#fff;border-radius:20px 20px 0 0;width:100%;max-width:480px;max-height:88vh;overflow-y:auto;padding:24px;}
    @media (min-width:640px){.sp-modal{border-radius:20px;}}
  `}</style>
);

/* ---------- estela / signature element ---------- */
const Swoosh = ({ flip }) => (
  <svg className="sp-swoosh" viewBox="0 0 400 34" preserveAspectRatio="none" style={flip ? { transform: "scaleY(-1)" } : {}}>
    <defs>
      <linearGradient id="swooshGrad" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stopColor="#8FE3C8" stopOpacity="0" />
        <stop offset="50%" stopColor="#8FE3C8" stopOpacity="0.9" />
        <stop offset="100%" stopColor="#1B4B43" stopOpacity="0" />
      </linearGradient>
    </defs>
    <path d="M0,28 C100,4 180,32 200,16 C220,2 300,26 400,8" fill="none" stroke="url(#swooshGrad)" strokeWidth="3" strokeLinecap="round" />
    <circle cx="60" cy="14" r="2" fill="#F4B860" />
    <circle cx="320" cy="18" r="2.5" fill="#FF6B5C" />
  </svg>
);

/* ---------- utilidades ---------- */
const uid = () => Math.random().toString(36).slice(2, 10);

const fmtFecha = (iso) => {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" });
};

const fmtHora12 = (hhmm) => {
  if (!hhmm) return "";
  let [h, m] = hhmm.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return `${h}:${String(m).padStart(2, "0")} ${ampm}`;
};

const addHoras = (hhmm, dur) => {
  if (!hhmm) return "";
  let [h, m] = hhmm.split(":").map(Number);
  let totalMin = h * 60 + m + dur * 60;
  totalMin = ((totalMin % (24 * 60)) + 24 * 60) % (24 * 60);
  const nh = Math.floor(totalMin / 60);
  const nm = totalMin % 60;
  return fmtHora12(`${String(nh).padStart(2, "0")}:${String(nm).padStart(2, "0")}`);
};

const mapsUrl = (direccion) =>
  `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(direccion)}`;

const telUrl = (tel) => `tel:${(tel || "").replace(/[^\d+]/g, "")}`;

const todayISO = () => new Date().toISOString().slice(0, 10);

const PAGOS = ["Cash", "Zelle", "Cheque"];

/* rangos de fecha para reportes */
const getRango = (periodo, ref) => {
  const d = new Date(ref + "T00:00:00");
  let start, end;
  if (periodo === "semana") {
    const dow = d.getDay();
    start = new Date(d); start.setDate(d.getDate() - dow);
    end = new Date(start); end.setDate(start.getDate() + 6);
  } else if (periodo === "mes") {
    start = new Date(d.getFullYear(), d.getMonth(), 1);
    end = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  } else if (periodo === "semestre") {
    const semStart = d.getMonth() < 6 ? 0 : 6;
    start = new Date(d.getFullYear(), semStart, 1);
    end = new Date(d.getFullYear(), semStart + 6, 0);
  } else {
    start = new Date(d.getFullYear(), 0, 1);
    end = new Date(d.getFullYear(), 11, 31);
  }
  return { start: start.toISOString().slice(0, 10), end: end.toISOString().slice(0, 10) };
};

/* ---------- App principal ---------- */
export default function App() {
  const [clientes, setClientes] = useState([]);
  const [trabajos, setTrabajos] = useState([]);
  const [agendadores, setAgendadores] = useState(["Esposa", "Juan Carlos"]);
  const [vista, setVista] = useState("agenda");
  const [modal, setModal] = useState(null); // {tipo: 'cliente'|'trabajo', data?}
  const [agendaFiltro, setAgendaFiltro] = useState("proximos"); // proximos | todos | pendientes
  const [busquedaCliente, setBusquedaCliente] = useState("");

  const [cargando, setCargando] = useState(true);

  /* mapeo db <-> app */
  const trabajoDesdeDb = (r) => ({
    id: r.id, clienteId: r.cliente_id, fecha: r.fecha, hora: r.hora?.slice(0, 5),
    duracion: Number(r.duracion), valor: Number(r.valor), pago: r.pago,
    estado: r.estado, agendadoPor: r.agendado_por, notas: r.notas || ""
  });
  const trabajoHaciaDb = (t) => ({
    cliente_id: t.clienteId, fecha: t.fecha, hora: t.hora, duracion: t.duracion,
    valor: t.valor, pago: t.pago, estado: t.estado, agendado_por: t.agendadoPor, notas: t.notas || ""
  });

  /* cargar datos iniciales */
  useEffect(() => {
    const cargar = async () => {
      const [{ data: c }, { data: t }, { data: a }] = await Promise.all([
        supabase.from("clientes").select("*").order("nombre"),
        supabase.from("trabajos").select("*"),
        supabase.from("agendadores").select("*").order("nombre")
      ]);
      if (c) setClientes(c);
      if (t) setTrabajos(t.map(trabajoDesdeDb));
      if (a) setAgendadores(a.map(r => r.nombre));
      setCargando(false);
    };
    cargar();
  }, []);

  /* ---------- acciones clientes ---------- */
  const guardarCliente = async (data) => {
    const { id, ...resto } = data;
    if (id) {
      const { data: actualizado } = await supabase.from("clientes").update(resto).eq("id", id).select().single();
      if (actualizado) setClientes(prev => prev.map(c => c.id === id ? actualizado : c));
    } else {
      const { data: creado } = await supabase.from("clientes").insert(resto).select().single();
      if (creado) setClientes(prev => [...prev, creado]);
    }
    setModal(null);
  };
  const borrarCliente = async (id) => {
    if (trabajos.some(t => t.clienteId === id)) {
      if (!window.confirm("Este cliente tiene trabajos agendados. ¿Eliminar de todas formas? (los trabajos quedarán sin cliente)")) return;
    }
    await supabase.from("clientes").delete().eq("id", id);
    setClientes(prev => prev.filter(c => c.id !== id));
  };

  const importarClientes = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const nuevos = rows.map(r => {
          const get = (...keys) => {
            for (const k of keys) {
              const found = Object.keys(r).find(rk => rk.toLowerCase().trim() === k);
              if (found) return String(r[found]).trim();
            }
            return "";
          };
          return {
            nombre: get("nombre", "cliente", "name"),
            telefono: get("telefono", "teléfono", "phone", "tel"),
            direccion: get("direccion", "dirección", "address"),
            notas: get("notas", "notes")
          };
        }).filter(c => c.nombre);
        const { data: creados } = await supabase.from("clientes").insert(nuevos).select();
        if (creados) setClientes(prev => [...prev, ...creados]);
        alert(`Se importaron ${creados?.length || 0} clientes.`);
      } catch (err) {
        alert("No se pudo leer el archivo. Asegúrate que tenga columnas: nombre, telefono, direccion");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  /* ---------- acciones trabajos ---------- */
  const guardarTrabajo = async (data) => {
    const { id, ...resto } = data;
    const payload = trabajoHaciaDb(resto);
    if (id) {
      const { data: actualizado } = await supabase.from("trabajos").update(payload).eq("id", id).select().single();
      if (actualizado) setTrabajos(prev => prev.map(t => t.id === id ? trabajoDesdeDb(actualizado) : t));
    } else {
      const { data: creado } = await supabase.from("trabajos").insert(payload).select().single();
      if (creado) setTrabajos(prev => [...prev, trabajoDesdeDb(creado)]);
    }
    setModal(null);
  };
  const borrarTrabajo = async (id) => {
    await supabase.from("trabajos").delete().eq("id", id);
    setTrabajos(prev => prev.filter(t => t.id !== id));
  };
  const toggleEstado = async (id) => {
    const actual = trabajos.find(t => t.id === id);
    const nuevoEstado = actual.estado === "Completado" ? "Pendiente" : "Completado";
    await supabase.from("trabajos").update({ estado: nuevoEstado }).eq("id", id);
    setTrabajos(prev => prev.map(t => t.id === id ? { ...t, estado: nuevoEstado } : t));
  };
  const agregarAgendador = async (nombre) => {
    const { data: creado } = await supabase.from("agendadores").insert({ nombre }).select().single();
    if (creado) setAgendadores(prev => [...prev, creado.nombre]);
  };

  /* ---------- datos derivados ---------- */
  const clientesPorId = useMemo(() => {
    const m = {};
    clientes.forEach(c => m[c.id] = c);
    return m;
  }, [clientes]);

  const trabajosOrdenados = useMemo(() => {
    let arr = [...trabajos].sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
    const hoy = todayISO();
    if (agendaFiltro === "proximos") arr = arr.filter(t => t.fecha >= hoy);
    if (agendaFiltro === "pendientes") arr = arr.filter(t => t.estado === "Pendiente");
    return arr;
  }, [trabajos, agendaFiltro]);

  const trabajosPorDia = useMemo(() => {
    const grupos = {};
    trabajosOrdenados.forEach(t => {
      grupos[t.fecha] = grupos[t.fecha] || [];
      grupos[t.fecha].push(t);
    });
    return grupos;
  }, [trabajosOrdenados]);

  const clientesFiltrados = useMemo(() => {
    const q = busquedaCliente.toLowerCase();
    return clientes.filter(c => c.nombre.toLowerCase().includes(q) || (c.telefono || "").includes(q));
  }, [clientes, busquedaCliente]);

  return (
    <div className="sp-root" style={{ paddingBottom: 90 }}>
      <Styles />
      <Header />
      <main style={{ maxWidth: 640, margin: "0 auto", padding: "0 16px" }}>
        {vista === "agenda" && (
          <VistaAgenda
            grupos={trabajosPorDia}
            clientesPorId={clientesPorId}
            filtro={agendaFiltro}
            setFiltro={setAgendaFiltro}
            onEditar={(t) => setModal({ tipo: "trabajo", data: t })}
            onBorrar={borrarTrabajo}
            onToggle={toggleEstado}
          />
        )}
        {vista === "clientes" && (
          <VistaClientes
            clientes={clientesFiltrados}
            busqueda={busquedaCliente}
            setBusqueda={setBusquedaCliente}
            onNuevo={() => setModal({ tipo: "cliente" })}
            onEditar={(c) => setModal({ tipo: "cliente", data: c })}
            onBorrar={borrarCliente}
            onImportar={importarClientes}
          />
        )}
        {vista === "reportes" && (
          <VistaReportes trabajos={trabajos} clientesPorId={clientesPorId} />
        )}
      </main>

      <button className="sp-fab" onClick={() => setModal({ tipo: "trabajo" })} aria-label="Agendar trabajo">
        <Plus size={26} />
      </button>

      <BottomNav vista={vista} setVista={setVista} />

      {modal?.tipo === "cliente" && (
        <ModalCliente data={modal.data} onGuardar={guardarCliente} onCerrar={() => setModal(null)} />
      )}
      {modal?.tipo === "trabajo" && (
        <ModalTrabajo
          data={modal.data}
          clientes={clientes}
          agendadores={agendadores}
          onAgregarAgendador={agregarAgendador}
          onGuardar={guardarTrabajo}
          onCrearClienteRapido={async (c) => {
            const { data: creado } = await supabase.from("clientes").insert(c).select().single();
            if (creado) setClientes(prev => [...prev, creado]);
            return creado;
          }}
          onCerrar={() => setModal(null)}
        />
      )}
    </div>
  );
}

/* ---------- header ---------- */
const Header = () => (
  <div style={{ background: "var(--teal)", color: "#fff", paddingTop: 22, paddingBottom: 6 }}>
    <div style={{ maxWidth: 640, margin: "0 auto", padding: "0 16px", display: "flex", alignItems: "center", gap: 10 }}>
      <div style={{ width: 38, height: 38, borderRadius: 10, background: "var(--mint)", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <Sparkles size={20} color="var(--teal)" />
      </div>
      <div>
        <div className="sp-display" style={{ fontSize: 22, fontWeight: 600, lineHeight: 1 }}>ShinyPro Clean</div>
        <div style={{ fontSize: 12, color: "var(--mint)", letterSpacing: ".03em" }}>agenda &amp; control de trabajos</div>
      </div>
    </div>
    <Swoosh />
  </div>
);

/* ---------- navegación inferior ---------- */
const BottomNav = ({ vista, setVista }) => (
  <nav className="sp-bottomnav">
    <button className={`sp-navbtn ${vista === "agenda" ? "active" : ""}`} onClick={() => setVista("agenda")}>
      <Calendar size={20} /> Agenda
    </button>
    <button className={`sp-navbtn ${vista === "clientes" ? "active" : ""}`} onClick={() => setVista("clientes")}>
      <Users size={20} /> Clientes
    </button>
    <button className={`sp-navbtn ${vista === "reportes" ? "active" : ""}`} onClick={() => setVista("reportes")}>
      <FileSpreadsheet size={20} /> Reportes
    </button>
  </nav>
);

/* ---------- vista agenda ---------- */
const VistaAgenda = ({ grupos, clientesPorId, filtro, setFiltro, onEditar, onBorrar, onToggle }) => {
  const dias = Object.keys(grupos).sort();
  return (
    <div style={{ paddingTop: 16 }}>
      <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto" }} className="sp-scroll">
        {[["proximos", "Próximos"], ["pendientes", "Pendientes"], ["todos", "Todos"]].map(([k, label]) => (
          <button key={k} onClick={() => setFiltro(k)}
            className="sp-btn"
            style={{
              padding: "8px 16px", fontSize: 13, whiteSpace: "nowrap",
              background: filtro === k ? "var(--teal)" : "#fff",
              color: filtro === k ? "#fff" : "var(--charcoal)",
              border: filtro === k ? "none" : "1.5px solid #E2DCD3"
            }}>
            {label}
          </button>
        ))}
      </div>

      {dias.length === 0 && (
        <div className="sp-card" style={{ padding: "40px 20px", textAlign: "center", color: "#A6B5AF" }}>
          <Calendar size={32} style={{ marginBottom: 10 }} />
          <div style={{ fontWeight: 600, color: "var(--charcoal)" }}>Sin trabajos agendados</div>
          <div style={{ fontSize: 13, marginTop: 4 }}>Toca el botón + para agendar el primer trabajo</div>
        </div>
      )}

      {dias.map(fecha => (
        <div key={fecha} style={{ marginBottom: 22 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "var(--teal)", textTransform: "capitalize", marginBottom: 8, paddingLeft: 4 }}>
            {fecha === todayISO() ? "Hoy · " : ""}{fmtFecha(fecha)}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {grupos[fecha].map(t => (
              <TrabajoCard key={t.id} t={t} cliente={clientesPorId[t.clienteId]}
                onEditar={() => onEditar(t)} onBorrar={() => onBorrar(t.id)} onToggle={() => onToggle(t.id)} />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
};

const TrabajoCard = ({ t, cliente, onEditar, onBorrar, onToggle }) => {
  const completado = t.estado === "Completado";
  return (
    <div className="sp-card" style={{ padding: 14, borderLeft: `4px solid ${completado ? "var(--mint)" : "var(--coral)"}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15 }}>{cliente?.nombre || "Cliente eliminado"}</div>
          <div style={{ fontSize: 13, color: "#7A8B85", display: "flex", alignItems: "center", gap: 4, marginTop: 2 }}>
            <Clock size={13} /> {fmtHora12(t.hora)} – {addHoras(t.hora, t.duracion)} · {t.duracion}h
          </div>
        </div>
        <button onClick={onToggle} className={`sp-tag ${completado ? "sp-tag-completado" : "sp-tag-pendiente"}`} style={{ border: "none", cursor: "pointer" }}>
          {completado ? <CheckCircle2 size={13} /> : <CircleDot size={13} />}
          {completado ? "Completado" : "Pendiente"}
        </button>
      </div>

      {cliente?.direccion && (
        <a href={mapsUrl(cliente.direccion)} target="_blank" rel="noreferrer"
          style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, fontSize: 13, color: "var(--teal)", textDecoration: "none" }}>
          <MapPin size={14} /> <span style={{ flex: 1 }}>{cliente.direccion}</span>
        </a>
      )}

      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10, alignItems: "center" }}>
        {cliente?.telefono && (
          <a href={telUrl(cliente.telefono)} className="sp-btn sp-btn-ghost" style={{ padding: "7px 12px", fontSize: 13, display: "flex", alignItems: "center", gap: 6, textDecoration: "none" }}>
            <Phone size={14} /> Llamar
          </a>
        )}
        <div className="sp-tag" style={{ background: "#F4EFE6", color: "var(--charcoal)" }}>
          <DollarSign size={13} /> {Number(t.valor || 0).toFixed(2)} · {t.pago}
        </div>
        {t.agendadoPor && (
          <div className="sp-tag" style={{ background: "#EFE8FB", color: "#5B4B9E" }}>
            Agendó: {t.agendadoPor}
          </div>
        )}
      </div>

      {t.notas && <div style={{ marginTop: 8, fontSize: 13, color: "#7A8B85", fontStyle: "italic" }}>{t.notas}</div>}

      <div style={{ display: "flex", gap: 8, marginTop: 12, borderTop: "1px solid #F1EBE2", paddingTop: 10 }}>
        <button onClick={onEditar} className="sp-btn" style={{ background: "none", color: "var(--teal)", fontSize: 13, padding: "4px 8px", display: "flex", alignItems: "center", gap: 4 }}>
          <Edit2 size={14} /> Editar
        </button>
        <button onClick={onBorrar} className="sp-btn" style={{ background: "none", color: "var(--coral)", fontSize: 13, padding: "4px 8px", display: "flex", alignItems: "center", gap: 4 }}>
          <Trash2 size={14} /> Eliminar
        </button>
      </div>
    </div>
  );
};

/* ---------- vista clientes ---------- */
const VistaClientes = ({ clientes, busqueda, setBusqueda, onNuevo, onEditar, onBorrar, onImportar }) => (
  <div style={{ paddingTop: 16 }}>
    <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
      <input className="sp-input" placeholder="Buscar cliente o teléfono..." value={busqueda} onChange={e => setBusqueda(e.target.value)} />
      <button onClick={onNuevo} className="sp-btn sp-btn-primary" style={{ padding: "0 16px", display: "flex", alignItems: "center", gap: 6 }}>
        <Plus size={16} /> Nuevo
      </button>
    </div>

    <label className="sp-btn sp-btn-ghost" style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "10px", marginBottom: 16, fontSize: 13 }}>
      <Upload size={16} /> Importar desde hoja de cálculo (Excel / CSV)
      <input type="file" accept=".xlsx,.xls,.csv" onChange={onImportar} style={{ display: "none" }} />
    </label>

    {clientes.length === 0 && (
      <div className="sp-card" style={{ padding: "40px 20px", textAlign: "center", color: "#A6B5AF" }}>
        <Users size={32} style={{ marginBottom: 10 }} />
        <div style={{ fontWeight: 600, color: "var(--charcoal)" }}>Sin clientes todavía</div>
      </div>
    )}

    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {clientes.map(c => (
        <div key={c.id} className="sp-card" style={{ padding: 14 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div style={{ fontWeight: 700, fontSize: 15 }}>{c.nombre}</div>
            <div style={{ display: "flex", gap: 6 }}>
              <button onClick={() => onEditar(c)} className="sp-btn" style={{ background: "var(--mint-light)", color: "var(--teal)", padding: 6 }}><Edit2 size={14} /></button>
              <button onClick={() => onBorrar(c.id)} className="sp-btn" style={{ background: "#FFE9E5", color: "var(--coral)", padding: 6 }}><Trash2 size={14} /></button>
            </div>
          </div>
          {c.telefono && (
            <a href={telUrl(c.telefono)} style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6, fontSize: 13, color: "var(--teal)", textDecoration: "none" }}>
              <Phone size={13} /> {c.telefono}
            </a>
          )}
          {c.direccion && (
            <a href={mapsUrl(c.direccion)} target="_blank" rel="noreferrer" style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 4, fontSize: 13, color: "#7A8B85", textDecoration: "none" }}>
              <MapPin size={13} /> {c.direccion}
            </a>
          )}
          {c.notas && <div style={{ marginTop: 6, fontSize: 13, color: "#A6B5AF" }}>{c.notas}</div>}
        </div>
      ))}
    </div>
  </div>
);

/* ---------- vista reportes ---------- */
const VistaReportes = ({ trabajos, clientesPorId }) => {
  const [periodo, setPeriodo] = useState("semana");
  const [refFecha, setRefFecha] = useState(todayISO());

  const { start, end } = getRango(periodo, refFecha);
  const filtrados = trabajos.filter(t => t.fecha >= start && t.fecha <= end)
    .sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));

  const total = filtrados.reduce((s, t) => s + Number(t.valor || 0), 0);
  const completados = filtrados.filter(t => t.estado === "Completado");
  const pendientes = filtrados.filter(t => t.estado === "Pendiente");
  const porPago = PAGOS.reduce((acc, p) => {
    acc[p] = filtrados.filter(t => t.pago === p).reduce((s, t) => s + Number(t.valor || 0), 0);
    return acc;
  }, {});

  const descargar = () => {
    const data = filtrados.map(t => ({
      Fecha: t.fecha,
      Hora: fmtHora12(t.hora),
      "Hora fin": addHoras(t.hora, t.duracion),
      "Duración (h)": t.duracion,
      Cliente: clientesPorId[t.clienteId]?.nombre || "",
      Teléfono: clientesPorId[t.clienteId]?.telefono || "",
      Dirección: clientesPorId[t.clienteId]?.direccion || "",
      Valor: Number(t.valor || 0),
      "Método de pago": t.pago,
      Estado: t.estado,
      "Agendado por": t.agendadoPor,
      Notas: t.notas || ""
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 12 }, { wch: 9 }, { wch: 9 }, { wch: 8 }, { wch: 22 }, { wch: 14 }, { wch: 30 }, { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 24 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Trabajos");

    const resumen = XLSX.utils.aoa_to_sheet([
      ["ShinyPro Clean - Resumen"],
      ["Período", `${start} a ${end}`],
      [],
      ["Total facturado", total],
      ["Trabajos completados", completados.length],
      ["Trabajos pendientes", pendientes.length],
      [],
      ["Por método de pago"],
      ...PAGOS.map(p => [p, porPago[p]])
    ]);
    XLSX.utils.book_append_sheet(wb, resumen, "Resumen");

    XLSX.writeFile(wb, `ShinyProClean_${periodo}_${start}_a_${end}.xlsx`);
  };

  return (
    <div style={{ paddingTop: 16 }}>
      <div className="sp-card" style={{ padding: 16, marginBottom: 16 }}>
        <span className="sp-label">Período</span>
        <div style={{ display: "flex", gap: 8, marginBottom: 12, overflowX: "auto" }} className="sp-scroll">
          {[["semana", "Semanal"], ["mes", "Mensual"], ["semestre", "Semestral"], ["año", "Anual"]].map(([k, label]) => (
            <button key={k} onClick={() => setPeriodo(k)} className="sp-btn"
              style={{
                padding: "8px 14px", fontSize: 13, whiteSpace: "nowrap",
                background: periodo === k ? "var(--teal)" : "#fff",
                color: periodo === k ? "#fff" : "var(--charcoal)",
                border: periodo === k ? "none" : "1.5px solid #E2DCD3"
              }}>{label}</button>
          ))}
        </div>
        <span className="sp-label">Fecha de referencia</span>
        <input type="date" className="sp-input" value={refFecha} onChange={e => setRefFecha(e.target.value)} />
        <div style={{ fontSize: 13, color: "#7A8B85", marginTop: 8 }}>Rango: {start} a {end}</div>
      </div>

      <div className="sp-card" style={{ padding: 16, marginBottom: 16 }}>
        <div style={{ fontSize: 13, color: "#7A8B85", marginBottom: 4 }}>Total facturado en el período</div>
        <div className="sp-display" style={{ fontSize: 32, fontWeight: 600, color: "var(--teal)" }}>
          ${total.toFixed(2)}
        </div>
        <div style={{ display: "flex", gap: 16, marginTop: 12, fontSize: 13 }}>
          <div><span className="sp-tag sp-tag-completado">{completados.length} completados</span></div>
          <div><span className="sp-tag sp-tag-pendiente">{pendientes.length} pendientes</span></div>
        </div>
        <div style={{ marginTop: 12, display: "flex", flexDirection: "column", gap: 4, fontSize: 13 }}>
          {PAGOS.map(p => (
            <div key={p} style={{ display: "flex", justifyContent: "space-between" }}>
              <span style={{ color: "#7A8B85" }}>{p}</span>
              <span style={{ fontWeight: 600 }}>${porPago[p].toFixed(2)}</span>
            </div>
          ))}
        </div>
      </div>

      <button onClick={descargar} className="sp-btn sp-btn-primary" style={{ width: "100%", padding: 14, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 24 }}>
        <FileSpreadsheet size={18} /> Descargar Excel ({filtrados.length} registros)
      </button>
    </div>
  );
};

/* ---------- modal cliente ---------- */
const ModalCliente = ({ data, onGuardar, onCerrar }) => {
  const [form, setForm] = useState(data || { nombre: "", telefono: "", direccion: "", notas: "" });
  return (
    <div className="sp-modalbg" onClick={onCerrar}>
      <div className="sp-modal" onClick={e => e.stopPropagation()}>
        <ModalHeader titulo={data ? "Editar cliente" : "Nuevo cliente"} onCerrar={onCerrar} />
        <div style={{ display: "flex", flexDirection: "column", gap: 14, marginTop: 16 }}>
          <div><span className="sp-label">Nombre</span>
            <input className="sp-input" value={form.nombre} onChange={e => setForm({ ...form, nombre: e.target.value })} placeholder="Nombre del cliente" /></div>
          <div><span className="sp-label">Teléfono</span>
            <input className="sp-input" value={form.telefono} onChange={e => setForm({ ...form, telefono: e.target.value })} placeholder="(555) 123-4567" /></div>
          <div><span className="sp-label">Dirección</span>
            <input className="sp-input" value={form.direccion} onChange={e => setForm({ ...form, direccion: e.target.value })} placeholder="Calle, ciudad, estado" /></div>
          <div><span className="sp-label">Notas</span>
            <input className="sp-input" value={form.notas} onChange={e => setForm({ ...form, notas: e.target.value })} placeholder="Opcional" /></div>
          <button className="sp-btn sp-btn-primary" style={{ padding: 14 }}
            disabled={!form.nombre.trim()}
            onClick={() => onGuardar({ ...form, id: data?.id })}>
            Guardar cliente
          </button>
        </div>
      </div>
    </div>
  );
};

/* ---------- modal trabajo ---------- */
const ModalTrabajo = ({ data, clientes, agendadores, onAgregarAgendador, onGuardar, onCrearClienteRapido, onCerrar }) => {
  const [form, setForm] = useState(data || {
    clienteId: clientes[0]?.id || "", fecha: todayISO(), hora: "09:00", duracion: 4,
    valor: "", pago: "Cash", estado: "Pendiente", agendadoPor: agendadores[0] || "", notas: ""
  });
  const [creandoCliente, setCreandoCliente] = useState(clientes.length === 0);
  const [nuevoCliente, setNuevoCliente] = useState({ nombre: "", telefono: "", direccion: "" });
  const [nuevoAgendador, setNuevoAgendador] = useState("");

  const handleGuardar = async () => {
    let clienteId = form.clienteId;
    if (creandoCliente) {
      if (!nuevoCliente.nombre.trim()) return alert("Escribe el nombre del cliente nuevo");
      const creado = await onCrearClienteRapido(nuevoCliente);
      clienteId = creado.id;
    }
    if (!clienteId) return alert("Selecciona o crea un cliente");
    onGuardar({ ...form, clienteId, id: data?.id, duracion: Number(form.duracion) || 4, valor: form.valor === "" ? 0 : Number(form.valor) });
  };

  return (
    <div className="sp-modalbg" onClick={onCerrar}>
      <div className="sp-modal" onClick={e => e.stopPropagation()}>
        <ModalHeader titulo={data ? "Editar trabajo" : "Agendar trabajo"} onCerrar={onCerrar} />
        <div style={{ display: "flex", flexDirection: "column", gap: 14, marginTop: 16 }}>

          <div>
            <span className="sp-label">Cliente</span>
            {!creandoCliente ? (
              <>
                <select className="sp-input" value={form.clienteId} onChange={e => setForm({ ...form, clienteId: e.target.value })}>
                  {clientes.map(c => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                </select>
                <button onClick={() => setCreandoCliente(true)} className="sp-btn" style={{ background: "none", color: "var(--teal)", fontSize: 12, padding: "6px 0" }}>
                  + Crear cliente nuevo
                </button>
              </>
            ) : (
              <div className="sp-card" style={{ padding: 12, background: "var(--mint-light)" }}>
                <input className="sp-input" placeholder="Nombre" value={nuevoCliente.nombre} onChange={e => setNuevoCliente({ ...nuevoCliente, nombre: e.target.value })} style={{ marginBottom: 8 }} />
                <input className="sp-input" placeholder="Teléfono" value={nuevoCliente.telefono} onChange={e => setNuevoCliente({ ...nuevoCliente, telefono: e.target.value })} style={{ marginBottom: 8 }} />
                <input className="sp-input" placeholder="Dirección" value={nuevoCliente.direccion} onChange={e => setNuevoCliente({ ...nuevoCliente, direccion: e.target.value })} style={{ marginBottom: 8 }} />
                {clientes.length > 0 && (
                  <button onClick={() => setCreandoCliente(false)} className="sp-btn" style={{ background: "none", color: "var(--teal)", fontSize: 12, padding: 0 }}>
                    Usar cliente existente
                  </button>
                )}
              </div>
            )}
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><span className="sp-label">Fecha</span>
              <input type="date" className="sp-input" value={form.fecha} onChange={e => setForm({ ...form, fecha: e.target.value })} /></div>
            <div style={{ flex: 1 }}><span className="sp-label">Hora inicio</span>
              <input type="time" className="sp-input" value={form.hora} onChange={e => setForm({ ...form, hora: e.target.value })} /></div>
          </div>

          <div>
            <span className="sp-label">Duración (horas)</span>
            <input type="number" min="0.5" step="0.5" className="sp-input" value={form.duracion} onChange={e => setForm({ ...form, duracion: e.target.value })} />
            <div style={{ fontSize: 12, color: "#A6B5AF", marginTop: 4 }}>Termina aprox. {addHoras(form.hora, Number(form.duracion) || 0)}</div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1 }}><span className="sp-label">Valor cobrado ($)</span>
              <input type="number" min="0" step="0.01" className="sp-input" value={form.valor} onChange={e => setForm({ ...form, valor: e.target.value })} placeholder="0.00" /></div>
            <div style={{ flex: 1 }}><span className="sp-label">Método de pago</span>
              <select className="sp-input" value={form.pago} onChange={e => setForm({ ...form, pago: e.target.value })}>
                {PAGOS.map(p => <option key={p} value={p}>{p}</option>)}
              </select></div>
          </div>

          <div>
            <span className="sp-label">Estado</span>
            <div style={{ display: "flex", gap: 8 }}>
              {["Pendiente", "Completado"].map(s => (
                <button key={s} onClick={() => setForm({ ...form, estado: s })} className="sp-btn"
                  style={{
                    flex: 1, padding: 10, fontSize: 13,
                    background: form.estado === s ? (s === "Completado" ? "var(--teal)" : "var(--coral)") : "#fff",
                    color: form.estado === s ? "#fff" : "var(--charcoal)",
                    border: form.estado === s ? "none" : "1.5px solid #E2DCD3"
                  }}>{s}</button>
              ))}
            </div>
          </div>

          <div>
            <span className="sp-label">Agendado por</span>
            <select className="sp-input" value={form.agendadoPor} onChange={e => setForm({ ...form, agendadoPor: e.target.value })}>
              {agendadores.map(a => <option key={a} value={a}>{a}</option>)}
              <option value="__nuevo__">+ Agregar nombre nuevo</option>
            </select>
            {form.agendadoPor === "__nuevo__" ? (
              <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                <input className="sp-input" placeholder="Nombre" value={nuevoAgendador} onChange={e => setNuevoAgendador(e.target.value)} />
                <button className="sp-btn sp-btn-ghost" style={{ padding: "0 14px" }} onClick={() => {
                  if (!nuevoAgendador.trim()) return;
                  onAgregarAgendador(nuevoAgendador.trim());
                  setForm({ ...form, agendadoPor: nuevoAgendador.trim() });
                  setNuevoAgendador("");
                }}>Agregar</button>
              </div>
            ) : null}
          </div>

          <div><span className="sp-label">Notas</span>
            <input className="sp-input" value={form.notas} onChange={e => setForm({ ...form, notas: e.target.value })} placeholder="Opcional" /></div>

          <button className="sp-btn sp-btn-primary" style={{ padding: 14, marginTop: 6 }} onClick={handleGuardar}>
            {data ? "Guardar cambios" : "Agendar trabajo"}
          </button>
        </div>
      </div>
    </div>
  );
};

const ModalHeader = ({ titulo, onCerrar }) => (
  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
    <div className="sp-display" style={{ fontSize: 20, fontWeight: 600, color: "var(--teal)" }}>{titulo}</div>
    <button onClick={onCerrar} className="sp-btn" style={{ background: "#F4EFE6", padding: 6, borderRadius: "50%" }}><X size={18} /></button>
  </div>
);
