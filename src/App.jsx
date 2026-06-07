import React, { useState, useEffect, useMemo } from 'react';
import {
  Plus, Trash2, Download, TrendingUp, TrendingDown, ArrowRightLeft,
  Calendar, Wallet, Target, FileSpreadsheet, Edit2, X, Check,
  ChevronLeft, ChevronRight, Sun, Moon, PiggyBank, AlertTriangle, RefreshCw,
} from 'lucide-react';
import * as XLSX from 'xlsx';

// ============ CONFIGURACIÓN BASE ============
const CUENTAS_DEFAULT = [
  { id: 'caja_gal_ars', nombre: 'Caja GAL ARS', moneda: 'ARS', tipo: 'liquido' },
  { id: 'caja_gal_usd', nombre: 'Caja GAL USD', moneda: 'USD', tipo: 'liquido' },
  { id: 'fima_gal_ars', nombre: 'FIMA GAL ARS', moneda: 'ARS', tipo: 'invertido' },
  { id: 'fima_gal_usd', nombre: 'FIMA GAL USD', moneda: 'USD', tipo: 'invertido' },
  { id: 'mp_ars', nombre: 'MP ARS', moneda: 'ARS', tipo: 'liquido' },
  { id: 'mp_usd', nombre: 'MP USD', moneda: 'USD', tipo: 'liquido' },
  { id: 'balanz', nombre: 'Balanz', moneda: 'USD', tipo: 'invertido' },
  { id: 'binance', nombre: 'Binance', moneda: 'USD', tipo: 'invertido' },
  { id: 'cash_ars', nombre: 'Cash ARS', moneda: 'ARS', tipo: 'liquido' },
  { id: 'cash_usd', nombre: 'Cash USD', moneda: 'USD', tipo: 'liquido' },
];

const MEDIOS_PAGO = ['Visa', 'Master', 'MercadoPago', 'Efectivo', 'Débito'];

const CATEGORIAS_EGRESO = {
  'Crédito Hipotecario': ['Cuota mensual'],
  'Tarjeta (cuotas)': ['Cuota tarjeta'],
  'Servicios e Impuestos': ['Ecogas', 'Edemsa', 'Aysam', 'Inmobiliario', 'Municipalidad', 'Internet', 'Expensas', 'Seguro', 'ABL/AFIP', 'Otros'],
  Supermercado: ['Compra grande', 'Compra chica', 'Verdulería', 'Carnicería'],
  'Comida afuera': ['Restaurante', 'Delivery', 'Almuerzo trabajo', 'Café'],
  'Ocio / Eventos': ['Bar', 'Boliche', 'Cine/Teatro', 'Eventos'],
  Auto: ['Combustible', 'Mantenimiento/Service', 'Seguro', 'Patente', 'Lavado', 'Estacionamiento'],
  Fitness: ['Gimnasio', 'Fútbol', 'Suplementos', 'Equipamiento'],
  Salud: ['Médico/Consultas', 'Farmacia', 'Estética', 'Obra social'],
  'Cuidado personal': ['Peluquería', 'Productos cuidado'],
  'Hogar / Depa': ['Limpieza', 'Mejoras', 'Decoración', 'Electrodomésticos'],
  Plataformas: ['Streaming', 'Música', 'Cloud', 'IA', 'Otros'],
  Educación: ['Cursos', 'Libros', 'Coaching'],
  Ropa: ['Urbano', 'Deportivo', 'Formal', 'Calzado'],
  Viajes: ['Pasajes', 'Alojamiento', 'Comida viaje', 'Actividades'],
  Regalos: ['Familia', 'Pareja', 'Amigos', 'Eventos sociales'],
  Otros: ['Otros'],
};

const CATEGORIAS_INGRESO = {
  Sueldo: ['Sueldo mensual', 'Aguinaldo', 'Bono/Premio'],
  Extras: ['Freelance', 'Ventas', 'Comisiones'],
  Inversiones: ['Renta FCI', 'Dividendos', 'Venta con ganancia', 'Intereses'],
  Otros: ['Regalos', 'Reintegros', 'Vouchers'],
};

const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

// ============ HELPERS ============
const fmtMoney = (n, moneda = 'ARS') => {
  if (n === null || n === undefined || isNaN(n)) return '-';
  const symbol = moneda === 'USD' ? 'US$' : '$';
  return `${symbol} ${Math.round(n).toLocaleString('es-AR')}`;
};

const fmtMoneyShort = (n) => {
  if (n === null || n === undefined || isNaN(n)) return '$ 0';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1000000) return `${sign}$ ${(abs / 1000000).toFixed(1)}M`;
  if (abs >= 1000) return `${sign}$ ${Math.round(abs / 1000)}k`;
  return `${sign}$ ${Math.round(abs)}`;
};

const fmtPct = (n) => {
  if (n === null || n === undefined || isNaN(n) || !isFinite(n)) return '-';
  return `${(n * 100).toFixed(0)}%`;
};

const hoy = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

const sumarMeses = (fechaStr, meses) => {
  const [y, m, d] = fechaStr.split('-').map(Number);
  const date = new Date(y, m - 1 + meses, d);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
};

const mesAnio = (fechaStr) => fechaStr.slice(0, 7);
const anio = (fechaStr) => fechaStr.slice(0, 4);

// ============ STORAGE ============
const STORAGE_KEYS = {
  movimientos: 'finanzas_movimientos_v2',
  presupuesto: 'finanzas_presupuesto_v2',
  config: 'finanzas_config_v2',
  patrimonio: 'finanzas_patrimonio_v2',
  snapshots: 'finanzas_snapshots_v2',
};

const loadData = (key, defaultValue) => {
  try {
    const raw = localStorage.getItem(key);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.error('Error cargando', e); }
  return defaultValue;
};

const saveData = (key, value) => {
  try { localStorage.setItem(key, JSON.stringify(value)); }
  catch (e) { console.error('Error guardando', e); }
};

function montoEnARS(m, config) {
  if (m.moneda === 'USD') return m.monto * (m.cotizacion || config.cotizacionDolar);
  return m.monto;
}

// ============ TEMA (claro/oscuro) ============
const TEMAS = {
  claro: {
    bg: 'bg-slate-50', card: 'bg-white', cardBorder: 'border-slate-200',
    text: 'text-slate-900', textMuted: 'text-slate-500', textSoft: 'text-slate-400',
    input: 'bg-white border-slate-300 text-slate-900', headerBg: 'bg-white',
    navBg: 'bg-white', surface: 'bg-slate-50', surfaceText: 'text-slate-600',
    chipBg: 'bg-slate-100', divide: 'border-slate-200',
  },
  oscuro: {
    bg: 'bg-slate-950', card: 'bg-slate-900', cardBorder: 'border-slate-800',
    text: 'text-slate-100', textMuted: 'text-slate-400', textSoft: 'text-slate-500',
    input: 'bg-slate-800 border-slate-700 text-slate-100', headerBg: 'bg-slate-900',
    navBg: 'bg-slate-900', surface: 'bg-slate-800', surfaceText: 'text-slate-300',
    chipBg: 'bg-slate-800', divide: 'border-slate-800',
  },
};

// ============ COMPONENTE PRINCIPAL ============
export default function FinanzasApp() {
  const [tab, setTab] = useState('dashboard');
  const [movimientos, setMovimientos] = useState([]);
  const [presupuesto, setPresupuesto] = useState({});
  const [patrimonio, setPatrimonio] = useState({});
  const [snapshots, setSnapshots] = useState([]);
  const [config, setConfig] = useState({ cotizacionDolar: 1400, mesActivo: mesAnio(hoy()), tema: 'claro', umbralConciliacion: 50000 });
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    setMovimientos(loadData(STORAGE_KEYS.movimientos, []));
    setPresupuesto(loadData(STORAGE_KEYS.presupuesto, {}));
    setPatrimonio(loadData(STORAGE_KEYS.patrimonio, {}));
    setSnapshots(loadData(STORAGE_KEYS.snapshots, []));
    setConfig(loadData(STORAGE_KEYS.config, { cotizacionDolar: 1400, mesActivo: mesAnio(hoy()), tema: 'claro', umbralConciliacion: 50000 }));
    setLoading(false);
  }, []);

  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.movimientos, movimientos); }, [movimientos, loading]);
  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.presupuesto, presupuesto); }, [presupuesto, loading]);
  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.patrimonio, patrimonio); }, [patrimonio, loading]);
  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.snapshots, snapshots); }, [snapshots, loading]);
  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.config, config); }, [config, loading]);

  const t = TEMAS[config.tema || 'claro'];

  const agregarMovimiento = (mov) => {
    const id = Date.now() + Math.random();
    if (mov.tipo === 'egreso' && mov.cuotas && mov.cuotas > 1) {
      const movs = [];
      const montoCuota = mov.monto / mov.cuotas;
      for (let i = 0; i < mov.cuotas; i++) {
        movs.push({
          ...mov, id: id + i, monto: montoCuota, fecha: sumarMeses(mov.fecha, i),
          descripcion: `${mov.descripcion || mov.categoria} (${i + 1}/${mov.cuotas})`,
          esCuota: true, grupoCuota: id, numeroCuota: i + 1, totalCuotas: mov.cuotas,
        });
      }
      setMovimientos((prev) => [...prev, ...movs]);
    } else {
      setMovimientos((prev) => [...prev, { ...mov, id }]);
    }
  };

  const eliminarMovimiento = (id) => {
    const mov = movimientos.find((m) => m.id === id);
    if (mov?.grupoCuota) {
      if (confirm('Es una cuota de un grupo. ¿Eliminar TODAS las cuotas de esta compra?')) {
        setMovimientos((prev) => prev.filter((m) => m.grupoCuota !== mov.grupoCuota));
        return;
      }
    }
    setMovimientos((prev) => prev.filter((m) => m.id !== id));
  };

  const actualizarMovimiento = (id, cambios) => {
    setMovimientos((prev) => prev.map((m) => (m.id === id ? { ...m, ...cambios } : m)));
  };

  const actualizarPresupuesto = (mes, categoria, monto) => {
    setPresupuesto((prev) => ({ ...prev, [mes]: { ...(prev[mes] || {}), [categoria]: monto } }));
  };

  const actualizarPatrimonio = (cuentaId, saldo) => {
    setPatrimonio((prev) => ({ ...prev, [cuentaId]: saldo }));
  };

  const guardarSnapshot = (totalARS, totalUSD) => {
    const mesActual = mesAnio(hoy());
    setSnapshots((prev) => {
      const sinEsteMes = prev.filter((s) => s.mes !== mesActual);
      return [...sinEsteMes, { mes: mesActual, fecha: hoy(), totalARS, totalUSD, detalle: { ...patrimonio } }].sort((a, b) => a.mes.localeCompare(b.mes));
    });
  };

  if (loading) {
    return <div className="min-h-screen bg-slate-50 flex items-center justify-center"><div className="text-slate-600">Cargando...</div></div>;
  }

  const tabs = [
    { id: 'dashboard', label: 'Inicio', icon: TrendingUp },
    { id: 'movimientos', label: 'Movim.', icon: Wallet },
    { id: 'nuevo', label: 'Nuevo', icon: Plus },
    { id: 'patrimonio', label: 'Patrim.', icon: PiggyBank },
    { id: 'mas', label: 'Más', icon: Target },
  ];

  return (
    <div className={`min-h-screen ${t.bg} pb-20 transition-colors`}>
      <div className={`${t.headerBg} border-b ${t.cardBorder} px-4 py-3 sticky top-0 z-10`}>
        <div className="flex items-center justify-between">
          <h1 className={`text-lg font-bold ${t.text}`}>Finanzas Agustín</h1>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setConfig({ ...config, tema: config.tema === 'oscuro' ? 'claro' : 'oscuro' })}
              className={`p-1.5 rounded-lg ${t.surface} ${t.surfaceText}`}
            >
              {config.tema === 'oscuro' ? <Sun size={18} /> : <Moon size={18} />}
            </button>
            <div className="flex items-center gap-1.5">
              <span className={`text-xs ${t.textMuted}`}>USD</span>
              <input
                type="number" inputMode="numeric" value={config.cotizacionDolar}
                onChange={(e) => setConfig({ ...config, cotizacionDolar: Number(e.target.value) })}
                className={`w-20 px-2 py-1 text-right border rounded text-sm ${t.input}`}
              />
            </div>
          </div>
        </div>
      </div>

      <div className="px-3 py-3 max-w-2xl mx-auto">
        {tab === 'dashboard' && <Dashboard movimientos={movimientos} presupuesto={presupuesto} config={config} setConfig={setConfig} t={t} />}
        {tab === 'movimientos' && <Movimientos movimientos={movimientos} eliminar={eliminarMovimiento} actualizar={actualizarMovimiento} config={config} t={t} />}
        {tab === 'nuevo' && <NuevoMovimiento agregar={agregarMovimiento} config={config} t={t} />}
        {tab === 'patrimonio' && <Patrimonio patrimonio={patrimonio} actualizar={actualizarPatrimonio} config={config} setConfig={setConfig} movimientos={movimientos} snapshots={snapshots} guardarSnapshot={guardarSnapshot} agregar={agregarMovimiento} t={t} />}
        {tab === 'mas' && <Mas presupuesto={presupuesto} actualizar={actualizarPresupuesto} movimientos={movimientos} config={config} patrimonio={patrimonio} snapshots={snapshots} t={t} />}
      </div>

      <div className={`fixed bottom-0 left-0 right-0 ${t.navBg} border-t ${t.cardBorder} z-20`}>
        <div className="flex max-w-2xl mx-auto">
          {tabs.map((tb) => {
            const activo = tab === tb.id;
            const esNuevo = tb.id === 'nuevo';
            return (
              <button key={tb.id} onClick={() => setTab(tb.id)} className={`flex-1 flex flex-col items-center justify-center gap-0.5 py-2 ${activo ? 'text-blue-500' : t.textMuted}`}>
                {esNuevo ? (
                  <div className={`-mt-5 w-12 h-12 rounded-full flex items-center justify-center shadow-lg ${activo ? 'bg-blue-700' : 'bg-blue-600'}`}>
                    <tb.icon size={24} className="text-white" />
                  </div>
                ) : (<tb.icon size={22} />)}
                <span className="text-[10px] font-medium">{tb.label}</span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ============ DASHBOARD (con selector de período) ============
function Dashboard({ movimientos, presupuesto, config, setConfig, t }) {
  const [periodo, setPeriodo] = useState('mes'); // mes | anio | historico
  const mesActivo = config.mesActivo || mesAnio(hoy());

  const cambiarMes = (delta) => {
    const [y, m] = mesActivo.split('-').map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    setConfig({ ...config, mesActivo: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}` });
  };

  const movsFiltrados = useMemo(() => {
    if (periodo === 'mes') return movimientos.filter((m) => mesAnio(m.fecha) === mesActivo);
    if (periodo === 'anio') return movimientos.filter((m) => anio(m.fecha) === anio(mesActivo + '-01'));
    return movimientos;
  }, [movimientos, mesActivo, periodo]);

  const totalIngresos = useMemo(() => movsFiltrados.filter((m) => m.tipo === 'ingreso').reduce((s, m) => s + montoEnARS(m, config), 0), [movsFiltrados, config]);
  const totalEgresos = useMemo(() => movsFiltrados.filter((m) => m.tipo === 'egreso').reduce((s, m) => s + montoEnARS(m, config), 0), [movsFiltrados, config]);

  const egresosPorCategoria = useMemo(() => {
    const acc = {};
    movsFiltrados.filter((m) => m.tipo === 'egreso').forEach((m) => { acc[m.categoria] = (acc[m.categoria] || 0) + montoEnARS(m, config); });
    return acc;
  }, [movsFiltrados, config]);

  // Presupuesto: para mes usa el mes; para año suma los 12; histórico sin plan
  const presupuestoComparable = useMemo(() => {
    if (periodo === 'mes') return presupuesto[mesActivo] || {};
    if (periodo === 'anio') {
      const acc = {};
      const y = anio(mesActivo + '-01');
      Object.keys(presupuesto).filter((k) => k.startsWith(y)).forEach((k) => {
        Object.entries(presupuesto[k]).forEach(([cat, v]) => { acc[cat] = (acc[cat] || 0) + (v || 0); });
      });
      return acc;
    }
    return {};
  }, [presupuesto, mesActivo, periodo]);

  const totalPlanificado = Object.values(presupuestoComparable).reduce((s, v) => s + (v || 0), 0);
  const balance = totalIngresos - totalEgresos;
  const cumplimiento = totalPlanificado > 0 ? totalEgresos / totalPlanificado : null;
  const [y, mNum] = mesActivo.split('-').map(Number);

  const categoriasConDatos = Object.keys(CATEGORIAS_EGRESO).filter((cat) => (presupuestoComparable[cat] || 0) > 0 || (egresosPorCategoria[cat] || 0) > 0);

  const tituloPeriodo = periodo === 'mes' ? `${MESES[mNum - 1]} ${y}` : periodo === 'anio' ? `Año ${y}` : 'Todo el historial';

  return (
    <div className="space-y-3">
      {/* Selector de período */}
      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-1 grid grid-cols-3 gap-1`}>
        {[{ id: 'mes', l: 'Mes' }, { id: 'anio', l: 'Año' }, { id: 'historico', l: 'Histórico' }].map((p) => (
          <button key={p.id} onClick={() => setPeriodo(p.id)} className={`py-2 rounded-lg text-sm font-medium ${periodo === p.id ? 'bg-blue-600 text-white' : `${t.textMuted}`}`}>{p.l}</button>
        ))}
      </div>

      {/* Navegación mes/año (solo si no es histórico) */}
      {periodo !== 'historico' && (
        <div className={`${t.card} rounded-xl border ${t.cardBorder} px-3 py-2.5 flex items-center justify-between`}>
          <button onClick={() => cambiarMes(periodo === 'anio' ? -12 : -1)} className={`p-2 rounded-lg ${t.textMuted}`}><ChevronLeft size={20} /></button>
          <h2 className={`text-base font-bold ${t.text}`}>{tituloPeriodo}</h2>
          <button onClick={() => cambiarMes(periodo === 'anio' ? 12 : 1)} className={`p-2 rounded-lg ${t.textMuted}`}><ChevronRight size={20} /></button>
        </div>
      )}
      {periodo === 'historico' && (
        <div className={`${t.card} rounded-xl border ${t.cardBorder} px-3 py-2.5 text-center`}>
          <h2 className={`text-base font-bold ${t.text}`}>Todo el historial</h2>
        </div>
      )}

      <div className="grid grid-cols-2 gap-2">
        <KPI label="Ingresos" valor={totalIngresos} color="green" t={t} />
        <KPI label="Egresos" valor={totalEgresos} color="red" t={t} />
        <KPI label="Balance" valor={balance} color={balance >= 0 ? 'green' : 'red'} t={t} />
        <KPI label="Cumplim. plan" valor={cumplimiento} esPct color={cumplimiento === null ? 'slate' : cumplimiento > 1 ? 'red' : 'green'} t={t} />
      </div>

      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
        <h3 className={`text-sm font-semibold ${t.text} mb-2`}>Seguimiento por categoría</h3>
        {categoriasConDatos.length === 0 && <p className={`text-sm ${t.textSoft} text-center py-4`}>Sin movimientos en este período</p>}
        <div className="space-y-2.5">
          {categoriasConDatos.map((cat) => {
            const plan = presupuestoComparable[cat] || 0;
            const real = egresosPorCategoria[cat] || 0;
            const pct = plan > 0 ? real / plan : null;
            const pctBarra = plan > 0 ? Math.min((real / plan) * 100, 100) : (real > 0 ? 100 : 0);
            const colorBarra = pct === null ? 'bg-slate-400' : pct > 1 ? 'bg-red-500' : pct > 0.9 ? 'bg-amber-500' : 'bg-green-500';
            return (
              <div key={cat}>
                <div className="flex items-center justify-between text-sm mb-1">
                  <span className={`font-medium ${t.surfaceText}`}>{cat}</span>
                  <span className={t.textMuted}>{fmtMoneyShort(real)}{plan > 0 && <span className={t.textSoft}> / {fmtMoneyShort(plan)}</span>}</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className={`flex-1 ${t.chipBg} rounded-full h-2 overflow-hidden`}>
                    <div className={`h-full ${colorBarra}`} style={{ width: `${pctBarra}%` }} />
                  </div>
                  <span className={`text-xs font-medium w-12 text-right ${pct === null ? t.textSoft : pct > 1 ? 'text-red-500' : pct > 0.9 ? 'text-amber-500' : 'text-green-500'}`}>{pct === null ? 's/plan' : fmtPct(pct)}</span>
                </div>
              </div>
            );
          })}
        </div>
        {categoriasConDatos.length > 0 && (
          <div className={`flex items-center justify-between mt-3 pt-3 border-t ${t.divide} text-sm font-bold`}>
            <span className={t.text}>Total</span>
            <span className={t.text}>{fmtMoney(totalEgresos)}{totalPlanificado > 0 && <span className={`${t.textSoft} font-medium`}> / {fmtMoney(totalPlanificado)}</span>}</span>
          </div>
        )}
      </div>

      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
        <h3 className={`text-sm font-semibold ${t.text} mb-2`}>Egresos por medio de pago</h3>
        <div className="grid grid-cols-2 gap-2">
          {MEDIOS_PAGO.map((mp) => {
            const total = movsFiltrados.filter((m) => m.tipo === 'egreso' && m.medioPago === mp).reduce((s, m) => s + montoEnARS(m, config), 0);
            const pct = totalEgresos > 0 ? total / totalEgresos : 0;
            if (total === 0) return null;
            return (
              <div key={mp} className={`${t.surface} rounded-lg p-2.5`}>
                <div className={`text-xs ${t.textMuted}`}>{mp}</div>
                <div className={`text-sm font-semibold ${t.text}`}>{fmtMoney(total)}</div>
                <div className={`text-xs ${t.textSoft}`}>{fmtPct(pct)}</div>
              </div>
            );
          })}
          {totalEgresos === 0 && <p className={`text-sm ${t.textSoft} col-span-2 text-center py-2`}>Sin egresos</p>}
        </div>
      </div>
    </div>
  );
}

function KPI({ label, valor, color = 'slate', esPct = false, t }) {
  const colorMap = {
    green: 'text-green-600 bg-green-500/10',
    red: 'text-red-600 bg-red-500/10',
    slate: `${t.textMuted} ${t.surface}`,
  };
  const display = esPct ? (valor === null ? 'sin plan' : fmtPct(valor)) : fmtMoney(valor);
  return (
    <div className={`rounded-xl p-3 ${colorMap[color]}`}>
      <div className="text-xs opacity-80">{label}</div>
      <div className="text-lg font-bold mt-0.5 leading-tight">{display}</div>
    </div>
  );
}

// ============ PATRIMONIO NETO ============
function Patrimonio({ patrimonio, actualizar, config, setConfig, movimientos, snapshots, guardarSnapshot, agregar, t }) {
  const [msg, setMsg] = useState('');

  const totalARS = useMemo(() => {
    return CUENTAS_DEFAULT.reduce((s, c) => {
      const saldo = patrimonio[c.id] || 0;
      const enARS = c.moneda === 'USD' ? saldo * config.cotizacionDolar : saldo;
      return s + enARS;
    }, 0);
  }, [patrimonio, config.cotizacionDolar]);

  const totalUSD = totalARS / config.cotizacionDolar;

  const liquido = useMemo(() => CUENTAS_DEFAULT.filter((c) => c.tipo === 'liquido').reduce((s, c) => {
    const saldo = patrimonio[c.id] || 0;
    return s + (c.moneda === 'USD' ? saldo * config.cotizacionDolar : saldo);
  }, 0), [patrimonio, config.cotizacionDolar]);

  const invertido = totalARS - liquido;

  // ===== CONCILIACIÓN =====
  // Variación esperada por flujos (todos los movimientos): ingresos - egresos
  const flujoNeto = useMemo(() => {
    return movimientos.reduce((s, m) => {
      if (m.tipo === 'ingreso') return s + montoEnARS(m, config);
      if (m.tipo === 'egreso') return s - montoEnARS(m, config);
      return s;
    }, 0);
  }, [movimientos, config]);

  // Snapshot anterior (último guardado)
  const snapshotPrevio = snapshots.length > 0 ? snapshots[snapshots.length - 1] : null;
  const patrimonioPrevio = snapshotPrevio ? snapshotPrevio.totalARS : null;

  // Conciliación: solo si hay un snapshot previo para comparar
  const conciliacion = useMemo(() => {
    if (patrimonioPrevio === null) return null;
    const esperado = patrimonioPrevio + flujoNeto;
    const real = totalARS;
    const diferencia = real - esperado;
    return { esperado, real, diferencia, abs: Math.abs(diferencia) };
  }, [patrimonioPrevio, flujoNeto, totalARS]);

  const umbral = config.umbralConciliacion || 50000;

  const handleGuardarSnapshot = () => {
    guardarSnapshot(totalARS, totalUSD);
    setMsg('✓ Foto del mes guardada');
    setTimeout(() => setMsg(''), 2000);
  };

  const ajustarDiferencia = () => {
    if (!conciliacion) return;
    const dif = conciliacion.diferencia;
    agregar({
      tipo: dif > 0 ? 'ingreso' : 'egreso',
      fecha: hoy(),
      monto: Math.abs(dif),
      moneda: 'ARS',
      cotizacion: null,
      categoria: dif > 0 ? 'Otros' : 'Otros',
      subcategoria: 'Ajuste conciliación',
      descripcion: 'Ajuste automático de conciliación',
    });
    setMsg('✓ Ajuste registrado');
    setTimeout(() => setMsg(''), 2000);
  };

  return (
    <div className="space-y-3">
      <h3 className={`text-base font-semibold ${t.text} px-1`}>Patrimonio neto</h3>

      {/* Total destacado */}
      <div className="bg-blue-600 rounded-xl p-4 text-center text-white">
        <div className="text-xs opacity-80">Patrimonio total</div>
        <div className="text-3xl font-bold mt-1">{fmtMoney(totalARS)}</div>
        <div className="text-sm opacity-90 mt-0.5">≈ {fmtMoney(totalUSD, 'USD')}</div>
      </div>

      {/* Líquido vs Invertido */}
      <div className="grid grid-cols-2 gap-2">
        <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
          <div className={`text-xs ${t.textMuted}`}>Líquido</div>
          <div className={`text-base font-bold ${t.text}`}>{fmtMoney(liquido)}</div>
          <div className={`text-xs ${t.textSoft}`}>{totalARS > 0 ? fmtPct(liquido / totalARS) : '-'}</div>
        </div>
        <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
          <div className={`text-xs ${t.textMuted}`}>Invertido</div>
          <div className={`text-base font-bold ${t.text}`}>{fmtMoney(invertido)}</div>
          <div className={`text-xs ${t.textSoft}`}>{totalARS > 0 ? fmtPct(invertido / totalARS) : '-'}</div>
        </div>
      </div>

      {/* Saldos por cuenta */}
      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
        <h4 className={`text-sm font-semibold ${t.text} mb-2`}>Saldos por cuenta</h4>
        <p className={`text-xs ${t.textSoft} mb-3`}>Actualizá cuánto tenés hoy en cada cuenta</p>
        <div className="space-y-2">
          {CUENTAS_DEFAULT.map((c) => (
            <div key={c.id} className="flex items-center justify-between gap-2">
              <div className="flex-1 min-w-0">
                <span className={`text-sm ${t.surfaceText}`}>{c.nombre}</span>
                <span className={`text-xs ml-1.5 px-1.5 py-0.5 rounded-full ${c.tipo === 'invertido' ? 'bg-purple-500/15 text-purple-500' : 'bg-blue-500/15 text-blue-500'}`}>{c.moneda}</span>
              </div>
              <input
                type="number" inputMode="decimal" value={patrimonio[c.id] || ''}
                onChange={(e) => actualizar(c.id, Number(e.target.value) || 0)}
                placeholder="0"
                className={`w-28 px-2 py-1.5 text-right border rounded-lg text-sm ${t.input}`}
              />
            </div>
          ))}
        </div>
      </div>

      {/* Conciliación */}
      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
        <div className="flex items-center justify-between mb-2">
          <h4 className={`text-sm font-semibold ${t.text}`}>Control / Conciliación</h4>
          <div className="flex items-center gap-1">
            <span className={`text-xs ${t.textSoft}`}>Umbral:</span>
            <input type="number" inputMode="numeric" value={config.umbralConciliacion || 50000}
              onChange={(e) => setConfig({ ...config, umbralConciliacion: Number(e.target.value) || 0 })}
              className={`w-20 px-1.5 py-1 text-right border rounded text-xs ${t.input}`} />
          </div>
        </div>

        {conciliacion === null ? (
          <p className={`text-xs ${t.textSoft}`}>
            Guardá una primera foto del patrimonio para empezar a conciliar. La próxima vez que vuelvas, la app comparará tus saldos reales contra lo esperado según tus ingresos y egresos.
          </p>
        ) : (
          <div className="space-y-2">
            <div className="flex justify-between text-sm">
              <span className={t.textMuted}>Esperado (foto previa + flujos)</span>
              <span className={t.text}>{fmtMoney(conciliacion.esperado)}</span>
            </div>
            <div className="flex justify-between text-sm">
              <span className={t.textMuted}>Real (saldos cargados)</span>
              <span className={t.text}>{fmtMoney(conciliacion.real)}</span>
            </div>
            <div className={`flex justify-between text-sm font-semibold pt-2 border-t ${t.divide}`}>
              <span className={t.text}>Diferencia</span>
              <span className={conciliacion.abs <= umbral ? 'text-green-500' : 'text-red-500'}>{fmtMoney(conciliacion.diferencia)}</span>
            </div>

            {conciliacion.abs <= umbral ? (
              <div className="bg-green-500/10 rounded-lg p-2.5 flex items-center gap-2">
                <Check size={16} className="text-green-500 shrink-0" />
                <span className="text-xs text-green-600 flex-1">Diferencia menor al umbral. Podés ajustar automáticamente.</span>
                {conciliacion.abs > 0 && <button onClick={ajustarDiferencia} className="text-xs bg-green-600 text-white px-2 py-1 rounded font-medium">Ajustar</button>}
              </div>
            ) : (
              <div className="bg-red-500/10 rounded-lg p-2.5 flex items-start gap-2">
                <AlertTriangle size={16} className="text-red-500 shrink-0 mt-0.5" />
                <span className="text-xs text-red-600">Diferencia grande ({fmtMoney(conciliacion.abs)}). Probablemente falta cargar un movimiento importante. Revisá antes de ajustar.</span>
              </div>
            )}
          </div>
        )}
      </div>

      <button onClick={handleGuardarSnapshot} className="w-full bg-blue-600 active:bg-blue-700 text-white font-semibold py-3 rounded-xl flex items-center justify-center gap-2">
        <RefreshCw size={18} /> Guardar foto del mes
      </button>
      {msg && <div className="text-center text-sm font-medium text-green-600">{msg}</div>}

      {/* Evolución */}
      {snapshots.length > 0 && (
        <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
          <h4 className={`text-sm font-semibold ${t.text} mb-2`}>Evolución patrimonial</h4>
          <div className="space-y-1.5">
            {snapshots.map((s, i) => {
              const prev = i > 0 ? snapshots[i - 1].totalARS : null;
              const variacion = prev !== null ? s.totalARS - prev : null;
              return (
                <div key={s.mes} className={`flex items-center justify-between text-sm py-1 border-b ${t.divide} last:border-0`}>
                  <span className={t.surfaceText}>{s.mes}</span>
                  <div className="text-right">
                    <span className={`font-medium ${t.text}`}>{fmtMoney(s.totalARS)}</span>
                    {variacion !== null && <span className={`text-xs ml-2 ${variacion >= 0 ? 'text-green-500' : 'text-red-500'}`}>{variacion >= 0 ? '+' : ''}{fmtMoneyShort(variacion)}</span>}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// ============ NUEVO MOVIMIENTO ============
function NuevoMovimiento({ agregar, config, t }) {
  const [tipo, setTipo] = useState('egreso');
  const [fecha, setFecha] = useState(hoy());
  const [monto, setMonto] = useState('');
  const [moneda, setMoneda] = useState('ARS');
  const [cotizacion, setCotizacion] = useState(config.cotizacionDolar);
  const [categoria, setCategoria] = useState('Supermercado');
  const [subcategoria, setSubcategoria] = useState('');
  const [medioPago, setMedioPago] = useState('Débito');
  const [cuentaOrigen, setCuentaOrigen] = useState('caja_gal_ars');
  const [cuentaDestino, setCuentaDestino] = useState('cash_usd');
  const [montoDestino, setMontoDestino] = useState('');
  const [descripcion, setDescripcion] = useState('');
  const [cuotas, setCuotas] = useState(2);
  const [esCuotas, setEsCuotas] = useState(false);
  const [mensaje, setMensaje] = useState('');

  const categorias = tipo === 'ingreso' ? CATEGORIAS_INGRESO : CATEGORIAS_EGRESO;
  const subcategoriasDisponibles = categorias[categoria] || [];

  useEffect(() => { const cats = Object.keys(categorias); if (!cats.includes(categoria)) setCategoria(cats[0]); }, [tipo]);
  useEffect(() => { if (subcategoriasDisponibles.length > 0 && !subcategoriasDisponibles.includes(subcategoria)) setSubcategoria(subcategoriasDisponibles[0]); }, [categoria]);

  const handleSubmit = () => {
    if (!monto || Number(monto) <= 0) { setMensaje('error:Ingresá un monto válido'); return; }
    if (tipo === 'transferencia') {
      const co = CUENTAS_DEFAULT.find((c) => c.id === cuentaOrigen);
      const cd = CUENTAS_DEFAULT.find((c) => c.id === cuentaDestino);
      agregar({ tipo: 'transferencia', fecha, monto: Number(monto), moneda: co.moneda, cotizacion: Number(cotizacion), cuentaOrigen, cuentaDestino, montoDestino: Number(montoDestino) || Number(monto), monedaDestino: cd.moneda, descripcion: descripcion || `${co.nombre} → ${cd.nombre}`, categoria: 'Transferencia' });
    } else {
      const mov = { tipo, fecha, monto: Number(monto), moneda, cotizacion: moneda === 'USD' ? Number(cotizacion) : null, categoria, subcategoria, descripcion };
      if (tipo === 'egreso') { mov.medioPago = medioPago; if (esCuotas && cuotas > 1) mov.cuotas = Number(cuotas); }
      agregar(mov);
    }
    setMensaje('ok:✓ Movimiento registrado');
    setMonto(''); setMontoDestino(''); setDescripcion(''); setCuotas(2); setEsCuotas(false);
    setTimeout(() => setMensaje(''), 1500);
  };

  const tipoBtns = [
    { id: 'egreso', label: 'Egreso', icon: TrendingDown, c: 'red' },
    { id: 'ingreso', label: 'Ingreso', icon: TrendingUp, c: 'green' },
    { id: 'transferencia', label: 'Transfer.', icon: ArrowRightLeft, c: 'blue' },
  ];

  return (
    <div className="space-y-3">
      <h3 className={`text-base font-semibold ${t.text} px-1`}>Nuevo movimiento</h3>
      <div className="grid grid-cols-3 gap-2">
        {tipoBtns.map((tb) => {
          const activo = tipo === tb.id;
          const cls = activo
            ? tb.c === 'red' ? 'border-red-500 bg-red-500/10 text-red-500' : tb.c === 'green' ? 'border-green-500 bg-green-500/10 text-green-500' : 'border-blue-500 bg-blue-500/10 text-blue-500'
            : `${t.cardBorder} ${t.card} ${t.textMuted}`;
          return (
            <button key={tb.id} onClick={() => setTipo(tb.id)} className={`flex flex-col items-center gap-1 py-3 rounded-xl border-2 ${cls}`}>
              <tb.icon size={20} /><span className="text-xs font-medium">{tb.label}</span>
            </button>
          );
        })}
      </div>

      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3 space-y-3`}>
        {(tipo === 'egreso' || tipo === 'ingreso') && (
          <>
            <Field label="Monto" t={t}>
              <div className="flex gap-2">
                <input type="number" inputMode="decimal" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" className={`flex-1 px-3 py-3 border rounded-lg text-2xl font-bold ${t.input}`} />
                <select value={moneda} onChange={(e) => setMoneda(e.target.value)} className={`px-3 py-3 border rounded-lg text-base font-medium ${t.input}`}>
                  <option value="ARS">ARS</option><option value="USD">USD</option>
                </select>
              </div>
            </Field>
            {moneda === 'USD' && <Field label="Cotización USD" t={t}><input type="number" inputMode="numeric" value={cotizacion} onChange={(e) => setCotizacion(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>}
            <Field label="Fecha" t={t}><input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
            <Field label="Categoría" t={t}>
              <select value={categoria} onChange={(e) => setCategoria(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`}>
                {Object.keys(categorias).map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            {subcategoriasDisponibles.length > 0 && (
              <Field label="Subcategoría" t={t}>
                <select value={subcategoria} onChange={(e) => setSubcategoria(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`}>
                  {subcategoriasDisponibles.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </Field>
            )}
            {tipo === 'egreso' && (
              <>
                <Field label="Medio de pago" t={t}>
                  <div className="grid grid-cols-3 gap-1.5">
                    {MEDIOS_PAGO.map((m) => (
                      <button key={m} onClick={() => setMedioPago(m)} className={`py-2 rounded-lg text-xs font-medium border ${medioPago === m ? 'border-blue-500 bg-blue-500/10 text-blue-500' : `${t.cardBorder} ${t.card} ${t.textMuted}`}`}>{m}</button>
                    ))}
                  </div>
                </Field>
                <div className="bg-amber-500/10 border border-amber-500/30 rounded-lg p-3">
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={esCuotas} onChange={(e) => setEsCuotas(e.target.checked)} className="w-4 h-4 rounded" />
                    <span className={`text-sm font-medium ${t.surfaceText}`}>¿Es compra en cuotas?</span>
                  </label>
                  {esCuotas && (
                    <div className="mt-2 flex items-center gap-2 flex-wrap">
                      <span className={`text-sm ${t.textMuted}`}>Cuotas:</span>
                      <input type="number" inputMode="numeric" value={cuotas} onChange={(e) => setCuotas(e.target.value)} min="2" max="60" className={`w-16 px-2 py-1.5 border rounded ${t.input}`} />
                      {monto > 0 && cuotas > 1 && <span className={`text-sm ${t.textMuted}`}>de {fmtMoney(Number(monto) / Number(cuotas), moneda)} c/u</span>}
                    </div>
                  )}
                </div>
              </>
            )}
            <Field label="Descripción (opcional)" t={t}><input type="text" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder="Detalle" className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
          </>
        )}
        {tipo === 'transferencia' && (
          <>
            <Field label="Fecha" t={t}><input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
            <Field label="Cuenta origen" t={t}><select value={cuentaOrigen} onChange={(e) => setCuentaOrigen(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`}>{CUENTAS_DEFAULT.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></Field>
            <Field label={`Monto origen (${CUENTAS_DEFAULT.find((c) => c.id === cuentaOrigen)?.moneda})`} t={t}><input type="number" inputMode="decimal" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" className={`w-full px-3 py-2.5 border rounded-lg text-lg font-semibold ${t.input}`} /></Field>
            <Field label="Cuenta destino" t={t}><select value={cuentaDestino} onChange={(e) => setCuentaDestino(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`}>{CUENTAS_DEFAULT.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></Field>
            <Field label={`Monto destino (${CUENTAS_DEFAULT.find((c) => c.id === cuentaDestino)?.moneda})`} t={t}><input type="number" inputMode="decimal" value={montoDestino} onChange={(e) => setMontoDestino(e.target.value)} placeholder="Si es misma moneda, dejá vacío" className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
            {monto && montoDestino && CUENTAS_DEFAULT.find((c) => c.id === cuentaOrigen)?.moneda !== CUENTAS_DEFAULT.find((c) => c.id === cuentaDestino)?.moneda && (
              <div className="bg-blue-500/10 border border-blue-500/30 rounded-lg p-2.5 text-sm text-blue-500">Tipo de cambio: <strong>{(Number(monto) / Number(montoDestino)).toFixed(2)}</strong></div>
            )}
            <Field label="Descripción" t={t}><input type="text" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder="Ej: Compra dólares" className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
          </>
        )}
      </div>

      <button onClick={handleSubmit} className="w-full bg-blue-600 active:bg-blue-700 text-white font-semibold py-3.5 rounded-xl flex items-center justify-center gap-2 text-base"><Plus size={20} /> Registrar</button>
      {mensaje && <div className={`text-center text-sm font-medium py-2 rounded-lg ${mensaje.startsWith('ok') ? 'text-green-600 bg-green-500/10' : 'text-red-600 bg-red-500/10'}`}>{mensaje.split(':')[1]}</div>}
    </div>
  );
}

function Field({ label, children, t }) {
  return (
    <label className="block">
      <span className={`text-xs font-medium ${t.textMuted} block mb-1`}>{label}</span>
      {children}
    </label>
  );
}

// ============ MOVIMIENTOS ============
function Movimientos({ movimientos, eliminar, actualizar, config, t }) {
  const [filtroTipo, setFiltroTipo] = useState('todos');
  const [filtroMes, setFiltroMes] = useState(mesAnio(hoy()));
  const [editando, setEditando] = useState(null);

  const filtrados = useMemo(() => movimientos
    .filter((m) => filtroTipo === 'todos' || m.tipo === filtroTipo)
    .filter((m) => filtroMes === 'todos' || mesAnio(m.fecha) === filtroMes)
    .sort((a, b) => b.fecha.localeCompare(a.fecha)), [movimientos, filtroTipo, filtroMes]);

  const mesesDisponibles = useMemo(() => Array.from(new Set(movimientos.map((m) => mesAnio(m.fecha)))).sort().reverse(), [movimientos]);

  return (
    <div className="space-y-3">
      <h3 className={`text-base font-semibold ${t.text} px-1`}>Movimientos ({filtrados.length})</h3>
      <div className="grid grid-cols-2 gap-2">
        <select value={filtroTipo} onChange={(e) => setFiltroTipo(e.target.value)} className={`px-3 py-2 border rounded-lg text-sm ${t.input}`}>
          <option value="todos">Todos los tipos</option><option value="ingreso">Ingresos</option><option value="egreso">Egresos</option><option value="transferencia">Transferencias</option>
        </select>
        <select value={filtroMes} onChange={(e) => setFiltroMes(e.target.value)} className={`px-3 py-2 border rounded-lg text-sm ${t.input}`}>
          <option value="todos">Todos los meses</option>{mesesDisponibles.map((m) => <option key={m} value={m}>{m}</option>)}
        </select>
      </div>
      <div className="space-y-2">
        {filtrados.length === 0 && <div className={`text-center py-8 ${t.textSoft} text-sm`}>No hay movimientos</div>}
        {filtrados.map((m) => <MovimientoCard key={m.id} mov={m} eliminar={eliminar} config={config} onEdit={() => setEditando(m)} t={t} />)}
      </div>
      {editando && <EditarModal mov={editando} onClose={() => setEditando(null)} onSave={(cambios) => { actualizar(editando.id, cambios); setEditando(null); }} t={t} />}
    </div>
  );
}

function MovimientoCard({ mov, eliminar, config, onEdit, t }) {
  const tipoColor = mov.tipo === 'ingreso' ? 'text-green-500' : mov.tipo === 'egreso' ? 'text-red-500' : 'text-blue-500';
  const tipoIcon = mov.tipo === 'ingreso' ? '+' : mov.tipo === 'egreso' ? '−' : '↔';
  return (
    <div className={`flex items-center gap-2 p-3 ${t.card} border ${t.cardBorder} rounded-xl`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className={`text-base font-bold ${tipoColor}`}>{tipoIcon}</span>
          <span className={`font-semibold ${t.text} text-sm`}>{fmtMoney(mov.monto, mov.moneda)}</span>
          {mov.moneda === 'USD' && <span className={`text-xs ${t.textSoft}`}>≈{fmtMoney(montoEnARS(mov, config))}</span>}
          <span className={`text-[11px] px-1.5 py-0.5 ${t.chipBg} rounded-full ${t.surfaceText}`}>{mov.categoria}</span>
        </div>
        <div className={`text-[11px] ${t.textSoft} mt-0.5 flex items-center gap-1.5 flex-wrap`}>
          <span>{mov.fecha}</span>
          {mov.subcategoria && <span>• {mov.subcategoria}</span>}
          {mov.medioPago && <span>• {mov.medioPago}</span>}
          {mov.descripcion && <span>• {mov.descripcion}</span>}
          {mov.totalCuotas && <span className="text-amber-500">• {mov.numeroCuota}/{mov.totalCuotas}</span>}
        </div>
      </div>
      {!mov.esCuota && <button onClick={onEdit} className={`p-2 ${t.textSoft} active:text-blue-500`}><Edit2 size={15} /></button>}
      <button onClick={() => eliminar(mov.id)} className={`p-2 ${t.textSoft} active:text-red-500`}><Trash2 size={15} /></button>
    </div>
  );
}

function EditarModal({ mov, onClose, onSave, t }) {
  const [monto, setMonto] = useState(mov.monto);
  const [descripcion, setDescripcion] = useState(mov.descripcion || '');
  const [fecha, setFecha] = useState(mov.fecha);
  return (
    <div className="fixed inset-0 bg-black/50 flex items-end sm:items-center justify-center z-50 p-3" onClick={onClose}>
      <div className={`${t.card} rounded-2xl w-full max-w-md p-4 space-y-3`} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between">
          <h3 className={`text-base font-semibold ${t.text}`}>Editar movimiento</h3>
          <button onClick={onClose} className={t.textSoft}><X size={20} /></button>
        </div>
        <Field label="Monto" t={t}><input type="number" inputMode="decimal" value={monto} onChange={(e) => setMonto(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg text-lg font-semibold ${t.input}`} /></Field>
        <Field label="Fecha" t={t}><input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
        <Field label="Descripción" t={t}><input type="text" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} className={`w-full px-3 py-2.5 border rounded-lg ${t.input}`} /></Field>
        <button onClick={() => onSave({ monto: Number(monto), descripcion, fecha })} className="w-full bg-blue-600 active:bg-blue-700 text-white font-semibold py-3 rounded-xl flex items-center justify-center gap-2"><Check size={18} /> Guardar</button>
      </div>
    </div>
  );
}

// ============ MÁS ============
function Mas({ presupuesto, actualizar, movimientos, config, patrimonio, snapshots, t }) {
  const [vista, setVista] = useState('menu');
  if (vista === 'presupuesto') return <PresupuestoMobile presupuesto={presupuesto} actualizar={actualizar} onBack={() => setVista('menu')} t={t} />;
  if (vista === 'cuotas') return <Cuotas movimientos={movimientos} onBack={() => setVista('menu')} t={t} />;
  if (vista === 'exportar') return <Exportar movimientos={movimientos} presupuesto={presupuesto} config={config} patrimonio={patrimonio} snapshots={snapshots} onBack={() => setVista('menu')} t={t} />;

  const items = [
    { id: 'presupuesto', icon: Target, color: 'blue', titulo: 'Presupuesto mensual', desc: 'Planificá cuánto gastar por categoría' },
    { id: 'cuotas', icon: Calendar, color: 'amber', titulo: 'Cuotas activas', desc: 'Seguimiento de compras en cuotas' },
    { id: 'exportar', icon: FileSpreadsheet, color: 'green', titulo: 'Exportar a Excel', desc: 'Descargá todo para analizar en la compu' },
  ];

  return (
    <div className="space-y-3">
      <h3 className={`text-base font-semibold ${t.text} px-1`}>Más opciones</h3>
      {items.map((it) => (
        <button key={it.id} onClick={() => setVista(it.id)} className={`w-full ${t.card} border ${t.cardBorder} rounded-xl p-4 flex items-center gap-3`}>
          <div className={`w-10 h-10 rounded-lg bg-${it.color}-500/10 flex items-center justify-center`}><it.icon size={20} className={`text-${it.color}-500`} /></div>
          <div className="text-left flex-1">
            <div className={`font-medium ${t.text} text-sm`}>{it.titulo}</div>
            <div className={`text-xs ${t.textSoft}`}>{it.desc}</div>
          </div>
          <ChevronRight size={18} className={t.textSoft} />
        </button>
      ))}
    </div>
  );
}

function PresupuestoMobile({ presupuesto, actualizar, onBack, t }) {
  const [mesSel, setMesSel] = useState(mesAnio(hoy()));
  const [copiarMsg, setCopiarMsg] = useState('');
  const cambiarMes = (delta) => {
    const [y, m] = mesSel.split('-').map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    setMesSel(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  };
  const [y, mNum] = mesSel.split('-').map(Number);
  const datosMes = presupuesto[mesSel] || {};
  const totalMes = Object.values(datosMes).reduce((s, v) => s + (v || 0), 0);
  const copiarMesAnterior = () => {
    const [yy, mm] = mesSel.split('-').map(Number);
    const dPrev = new Date(yy, mm - 2, 1);
    const keyPrev = `${dPrev.getFullYear()}-${String(dPrev.getMonth() + 1).padStart(2, '0')}`;
    const datosPrev = presupuesto[keyPrev] || {};
    if (Object.keys(datosPrev).length === 0) { setCopiarMsg('El mes anterior no tiene presupuesto'); setTimeout(() => setCopiarMsg(''), 2000); return; }
    Object.keys(datosPrev).forEach((cat) => actualizar(mesSel, cat, datosPrev[cat]));
    setCopiarMsg('✓ Copiado del mes anterior');
    setTimeout(() => setCopiarMsg(''), 2000);
  };
  return (
    <div className="space-y-3">
      <div className="flex items-center gap-2 px-1">
        <button onClick={onBack} className={`p-1.5 -ml-1.5 ${t.textMuted}`}><ChevronLeft size={22} /></button>
        <h3 className={`text-base font-semibold ${t.text}`}>Presupuesto</h3>
      </div>
      <div className={`${t.card} rounded-xl border ${t.cardBorder} px-3 py-2.5 flex items-center justify-between`}>
        <button onClick={() => cambiarMes(-1)} className={`p-2 ${t.textMuted}`}><ChevronLeft size={20} /></button>
        <h2 className={`text-base font-bold ${t.text}`}>{MESES[mNum - 1]} {y}</h2>
        <button onClick={() => cambiarMes(1)} className={`p-2 ${t.textMuted}`}><ChevronRight size={20} /></button>
      </div>
      <button onClick={copiarMesAnterior} className="w-full bg-blue-500/10 border border-blue-500/30 text-blue-500 rounded-lg py-2 text-sm font-medium">Copiar del mes anterior</button>
      {copiarMsg && <div className={`text-center text-xs ${t.textMuted}`}>{copiarMsg}</div>}
      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3 space-y-2`}>
        {Object.keys(CATEGORIAS_EGRESO).map((cat) => (
          <div key={cat} className="flex items-center justify-between gap-2">
            <span className={`text-sm ${t.surfaceText} flex-1`}>{cat}</span>
            <input type="number" inputMode="numeric" value={datosMes[cat] || ''} onChange={(e) => actualizar(mesSel, cat, Number(e.target.value) || 0)} placeholder="0" className={`w-28 px-2 py-1.5 text-right border rounded-lg text-sm ${t.input}`} />
          </div>
        ))}
        <div className={`flex items-center justify-between pt-2 mt-2 border-t ${t.divide} font-bold`}>
          <span className={`text-sm ${t.text}`}>Total mes</span>
          <span className={`text-sm ${t.text}`}>{fmtMoney(totalMes)}</span>
        </div>
      </div>
    </div>
  );
}

function Cuotas({ movimientos, onBack, t }) {
  const cuotasGrupos = useMemo(() => {
    const grupos = {};
    movimientos.filter((m) => m.grupoCuota).forEach((m) => {
      if (!grupos[m.grupoCuota]) grupos[m.grupoCuota] = { descripcion: m.descripcion?.split('(')[0]?.trim() || m.categoria, categoria: m.categoria, medioPago: m.medioPago, montoTotal: 0, montoCuota: m.monto, totalCuotas: m.totalCuotas, fechaInicio: m.fecha, fechaFin: '', cuotasPagadas: 0, cuotasPendientes: 0 };
      grupos[m.grupoCuota].montoTotal += m.monto;
      if (m.fecha <= hoy()) grupos[m.grupoCuota].cuotasPagadas += 1; else grupos[m.grupoCuota].cuotasPendientes += 1;
      if (m.fecha > grupos[m.grupoCuota].fechaFin) grupos[m.grupoCuota].fechaFin = m.fecha;
    });
    return Object.values(grupos).sort((a, b) => b.fechaInicio.localeCompare(a.fechaInicio));
  }, [movimientos]);
  const pendientesTotal = cuotasGrupos.reduce((s, g) => s + g.cuotasPendientes * g.montoCuota, 0);
  return (
    <div className="space-y-3">
      <div className="flex items-center gap-2 px-1">
        <button onClick={onBack} className={`p-1.5 -ml-1.5 ${t.textMuted}`}><ChevronLeft size={22} /></button>
        <h3 className={`text-base font-semibold ${t.text}`}>Cuotas activas</h3>
      </div>
      <div className="bg-amber-500/10 border border-amber-500/30 rounded-xl p-3 text-center">
        <div className="text-xs text-amber-600">Pendiente total en cuotas</div>
        <div className="text-xl font-bold text-amber-600">{fmtMoney(pendientesTotal)}</div>
      </div>
      {cuotasGrupos.length === 0 && <div className={`text-center py-8 ${t.textSoft} text-sm`}>No hay compras en cuotas</div>}
      <div className="space-y-2">
        {cuotasGrupos.map((g, i) => (
          <div key={i} className={`${t.card} border ${t.cardBorder} rounded-xl p-3`}>
            <div className="flex items-start justify-between gap-2">
              <div className="min-w-0"><div className={`font-semibold ${t.text} text-sm truncate`}>{g.descripcion}</div><div className={`text-xs ${t.textSoft}`}>{g.categoria} • {g.medioPago}</div></div>
              <div className="text-right shrink-0"><div className={`font-semibold text-sm ${t.text}`}>{fmtMoney(g.montoCuota)} × {g.totalCuotas}</div><div className={`text-xs ${t.textSoft}`}>Total: {fmtMoney(g.montoTotal)}</div></div>
            </div>
            <div className="mt-2 flex items-center gap-2">
              <div className={`flex-1 ${t.chipBg} rounded-full h-2 overflow-hidden`}><div className="bg-blue-500 h-full" style={{ width: `${(g.cuotasPagadas / g.totalCuotas) * 100}%` }} /></div>
              <span className={`text-xs ${t.textMuted}`}>{g.cuotasPagadas}/{g.totalCuotas}</span>
            </div>
            <div className={`text-[11px] ${t.textSoft} mt-1`}>{g.fechaInicio} → {g.fechaFin}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Exportar({ movimientos, presupuesto, config, patrimonio, snapshots, onBack, t }) {
  const exportarExcel = () => {
    const wb = XLSX.utils.book_new();
    const movsData = movimientos.map((m) => ({ Fecha: m.fecha, Tipo: m.tipo, Categoría: m.categoria, Subcategoría: m.subcategoria || '', Descripción: m.descripcion || '', Monto: m.monto, Moneda: m.moneda, Cotización: m.cotizacion || '', 'Monto ARS': montoEnARS(m, config), 'Medio Pago': m.medioPago || '', 'Cuenta Origen': m.cuentaOrigen || '', 'Cuenta Destino': m.cuentaDestino || '', 'Es cuota': m.esCuota ? 'Sí' : 'No', 'N° Cuota': m.numeroCuota || '', 'Total Cuotas': m.totalCuotas || '' }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(movsData), 'Movimientos');

    const meses = [...new Set([...Object.keys(presupuesto), ...movimientos.map((m) => mesAnio(m.fecha))])].sort();
    const categorias = Object.keys(CATEGORIAS_EGRESO);
    const seguimiento = [];
    const header = ['Categoría'];
    meses.forEach((m) => header.push(`${m} Plan`, `${m} Real`, `${m} %`));
    seguimiento.push(header);
    categorias.forEach((cat) => {
      const row = [cat];
      meses.forEach((m) => {
        const plan = (presupuesto[m] || {})[cat] || 0;
        const real = movimientos.filter((mv) => mv.tipo === 'egreso' && mv.categoria === cat && mesAnio(mv.fecha) === m).reduce((s, mv) => s + montoEnARS(mv, config), 0);
        row.push(plan, real, plan > 0 ? real / plan : 0);
      });
      seguimiento.push(row);
    });
    const totalRow = ['TOTAL'];
    meses.forEach((m) => {
      const tp = Object.values(presupuesto[m] || {}).reduce((s, v) => s + v, 0);
      const tr = movimientos.filter((mv) => mv.tipo === 'egreso' && mesAnio(mv.fecha) === m).reduce((s, mv) => s + montoEnARS(mv, config), 0);
      totalRow.push(tp, tr, tp > 0 ? tr / tp : 0);
    });
    seguimiento.push(totalRow);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(seguimiento), 'Seguimiento');

    const presData = [['Categoría', ...meses, 'Anual']];
    categorias.forEach((cat) => { const row = [cat]; let anual = 0; meses.forEach((m) => { const v = (presupuesto[m] || {})[cat] || 0; row.push(v); anual += v; }); row.push(anual); presData.push(row); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(presData), 'Presupuesto');

    const resumen = {};
    movimientos.filter((m) => m.tipo === 'egreso').forEach((m) => { const key = `${m.categoria} - ${m.subcategoria || 'General'}`; resumen[key] = (resumen[key] || 0) + montoEnARS(m, config); });
    const resumenData = [['Categoría/Subcategoría', 'Total ARS']];
    Object.entries(resumen).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => resumenData.push([k, v]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), 'Resumen Cat');

    // Hoja patrimonio
    const patData = [['Cuenta', 'Moneda', 'Tipo', 'Saldo', 'Saldo ARS']];
    CUENTAS_DEFAULT.forEach((c) => { const saldo = patrimonio[c.id] || 0; patData.push([c.nombre, c.moneda, c.tipo, saldo, c.moneda === 'USD' ? saldo * config.cotizacionDolar : saldo]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(patData), 'Patrimonio');

    // Hoja evolución
    if (snapshots.length > 0) {
      const snapData = [['Mes', 'Fecha', 'Total ARS', 'Total USD']];
      snapshots.forEach((s) => snapData.push([s.mes, s.fecha, s.totalARS, s.totalUSD]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(snapData), 'Evolución Patrim');
    }

    XLSX.writeFile(wb, `Finanzas_Agustin_${hoy()}.xlsx`);
  };
  return (
    <div className="space-y-3">
      <div className="flex items-center gap-2 px-1">
        <button onClick={onBack} className={`p-1.5 -ml-1.5 ${t.textMuted}`}><ChevronLeft size={22} /></button>
        <h3 className={`text-base font-semibold ${t.text}`}>Exportar a Excel</h3>
      </div>
      <div className={`${t.card} rounded-xl border ${t.cardBorder} p-3`}>
        <p className={`text-sm ${t.textMuted} mb-2`}>El Excel incluye:</p>
        <ul className={`text-sm ${t.surfaceText} space-y-1.5`}>
          <li>• <strong>Movimientos</strong>: todos los registros</li>
          <li>• <strong>Seguimiento</strong>: Plan vs Real por mes</li>
          <li>• <strong>Presupuesto</strong>: plan anual</li>
          <li>• <strong>Resumen Cat</strong>: total por subcategoría</li>
          <li>• <strong>Patrimonio</strong>: saldos por cuenta</li>
          <li>• <strong>Evolución</strong>: histórico de patrimonio</li>
        </ul>
      </div>
      <button onClick={exportarExcel} className="w-full bg-green-600 active:bg-green-700 text-white font-semibold py-3.5 rounded-xl flex items-center justify-center gap-2"><Download size={20} /> Descargar Excel</button>
      <div className={`text-center text-xs ${t.textSoft}`}>{movimientos.length} movimientos guardados</div>
    </div>
  );
}
