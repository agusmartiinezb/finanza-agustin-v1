import React, { useState, useEffect, useMemo } from 'react';
import {
  Plus,
  Trash2,
  Download,
  TrendingUp,
  TrendingDown,
  ArrowRightLeft,
  Calendar,
  Wallet,
  Target,
  FileSpreadsheet,
  Edit2,
  X,
  Check,
  ChevronLeft,
  ChevronRight,
} from 'lucide-react';
import * as XLSX from 'xlsx';

// ============ CONFIGURACIÓN BASE ============
const CUENTAS_DEFAULT = [
  { id: 'caja_gal_ars', nombre: 'Caja GAL ARS', moneda: 'ARS' },
  { id: 'caja_gal_usd', nombre: 'Caja GAL USD', moneda: 'USD' },
  { id: 'fima_gal_ars', nombre: 'FIMA GAL ARS', moneda: 'ARS' },
  { id: 'fima_gal_usd', nombre: 'FIMA GAL USD', moneda: 'USD' },
  { id: 'mp_ars', nombre: 'MP ARS', moneda: 'ARS' },
  { id: 'mp_usd', nombre: 'MP USD', moneda: 'USD' },
  { id: 'balanz', nombre: 'Balanz', moneda: 'USD' },
  { id: 'binance', nombre: 'Binance', moneda: 'USD' },
  { id: 'cash_ars', nombre: 'Cash ARS', moneda: 'ARS' },
  { id: 'cash_usd', nombre: 'Cash USD', moneda: 'USD' },
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
  if (abs >= 1000000) return `$ ${(n / 1000000).toFixed(1)}M`;
  if (abs >= 1000) return `$ ${Math.round(n / 1000)}k`;
  return `$ ${Math.round(n)}`;
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

// ============ STORAGE (localStorage — funciona en celular y web) ============
const STORAGE_KEYS = {
  movimientos: 'finanzas_movimientos_v2',
  presupuesto: 'finanzas_presupuesto_v2',
  config: 'finanzas_config_v2',
};

const loadData = (key, defaultValue) => {
  try {
    const raw = localStorage.getItem(key);
    if (raw) return JSON.parse(raw);
  } catch (e) {
    console.error('Error cargando', e);
  }
  return defaultValue;
};

const saveData = (key, value) => {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.error('Error guardando', e);
  }
};

function montoEnARS(m, config) {
  if (m.moneda === 'USD') return m.monto * (m.cotizacion || config.cotizacionDolar);
  return m.monto;
}

// ============ COMPONENTE PRINCIPAL ============
export default function FinanzasApp() {
  const [tab, setTab] = useState('dashboard');
  const [movimientos, setMovimientos] = useState([]);
  const [presupuesto, setPresupuesto] = useState({});
  const [config, setConfig] = useState({ cotizacionDolar: 1400, mesActivo: mesAnio(hoy()) });
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    setMovimientos(loadData(STORAGE_KEYS.movimientos, []));
    setPresupuesto(loadData(STORAGE_KEYS.presupuesto, {}));
    setConfig(loadData(STORAGE_KEYS.config, { cotizacionDolar: 1400, mesActivo: mesAnio(hoy()) }));
    setLoading(false);
  }, []);

  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.movimientos, movimientos); }, [movimientos, loading]);
  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.presupuesto, presupuesto); }, [presupuesto, loading]);
  useEffect(() => { if (!loading) saveData(STORAGE_KEYS.config, config); }, [config, loading]);

  const agregarMovimiento = (mov) => {
    const id = Date.now() + Math.random();
    if (mov.tipo === 'egreso' && mov.cuotas && mov.cuotas > 1) {
      const movs = [];
      const montoCuota = mov.monto / mov.cuotas;
      for (let i = 0; i < mov.cuotas; i++) {
        movs.push({
          ...mov,
          id: id + i,
          monto: montoCuota,
          fecha: sumarMeses(mov.fecha, i),
          descripcion: `${mov.descripcion || mov.categoria} (${i + 1}/${mov.cuotas})`,
          esCuota: true,
          grupoCuota: id,
          numeroCuota: i + 1,
          totalCuotas: mov.cuotas,
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

  if (loading) {
    return <div className="min-h-screen bg-slate-50 flex items-center justify-center"><div className="text-slate-600">Cargando...</div></div>;
  }

  const tabs = [
    { id: 'dashboard', label: 'Inicio', icon: TrendingUp },
    { id: 'movimientos', label: 'Movim.', icon: Wallet },
    { id: 'nuevo', label: 'Nuevo', icon: Plus },
    { id: 'cuotas', label: 'Cuotas', icon: Calendar },
    { id: 'mas', label: 'Más', icon: Target },
  ];

  return (
    <div className="min-h-screen bg-slate-50 pb-20">
      {/* Header compacto */}
      <div className="bg-white border-b border-slate-200 px-4 py-3 sticky top-0 z-10">
        <div className="flex items-center justify-between">
          <h1 className="text-lg font-bold text-slate-900">Finanzas Agustín</h1>
          <div className="flex items-center gap-1.5">
            <span className="text-xs text-slate-500">USD</span>
            <input
              type="number"
              inputMode="numeric"
              value={config.cotizacionDolar}
              onChange={(e) => setConfig({ ...config, cotizacionDolar: Number(e.target.value) })}
              className="w-20 px-2 py-1 text-right border border-slate-300 rounded text-sm"
            />
          </div>
        </div>
      </div>

      {/* Contenido */}
      <div className="px-3 py-3 max-w-2xl mx-auto">
        {tab === 'dashboard' && <Dashboard movimientos={movimientos} presupuesto={presupuesto} config={config} setConfig={setConfig} />}
        {tab === 'movimientos' && <Movimientos movimientos={movimientos} eliminar={eliminarMovimiento} actualizar={actualizarMovimiento} config={config} />}
        {tab === 'nuevo' && <NuevoMovimiento agregar={agregarMovimiento} config={config} onDone={() => setTab('dashboard')} />}
        {tab === 'cuotas' && <Cuotas movimientos={movimientos} />}
        {tab === 'mas' && <Mas presupuesto={presupuesto} actualizar={actualizarPresupuesto} movimientos={movimientos} config={config} />}
      </div>

      {/* Navegación inferior fija (estilo app nativa) */}
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 z-20">
        <div className="flex max-w-2xl mx-auto">
          {tabs.map((t) => {
            const activo = tab === t.id;
            const esNuevo = t.id === 'nuevo';
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className={`flex-1 flex flex-col items-center justify-center gap-0.5 py-2 ${activo ? 'text-blue-600' : 'text-slate-500'}`}
              >
                {esNuevo ? (
                  <div className={`-mt-5 w-12 h-12 rounded-full flex items-center justify-center shadow-lg ${activo ? 'bg-blue-700' : 'bg-blue-600'}`}>
                    <t.icon size={24} className="text-white" />
                  </div>
                ) : (
                  <t.icon size={22} />
                )}
                <span className="text-[10px] font-medium">{t.label}</span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ============ DASHBOARD ============
function Dashboard({ movimientos, presupuesto, config, setConfig }) {
  const mesActivo = config.mesActivo || mesAnio(hoy());

  const cambiarMes = (delta) => {
    const [y, m] = mesActivo.split('-').map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    setConfig({ ...config, mesActivo: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}` });
  };

  const movsDelMes = useMemo(() => movimientos.filter((m) => mesAnio(m.fecha) === mesActivo), [movimientos, mesActivo]);
  const totalIngresos = useMemo(() => movsDelMes.filter((m) => m.tipo === 'ingreso').reduce((s, m) => s + montoEnARS(m, config), 0), [movsDelMes, config]);
  const totalEgresos = useMemo(() => movsDelMes.filter((m) => m.tipo === 'egreso').reduce((s, m) => s + montoEnARS(m, config), 0), [movsDelMes, config]);

  const egresosPorCategoria = useMemo(() => {
    const acc = {};
    movsDelMes.filter((m) => m.tipo === 'egreso').forEach((m) => { acc[m.categoria] = (acc[m.categoria] || 0) + montoEnARS(m, config); });
    return acc;
  }, [movsDelMes, config]);

  const presupuestoMes = presupuesto[mesActivo] || {};
  const totalPlanificado = Object.values(presupuestoMes).reduce((s, v) => s + (v || 0), 0);
  const balance = totalIngresos - totalEgresos;
  const [y, mNum] = mesActivo.split('-').map(Number);
  const nombreMes = MESES[mNum - 1];

  // Bug arreglado: solo calcula % si hay plan
  const cumplimiento = totalPlanificado > 0 ? totalEgresos / totalPlanificado : null;

  const categoriasConDatos = Object.keys(CATEGORIAS_EGRESO).filter(
    (cat) => (presupuestoMes[cat] || 0) > 0 || (egresosPorCategoria[cat] || 0) > 0
  );

  return (
    <div className="space-y-3">
      {/* Selector de mes */}
      <div className="bg-white rounded-xl border border-slate-200 px-3 py-2.5 flex items-center justify-between">
        <button onClick={() => cambiarMes(-1)} className="p-2 rounded-lg hover:bg-slate-100 active:bg-slate-200"><ChevronLeft size={20} className="text-slate-600" /></button>
        <h2 className="text-base font-bold text-slate-900">{nombreMes} {y}</h2>
        <button onClick={() => cambiarMes(1)} className="p-2 rounded-lg hover:bg-slate-100 active:bg-slate-200"><ChevronRight size={20} className="text-slate-600" /></button>
      </div>

      {/* KPIs en grilla 2x2 */}
      <div className="grid grid-cols-2 gap-2">
        <KPI label="Ingresos" valor={totalIngresos} color="green" />
        <KPI label="Egresos" valor={totalEgresos} color="red" />
        <KPI label="Balance" valor={balance} color={balance >= 0 ? 'green' : 'red'} />
        <KPI label="Cumplimiento plan" valor={cumplimiento} esPct color={cumplimiento === null ? 'slate' : cumplimiento > 1 ? 'red' : 'green'} />
      </div>

      {/* Seguimiento por categoría — versión mobile simplificada */}
      <div className="bg-white rounded-xl border border-slate-200 p-3">
        <h3 className="text-sm font-semibold text-slate-900 mb-2">Seguimiento por categoría</h3>
        {categoriasConDatos.length === 0 && (
          <p className="text-sm text-slate-400 text-center py-4">Sin movimientos este mes</p>
        )}
        <div className="space-y-2.5">
          {categoriasConDatos.map((cat) => {
            const plan = presupuestoMes[cat] || 0;
            const real = egresosPorCategoria[cat] || 0;
            const pct = plan > 0 ? real / plan : null;
            const pctBarra = plan > 0 ? Math.min((real / plan) * 100, 100) : (real > 0 ? 100 : 0);
            const colorBarra = pct === null ? 'bg-slate-400' : pct > 1 ? 'bg-red-500' : pct > 0.9 ? 'bg-amber-500' : 'bg-green-500';
            return (
              <div key={cat}>
                <div className="flex items-center justify-between text-sm mb-1">
                  <span className="font-medium text-slate-700">{cat}</span>
                  <span className="text-slate-600">
                    {fmtMoneyShort(real)}
                    {plan > 0 && <span className="text-slate-400"> / {fmtMoneyShort(plan)}</span>}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="flex-1 bg-slate-100 rounded-full h-2 overflow-hidden">
                    <div className={`h-full ${colorBarra}`} style={{ width: `${pctBarra}%` }} />
                  </div>
                  <span className={`text-xs font-medium w-12 text-right ${pct === null ? 'text-slate-400' : pct > 1 ? 'text-red-600' : pct > 0.9 ? 'text-amber-600' : 'text-green-600'}`}>
                    {pct === null ? 's/plan' : fmtPct(pct)}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
        {categoriasConDatos.length > 0 && (
          <div className="flex items-center justify-between mt-3 pt-3 border-t border-slate-200 text-sm font-bold">
            <span className="text-slate-900">Total</span>
            <span className="text-slate-900">
              {fmtMoney(totalEgresos)}
              {totalPlanificado > 0 && <span className="text-slate-400 font-medium"> / {fmtMoney(totalPlanificado)}</span>}
            </span>
          </div>
        )}
      </div>

      {/* Medios de pago */}
      <div className="bg-white rounded-xl border border-slate-200 p-3">
        <h3 className="text-sm font-semibold text-slate-900 mb-2">Egresos por medio de pago</h3>
        <div className="grid grid-cols-2 gap-2">
          {MEDIOS_PAGO.map((mp) => {
            const total = movsDelMes.filter((m) => m.tipo === 'egreso' && m.medioPago === mp).reduce((s, m) => s + montoEnARS(m, config), 0);
            const pct = totalEgresos > 0 ? total / totalEgresos : 0;
            if (total === 0) return null;
            return (
              <div key={mp} className="bg-slate-50 rounded-lg p-2.5">
                <div className="text-xs text-slate-500">{mp}</div>
                <div className="text-sm font-semibold text-slate-900">{fmtMoney(total)}</div>
                <div className="text-xs text-slate-400">{fmtPct(pct)}</div>
              </div>
            );
          })}
          {totalEgresos === 0 && <p className="text-sm text-slate-400 col-span-2 text-center py-2">Sin egresos este mes</p>}
        </div>
      </div>
    </div>
  );
}

function KPI({ label, valor, color = 'slate', esPct = false }) {
  const colorMap = {
    green: 'text-green-700 bg-green-50',
    red: 'text-red-700 bg-red-50',
    slate: 'text-slate-600 bg-slate-100',
  };
  const display = esPct ? (valor === null ? 'sin plan' : fmtPct(valor)) : fmtMoney(valor);
  return (
    <div className={`rounded-xl p-3 ${colorMap[color]}`}>
      <div className="text-xs opacity-70">{label}</div>
      <div className="text-lg font-bold mt-0.5 leading-tight">{display}</div>
    </div>
  );
}

// ============ NUEVO MOVIMIENTO ============
function NuevoMovimiento({ agregar, config, onDone }) {
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

  useEffect(() => {
    const cats = Object.keys(categorias);
    if (!cats.includes(categoria)) setCategoria(cats[0]);
  }, [tipo]);

  useEffect(() => {
    if (subcategoriasDisponibles.length > 0 && !subcategoriasDisponibles.includes(subcategoria)) {
      setSubcategoria(subcategoriasDisponibles[0]);
    }
  }, [categoria]);

  const handleSubmit = () => {
    if (!monto || Number(monto) <= 0) { setMensaje('error:Ingresá un monto válido'); return; }

    if (tipo === 'transferencia') {
      const cuentaOrig = CUENTAS_DEFAULT.find((c) => c.id === cuentaOrigen);
      const cuentaDest = CUENTAS_DEFAULT.find((c) => c.id === cuentaDestino);
      agregar({
        tipo: 'transferencia', fecha, monto: Number(monto), moneda: cuentaOrig.moneda, cotizacion: Number(cotizacion),
        cuentaOrigen, cuentaDestino, montoDestino: Number(montoDestino) || Number(monto), monedaDestino: cuentaDest.moneda,
        descripcion: descripcion || `${cuentaOrig.nombre} → ${cuentaDest.nombre}`, categoria: 'Transferencia',
      });
    } else {
      const mov = { tipo, fecha, monto: Number(monto), moneda, cotizacion: moneda === 'USD' ? Number(cotizacion) : null, categoria, subcategoria, descripcion };
      if (tipo === 'egreso') {
        mov.medioPago = medioPago;
        if (esCuotas && cuotas > 1) mov.cuotas = Number(cuotas);
      }
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
      <h3 className="text-base font-semibold text-slate-900 px-1">Nuevo movimiento</h3>

      {/* Selector tipo */}
      <div className="grid grid-cols-3 gap-2">
        {tipoBtns.map((t) => {
          const activo = tipo === t.id;
          const cls = activo
            ? t.c === 'red' ? 'border-red-500 bg-red-50 text-red-700'
              : t.c === 'green' ? 'border-green-500 bg-green-50 text-green-700'
                : 'border-blue-500 bg-blue-50 text-blue-700'
            : 'border-slate-200 bg-white text-slate-500';
          return (
            <button key={t.id} onClick={() => setTipo(t.id)} className={`flex flex-col items-center gap-1 py-3 rounded-xl border-2 transition ${cls}`}>
              <t.icon size={20} />
              <span className="text-xs font-medium">{t.label}</span>
            </button>
          );
        })}
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-3">
        {(tipo === 'egreso' || tipo === 'ingreso') && (
          <>
            {/* Monto grande primero — lo más importante */}
            <Field label="Monto">
              <div className="flex gap-2">
                <input type="number" inputMode="decimal" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" className="flex-1 px-3 py-3 border border-slate-300 rounded-lg text-2xl font-bold text-slate-900" />
                <select value={moneda} onChange={(e) => setMoneda(e.target.value)} className="px-3 py-3 border border-slate-300 rounded-lg text-base font-medium">
                  <option value="ARS">ARS</option>
                  <option value="USD">USD</option>
                </select>
              </div>
            </Field>

            {moneda === 'USD' && (
              <Field label="Cotización USD">
                <input type="number" inputMode="numeric" value={cotizacion} onChange={(e) => setCotizacion(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
              </Field>
            )}

            <Field label="Fecha">
              <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
            </Field>

            <Field label="Categoría">
              <select value={categoria} onChange={(e) => setCategoria(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg">
                {Object.keys(categorias).map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>

            {subcategoriasDisponibles.length > 0 && (
              <Field label="Subcategoría">
                <select value={subcategoria} onChange={(e) => setSubcategoria(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg">
                  {subcategoriasDisponibles.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </Field>
            )}

            {tipo === 'egreso' && (
              <>
                <Field label="Medio de pago">
                  <div className="grid grid-cols-3 gap-1.5">
                    {MEDIOS_PAGO.map((m) => (
                      <button key={m} onClick={() => setMedioPago(m)} className={`py-2 rounded-lg text-xs font-medium border ${medioPago === m ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-slate-200 bg-white text-slate-600'}`}>
                        {m}
                      </button>
                    ))}
                  </div>
                </Field>

                <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={esCuotas} onChange={(e) => setEsCuotas(e.target.checked)} className="w-4 h-4 rounded" />
                    <span className="text-sm font-medium text-slate-700">¿Es compra en cuotas?</span>
                  </label>
                  {esCuotas && (
                    <div className="mt-2 flex items-center gap-2 flex-wrap">
                      <span className="text-sm text-slate-600">Cuotas:</span>
                      <input type="number" inputMode="numeric" value={cuotas} onChange={(e) => setCuotas(e.target.value)} min="2" max="60" className="w-16 px-2 py-1.5 border border-slate-300 rounded" />
                      {monto > 0 && cuotas > 1 && <span className="text-sm text-slate-500">de {fmtMoney(Number(monto) / Number(cuotas), moneda)} c/u</span>}
                    </div>
                  )}
                </div>
              </>
            )}

            <Field label="Descripción (opcional)">
              <input type="text" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder="Detalle" className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
            </Field>
          </>
        )}

        {tipo === 'transferencia' && (
          <>
            <Field label="Fecha">
              <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
            </Field>
            <Field label="Cuenta origen">
              <select value={cuentaOrigen} onChange={(e) => setCuentaOrigen(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg">
                {CUENTAS_DEFAULT.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
              </select>
            </Field>
            <Field label={`Monto origen (${CUENTAS_DEFAULT.find((c) => c.id === cuentaOrigen)?.moneda})`}>
              <input type="number" inputMode="decimal" value={monto} onChange={(e) => setMonto(e.target.value)} placeholder="0" className="w-full px-3 py-2.5 border border-slate-300 rounded-lg text-lg font-semibold" />
            </Field>
            <Field label="Cuenta destino">
              <select value={cuentaDestino} onChange={(e) => setCuentaDestino(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg">
                {CUENTAS_DEFAULT.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
              </select>
            </Field>
            <Field label={`Monto destino (${CUENTAS_DEFAULT.find((c) => c.id === cuentaDestino)?.moneda})`}>
              <input type="number" inputMode="decimal" value={montoDestino} onChange={(e) => setMontoDestino(e.target.value)} placeholder="Si es misma moneda, dejá vacío" className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
            </Field>
            {monto && montoDestino && CUENTAS_DEFAULT.find((c) => c.id === cuentaOrigen)?.moneda !== CUENTAS_DEFAULT.find((c) => c.id === cuentaDestino)?.moneda && (
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-2.5 text-sm text-blue-900">
                Tipo de cambio: <strong>{(Number(monto) / Number(montoDestino)).toFixed(2)}</strong>
              </div>
            )}
            <Field label="Descripción">
              <input type="text" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} placeholder="Ej: Compra dólares" className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
            </Field>
          </>
        )}
      </div>

      <button onClick={handleSubmit} className="w-full bg-blue-600 active:bg-blue-700 text-white font-semibold py-3.5 rounded-xl flex items-center justify-center gap-2 text-base">
        <Plus size={20} /> Registrar
      </button>

      {mensaje && (
        <div className={`text-center text-sm font-medium py-2 rounded-lg ${mensaje.startsWith('ok') ? 'text-green-700 bg-green-50' : 'text-red-700 bg-red-50'}`}>
          {mensaje.split(':')[1]}
        </div>
      )}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="text-xs font-medium text-slate-600 block mb-1">{label}</span>
      {children}
    </label>
  );
}

// ============ MOVIMIENTOS ============
function Movimientos({ movimientos, eliminar, actualizar, config }) {
  const [filtroTipo, setFiltroTipo] = useState('todos');
  const [filtroMes, setFiltroMes] = useState(mesAnio(hoy()));
  const [editando, setEditando] = useState(null);

  const filtrados = useMemo(() => {
    return movimientos
      .filter((m) => filtroTipo === 'todos' || m.tipo === filtroTipo)
      .filter((m) => filtroMes === 'todos' || mesAnio(m.fecha) === filtroMes)
      .sort((a, b) => b.fecha.localeCompare(a.fecha));
  }, [movimientos, filtroTipo, filtroMes]);

  const mesesDisponibles = useMemo(() => {
    const set = new Set(movimientos.map((m) => mesAnio(m.fecha)));
    return Array.from(set).sort().reverse();
  }, [movimientos]);

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between px-1">
        <h3 className="text-base font-semibold text-slate-900">Movimientos ({filtrados.length})</h3>
      </div>

      <div className="grid grid-cols-2 gap-2">
        <select value={filtroTipo} onChange={(e) => setFiltroTipo(e.target.value)} className="px-3 py-2 border border-slate-300 rounded-lg text-sm bg-white">
          <option value="todos">Todos los tipos</option>
          <option value="ingreso">Ingresos</option>
          <option value="egreso">Egresos</option>
          <option value="transferencia">Transferencias</option>
        </select>
        <select value={filtroMes} onChange={(e) => setFiltroMes(e.target.value)} className="px-3 py-2 border border-slate-300 rounded-lg text-sm bg-white">
          <option value="todos">Todos los meses</option>
          {mesesDisponibles.map((m) => <option key={m} value={m}>{m}</option>)}
        </select>
      </div>

      <div className="space-y-2">
        {filtrados.length === 0 && <div className="text-center py-8 text-slate-400 text-sm">No hay movimientos</div>}
        {filtrados.map((m) => (
          <MovimientoCard key={m.id} mov={m} eliminar={eliminar} config={config} onEdit={() => setEditando(m)} />
        ))}
      </div>

      {editando && (
        <EditarModal mov={editando} onClose={() => setEditando(null)} onSave={(cambios) => { actualizar(editando.id, cambios); setEditando(null); }} />
      )}
    </div>
  );
}

function MovimientoCard({ mov, eliminar, config, onEdit }) {
  const tipoColor = mov.tipo === 'ingreso' ? 'text-green-600' : mov.tipo === 'egreso' ? 'text-red-600' : 'text-blue-600';
  const tipoIcon = mov.tipo === 'ingreso' ? '+' : mov.tipo === 'egreso' ? '−' : '↔';

  return (
    <div className="flex items-center gap-2 p-3 bg-white border border-slate-200 rounded-xl">
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className={`text-base font-bold ${tipoColor}`}>{tipoIcon}</span>
          <span className="font-semibold text-slate-900 text-sm">{fmtMoney(mov.monto, mov.moneda)}</span>
          {mov.moneda === 'USD' && <span className="text-xs text-slate-400">≈{fmtMoney(montoEnARS(mov, config))}</span>}
          <span className="text-[11px] px-1.5 py-0.5 bg-slate-100 rounded-full text-slate-600">{mov.categoria}</span>
        </div>
        <div className="text-[11px] text-slate-400 mt-0.5 flex items-center gap-1.5 flex-wrap">
          <span>{mov.fecha}</span>
          {mov.subcategoria && <span>• {mov.subcategoria}</span>}
          {mov.medioPago && <span>• {mov.medioPago}</span>}
          {mov.descripcion && <span>• {mov.descripcion}</span>}
          {mov.totalCuotas && <span className="text-amber-600">• {mov.numeroCuota}/{mov.totalCuotas}</span>}
        </div>
      </div>
      {!mov.esCuota && (
        <button onClick={onEdit} className="p-2 text-slate-400 active:text-blue-600"><Edit2 size={15} /></button>
      )}
      <button onClick={() => eliminar(mov.id)} className="p-2 text-slate-400 active:text-red-600"><Trash2 size={15} /></button>
    </div>
  );
}

function EditarModal({ mov, onClose, onSave }) {
  const [monto, setMonto] = useState(mov.monto);
  const [descripcion, setDescripcion] = useState(mov.descripcion || '');
  const [fecha, setFecha] = useState(mov.fecha);

  return (
    <div className="fixed inset-0 bg-black/40 flex items-end sm:items-center justify-center z-50 p-3" onClick={onClose}>
      <div className="bg-white rounded-2xl w-full max-w-md p-4 space-y-3" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between">
          <h3 className="text-base font-semibold text-slate-900">Editar movimiento</h3>
          <button onClick={onClose} className="p-1 text-slate-400"><X size={20} /></button>
        </div>
        <Field label="Monto">
          <input type="number" inputMode="decimal" value={monto} onChange={(e) => setMonto(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg text-lg font-semibold" />
        </Field>
        <Field label="Fecha">
          <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
        </Field>
        <Field label="Descripción">
          <input type="text" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} className="w-full px-3 py-2.5 border border-slate-300 rounded-lg" />
        </Field>
        <button onClick={() => onSave({ monto: Number(monto), descripcion, fecha })} className="w-full bg-blue-600 active:bg-blue-700 text-white font-semibold py-3 rounded-xl flex items-center justify-center gap-2">
          <Check size={18} /> Guardar
        </button>
      </div>
    </div>
  );
}

// ============ CUOTAS ============
function Cuotas({ movimientos }) {
  const cuotasGrupos = useMemo(() => {
    const grupos = {};
    movimientos.filter((m) => m.grupoCuota).forEach((m) => {
      if (!grupos[m.grupoCuota]) {
        grupos[m.grupoCuota] = {
          descripcion: m.descripcion?.split('(')[0]?.trim() || m.categoria,
          categoria: m.categoria, medioPago: m.medioPago, montoTotal: 0, montoCuota: m.monto,
          totalCuotas: m.totalCuotas, fechaInicio: m.fecha, fechaFin: '', cuotasPagadas: 0, cuotasPendientes: 0,
        };
      }
      grupos[m.grupoCuota].montoTotal += m.monto;
      if (m.fecha <= hoy()) grupos[m.grupoCuota].cuotasPagadas += 1;
      else grupos[m.grupoCuota].cuotasPendientes += 1;
      if (m.fecha > grupos[m.grupoCuota].fechaFin) grupos[m.grupoCuota].fechaFin = m.fecha;
    });
    return Object.values(grupos).sort((a, b) => b.fechaInicio.localeCompare(a.fechaInicio));
  }, [movimientos]);

  const pendientesTotal = cuotasGrupos.reduce((s, g) => s + g.cuotasPendientes * g.montoCuota, 0);

  return (
    <div className="space-y-3">
      <h3 className="text-base font-semibold text-slate-900 px-1">Cuotas activas</h3>

      <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-center">
        <div className="text-xs text-amber-700">Pendiente total en cuotas</div>
        <div className="text-xl font-bold text-amber-900">{fmtMoney(pendientesTotal)}</div>
      </div>

      {cuotasGrupos.length === 0 && <div className="text-center py-8 text-slate-400 text-sm">No hay compras en cuotas</div>}

      <div className="space-y-2">
        {cuotasGrupos.map((g, i) => (
          <div key={i} className="bg-white border border-slate-200 rounded-xl p-3">
            <div className="flex items-start justify-between gap-2">
              <div className="min-w-0">
                <div className="font-semibold text-slate-900 text-sm truncate">{g.descripcion}</div>
                <div className="text-xs text-slate-400">{g.categoria} • {g.medioPago}</div>
              </div>
              <div className="text-right shrink-0">
                <div className="font-semibold text-sm text-slate-900">{fmtMoney(g.montoCuota)} × {g.totalCuotas}</div>
                <div className="text-xs text-slate-400">Total: {fmtMoney(g.montoTotal)}</div>
              </div>
            </div>
            <div className="mt-2 flex items-center gap-2">
              <div className="flex-1 bg-slate-100 rounded-full h-2 overflow-hidden">
                <div className="bg-blue-500 h-full" style={{ width: `${(g.cuotasPagadas / g.totalCuotas) * 100}%` }} />
              </div>
              <span className="text-xs text-slate-500">{g.cuotasPagadas}/{g.totalCuotas}</span>
            </div>
            <div className="text-[11px] text-slate-400 mt-1">{g.fechaInicio} → {g.fechaFin}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ============ MÁS (Presupuesto + Exportar) ============
function Mas({ presupuesto, actualizar, movimientos, config }) {
  const [vista, setVista] = useState('menu');

  if (vista === 'presupuesto') return <PresupuestoMobile presupuesto={presupuesto} actualizar={actualizar} onBack={() => setVista('menu')} />;
  if (vista === 'exportar') return <Exportar movimientos={movimientos} presupuesto={presupuesto} config={config} onBack={() => setVista('menu')} />;

  return (
    <div className="space-y-3">
      <h3 className="text-base font-semibold text-slate-900 px-1">Más opciones</h3>
      <button onClick={() => setVista('presupuesto')} className="w-full bg-white border border-slate-200 rounded-xl p-4 flex items-center gap-3 active:bg-slate-50">
        <div className="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center"><Target size={20} className="text-blue-600" /></div>
        <div className="text-left flex-1">
          <div className="font-medium text-slate-900 text-sm">Presupuesto mensual</div>
          <div className="text-xs text-slate-400">Planificá cuánto querés gastar por categoría</div>
        </div>
        <ChevronRight size={18} className="text-slate-400" />
      </button>
      <button onClick={() => setVista('exportar')} className="w-full bg-white border border-slate-200 rounded-xl p-4 flex items-center gap-3 active:bg-slate-50">
        <div className="w-10 h-10 rounded-lg bg-green-50 flex items-center justify-center"><FileSpreadsheet size={20} className="text-green-600" /></div>
        <div className="text-left flex-1">
          <div className="font-medium text-slate-900 text-sm">Exportar a Excel</div>
          <div className="text-xs text-slate-400">Descargá todo para analizar en la compu</div>
        </div>
        <ChevronRight size={18} className="text-slate-400" />
      </button>
    </div>
  );
}

// Presupuesto mobile: mes por mes en vez de grilla 12 meses
function PresupuestoMobile({ presupuesto, actualizar, onBack }) {
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
        <button onClick={onBack} className="p-1.5 -ml-1.5 text-slate-500"><ChevronLeft size={22} /></button>
        <h3 className="text-base font-semibold text-slate-900">Presupuesto</h3>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 px-3 py-2.5 flex items-center justify-between">
        <button onClick={() => cambiarMes(-1)} className="p-2 rounded-lg active:bg-slate-100"><ChevronLeft size={20} className="text-slate-600" /></button>
        <h2 className="text-base font-bold text-slate-900">{MESES[mNum - 1]} {y}</h2>
        <button onClick={() => cambiarMes(1)} className="p-2 rounded-lg active:bg-slate-100"><ChevronRight size={20} className="text-slate-600" /></button>
      </div>

      <button onClick={copiarMesAnterior} className="w-full bg-blue-50 border border-blue-200 text-blue-700 rounded-lg py-2 text-sm font-medium active:bg-blue-100">
        Copiar del mes anterior
      </button>
      {copiarMsg && <div className="text-center text-xs text-slate-600">{copiarMsg}</div>}

      <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-2">
        {Object.keys(CATEGORIAS_EGRESO).map((cat) => (
          <div key={cat} className="flex items-center justify-between gap-2">
            <span className="text-sm text-slate-700 flex-1">{cat}</span>
            <input
              type="number"
              inputMode="numeric"
              value={datosMes[cat] || ''}
              onChange={(e) => actualizar(mesSel, cat, Number(e.target.value) || 0)}
              placeholder="0"
              className="w-28 px-2 py-1.5 text-right border border-slate-200 rounded-lg text-sm"
            />
          </div>
        ))}
        <div className="flex items-center justify-between pt-2 mt-2 border-t border-slate-200 font-bold">
          <span className="text-sm text-slate-900">Total mes</span>
          <span className="text-sm text-slate-900">{fmtMoney(totalMes)}</span>
        </div>
      </div>
    </div>
  );
}

// ============ EXPORTAR ============
function Exportar({ movimientos, presupuesto, config, onBack }) {
  const exportarExcel = () => {
    const wb = XLSX.utils.book_new();

    const movsData = movimientos.map((m) => ({
      Fecha: m.fecha, Tipo: m.tipo, Categoría: m.categoria, Subcategoría: m.subcategoria || '', Descripción: m.descripcion || '',
      Monto: m.monto, Moneda: m.moneda, Cotización: m.cotizacion || '', 'Monto ARS': montoEnARS(m, config),
      'Medio Pago': m.medioPago || '', 'Cuenta Origen': m.cuentaOrigen || '', 'Cuenta Destino': m.cuentaDestino || '',
      'Es cuota': m.esCuota ? 'Sí' : 'No', 'N° Cuota': m.numeroCuota || '', 'Total Cuotas': m.totalCuotas || '',
    }));
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
    categorias.forEach((cat) => {
      const row = [cat];
      let anual = 0;
      meses.forEach((m) => { const v = (presupuesto[m] || {})[cat] || 0; row.push(v); anual += v; });
      row.push(anual);
      presData.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(presData), 'Presupuesto');

    const resumen = {};
    movimientos.filter((m) => m.tipo === 'egreso').forEach((m) => {
      const key = `${m.categoria} - ${m.subcategoria || 'General'}`;
      resumen[key] = (resumen[key] || 0) + montoEnARS(m, config);
    });
    const resumenData = [['Categoría/Subcategoría', 'Total ARS']];
    Object.entries(resumen).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => resumenData.push([k, v]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), 'Resumen Cat');

    XLSX.writeFile(wb, `Finanzas_Agustin_${hoy()}.xlsx`);
  };

  return (
    <div className="space-y-3">
      <div className="flex items-center gap-2 px-1">
        <button onClick={onBack} className="p-1.5 -ml-1.5 text-slate-500"><ChevronLeft size={22} /></button>
        <h3 className="text-base font-semibold text-slate-900">Exportar a Excel</h3>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-3">
        <p className="text-sm text-slate-600 mb-2">El Excel incluye 4 hojas:</p>
        <ul className="text-sm text-slate-700 space-y-1.5">
          <li>• <strong>Movimientos</strong>: todos los registros con detalle</li>
          <li>• <strong>Seguimiento</strong>: Plan vs Real por mes y categoría</li>
          <li>• <strong>Presupuesto</strong>: plan anual por categoría</li>
          <li>• <strong>Resumen Cat</strong>: total por subcategoría</li>
        </ul>
      </div>

      <button onClick={exportarExcel} className="w-full bg-green-600 active:bg-green-700 text-white font-semibold py-3.5 rounded-xl flex items-center justify-center gap-2">
        <Download size={20} /> Descargar Excel
      </button>
      <div className="text-center text-xs text-slate-400">{movimientos.length} movimientos guardados</div>
    </div>
  );
}
