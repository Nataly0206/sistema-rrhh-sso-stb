<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<title>Identificación de Riesgo por Puesto</title>
<style>
  :root{
    --border:#000;
    --gris:#bfbfbf;
    --font: "Arial", Helvetica, sans-serif;
  }
  *{ box-sizing:border-box; }
  body{ margin:0; font-family:var(--font); color:#000; }
  .a4{ width:210mm; margin:0 auto; padding:10mm 12mm; }

  /* Encabezado */
  .head{ display:flex; align-items:flex-start; gap:12px; }
  .head img{ height:48px; width:auto; object-fit:contain; margin-top:2px; }
  .head-text{ text-align:center; flex:1; line-height:1.2; font-size:12px; text-transform:uppercase; }
  .head-text .line1{ font-weight:700; }
  .head-hr{ border:none; border-top:1px solid var(--border); margin:6px 0 0; }

  /* Barras de sección */
  .bar{ background:var(--gris); color:#000; text-transform:uppercase; 
        font-weight:700; text-align:center; padding:6px; border:1px solid var(--border); }

  /* Tabla principal */
  table.sheet{ width:100%; border-collapse:collapse; table-layout:fixed; }
  .sheet td, .sheet th{ border:1px solid var(--border); vertical-align:top; }
  .label{ width:36%; padding:6px; font-size:12px; }
  .field{ width:64%; padding:0; }

  /* Campos (estilo “celda vacía”) */
  .cell-input, .cell-textarea{
    width:100%; border:0; padding:8px; font:inherit; resize:none; 
    outline:none; background:transparent;
  }
  .cell-textarea{ height:90px; } /* alto para “Descripción General de la labor” */
  .objetivo{ height:110px; }

  /* Imprime bonito */
  @media print{
    body{ margin:0; }
    .a4{ width:auto; padding:0; }
  }
  /* --- Ajustes globales (sin cambiar contenido) --- */
  /*.sheet col{ width:auto !important; }                 /* anula colgroups desparejos */
  .sheet{ table-layout:fixed; }                        /* grilla estable para imprimir */
  .sheet th, .sheet td{ padding:6px 8px; line-height:1.25; }
  .sheet th{ font-size:11px; }
  .sheet th, .sheet td{ white-space:normal; overflow-wrap:anywhere; word-break:normal; hyphens:auto; }
  .cell-input, .cell-textarea{ padding:6px 8px; }

  /* 7) Firmas/fechas: evita cortes raros en las etiquetas */
  .sheet .label-cell{ white-space:normal; }

  /* Poner TODAS las celdas TIPO en vertical */
.risk-table td.tipo{
  writing-mode: vertical-rl;   /* texto de arriba hacia abajo */
  text-orientation: mixed;     /* mantiene letras latinas derechas */
  white-space: nowrap;         /* evita cortes o saltos raros */
  text-align: center;          /* centra horizontalmente */
  vertical-align: middle;      /* centra verticalmente (sobre el rowspan) */
  padding: 8px 4px;
}

/* Oculta los <br> dentro de la celda tipo para que sea una sola línea vertical */
.risk-table td.tipo br{ display:none; }

/* (Opcional) si quieres estrechar la primera columna:
.risk-table td.tipo{ width:32px !important; }
*/

  /* Vertical abajo → arriba (letras derechas) */
  .risk-table td.tipo{
    writing-mode: vertical-rl;   /* columna vertical */
    text-orientation: mixed;     /* letras latinas derechas */
    transform: rotate(180deg);   /* invierte el sentido a abajo→arriba */
    white-space: nowrap;
    text-align: center;
    vertical-align: middle;
    padding: 8px 4px;
  }
  /* Re-endereza el texto dentro para que no quede cabeza abajo */
  .risk-table td.tipo .tipo-inner{
    display: inline-block;
    transform: rotate(180deg);
  }
  /* Opcional: ocultar <br> dentro de TIPO para que sea una sola línea vertical */
  .risk-table td.tipo br{ display:none; }

  /* Opcional: ajustar ancho de la columna TIPO si quieres más angosta */
  /* .risk-table td.tipo{ width:32px !important; } */

  .puesto-layout{ display:flex; gap:16px; align-items:stretch; }
  .puesto-main{ flex:1 1 auto; min-width:0; display:flex; flex-direction:column; }
  .puesto-main.sheet{ flex:1 1 auto; }
  .puestos-sidebar{ width:240px; border:1px solid var(--border); border-radius:6px; padding:12px; background:#f6f6f6; font-size:12px; line-height:1.3; display:flex; flex-direction:column; height:490px; }
  .puestos-sidebar__header{ display:flex; flex-direction:column; gap:6px; margin-bottom:8px; }
  .puestos-sidebar__title{ font-weight:700; text-transform:uppercase; letter-spacing:0.5px; }
  .puestos-sidebar__legend{ display:flex; gap:10px; flex-wrap:wrap; font-size:11px; align-items:center; }
  .puestos-sidebar__legend span{ display:flex; align-items:center; gap:4px; }
  .puestos-sidebar__search{ margin-top:4px; }
  .puestos-sidebar__input{ width:100%; padding:6px 8px; border:1px solid var(--border); border-radius:4px; font:inherit; background:#fff; }
  .puestos-sidebar__scroll{ flex:1 1 auto; overflow:auto; border:1px solid var(--border); border-radius:4px; background:#fff; }
  .puestos-sidebar__table{ width:100%; border-collapse:collapse; font-size:11px; }
  .puestos-sidebar__table td{ padding:6px 8px; border-bottom:1px solid rgba(0,0,0,.08); }
  .puestos-sidebar__name{ width:60%; }
  .puestos-sidebar__status{ width:40%; text-align:right; white-space:nowrap; }
  .puestos-sidebar__row--with .puestos-sidebar__status{ color:#1f7a35; font-weight:600; }
  .puestos-sidebar__row--pending .puestos-sidebar__status{ color:#b3261e; font-weight:600; }
  .puestos-sidebar__row--active{ background:#e8f3ff; }
  .puestos-sidebar__row--active td{ font-weight:600; }
  .puestos-sidebar__row--empty td,
  .puestos-sidebar__row--no-results td{ text-align:center; color:#666; font-style:italic; }
  .puestos-sidebar__row:last-child td{ border-bottom:none; }
  .status-dot{ display:inline-block; width:10px; height:10px; border-radius:50%; background:#9e9e9e; }
  .status-dot--con{ background:#1f7a35; }
  .status-dot--sin{ background:#b3261e; }

  .puesto-select{ display:flex; gap:8px; align-items:center; }
  .puesto-select__input{ flex:1 1 auto; min-width:0; }
  .puesto-select__btn{ flex:0 0 auto; padding:8px 12px; border:1px solid var(--border); border-radius:6px; background:#e7f6ff; font-size:12px; font-weight:600; cursor:pointer; color:#0f172a; transition:background-color .2s ease, box-shadow .2s ease; }
  .puesto-select__btn:hover{ background:#d4ecff; box-shadow:0 1px 0 rgba(15,23,42,.18); }
  .puesto-select__btn:focus{ outline:2px solid #44B3E1; outline-offset:2px; }

  .modal{ position:fixed; inset:0; z-index:1100; display:none; align-items:center; justify-content:center; }
  .modal.is-open{ display:flex; }
  .modal__backdrop{ position:absolute; inset:0; background:rgba(15,23,42,.45); }
  .modal__dialog{ position:relative; width:min(720px, 92%); max-height:85vh; background:#fff; border-radius:14px; box-shadow:0 24px 48px rgba(15,23,42,.25); display:flex; flex-direction:column; overflow:hidden; }
  .modal__header{ padding:16px 20px; border-bottom:1px solid rgba(0,0,0,.08); display:flex; align-items:center; justify-content:space-between; gap:12px; }
  .modal__title{ font-size:16px; font-weight:700; margin:0; }
  .modal__close{ background:none; border:0; font-size:24px; line-height:1; cursor:pointer; color:#64748b; padding:4px; border-radius:6px; }
  .modal__close:hover{ background:rgba(148,163,184,.15); color:#0f172a; }
  .modal__body{ padding:16px 20px; display:flex; flex-direction:column; gap:12px; }
  .modal__search input{ width:100%; padding:10px 12px; border:1px solid var(--border); border-radius:8px; font:inherit; }
  .modal__list{ max-height:50vh; overflow:auto; border:1px solid var(--border); border-radius:10px; background:#fff; }
  .modal__table{ width:100%; border-collapse:collapse; font-size:13px; }
  .modal__table tbody tr{ cursor:pointer; transition:background-color .15s ease; }
  .modal__table tbody tr:hover{ background:#f1f5f9; }
  .modal__table td{ padding:10px 14px; border-bottom:1px solid rgba(15,23,42,.08); }
  .modal__status{ text-align:right; white-space:nowrap; font-size:12px; color:#475569; }
  .modal__row--ready .modal__status{ color:#1f7a35; }
  .modal__row--pending .modal__status{ color:#b45309; }
  .modal__row--active{ background:#e0f2fe !important; }
  .modal__footer{ padding:12px 20px; border-top:1px solid rgba(0,0,0,.08); display:flex; justify-content:flex-end; gap:8px; }
  .modal__btn{ appearance:none; border-radius:8px; border:1px solid var(--border); padding:8px 14px; font-weight:600; cursor:pointer; background:#fff; color:#0f172a; transition:background-color .2s ease; }
  .modal__btn:hover{ background:#f1f5f9; }
  .modal__btn--primary{ background:#44B3E1; border-color:#44B3E1; color:#fff; }
  .modal__btn--primary:hover{ background:#1996cd; }
  .modal__empty{ padding:16px; text-align:center; color:#64748b; display:none; }
  body.modal-open{ overflow:hidden; }

  .actions{ margin-top:12px; }
.btn-guardar{
  display:block;
  width:100%;
  padding:12px 16px;
  font:700 16px var(--font);
  text-transform:uppercase;
  border:1px solid var(--border);
  border-radius:8px;
  background: linear-gradient(180deg, #2ea043, #1f7a35); /* verde agradable */
  color:;
  box-shadow: 0 2px 6px rgba(0,0,0,.15);
  cursor:pointer;
}
.btn-guardar:hover{ filter:brightness(1.05); }
.btn-guardar:active{ transform:translateY(1px); box-shadow: 0 1px 4px rgba(0,0,0,.2); }
.btn-guardar:disabled{
  background:#9ccaa8; color:#fff; cursor:not-allowed; box-shadow:none;
}
/* No mostrar al imprimir */
@media print{ .actions{ display:none; } .puestos-sidebar{ display:none; } .puesto-layout{ display:block; } }
</style>

<script>
  // Envuelve el contenido actual de cada TIPO en un span (una sola vez)
  document.querySelectorAll('.risk-table td.tipo').forEach(td => {
    if (!td.querySelector('.tipo-inner')) {
      td.innerHTML = '<span class="tipo-inner">' + td.innerHTML + '</span>';
    }
  });
</script>

</head>
<body>

<a href="/verificacion" class="btn-volver" type="button" aria-label="Volver a inicio">
  <span class="ico" aria-hidden="true"></span>
  <span>Volver a inicio</span>
</a>

<style>
  .btn-volver{
    --brand1:#00B0F0; --brand2:#0088BC;
    display:inline-flex; align-items:center; gap:.55rem;
    padding:.55rem 1rem; border:0; cursor:pointer;
    border-radius:9999px; color:#fff; font-weight:600;
    background:linear-gradient(90deg,var(--brand1),var(--brand2));
    box-shadow:0 6px 16px rgba(0,0,0,.12);
    transition:transform .15s ease, box-shadow .15s ease, filter .15s ease;
    margin-top: 20px; 
  }
  .btn-volver:hover{ transform:translateX(-2px); box-shadow:0 10px 22px rgba(0,0,0,.16); filter:saturate(1.05); }
  .btn-volver:active{ transform:translateX(-1px) scale(.98); }
  .btn-volver:focus-visible{ outline:3px solid #00B0F0; outline-offset:2px; }

  /* Flecha izquierda (puro CSS) */
  .btn-volver .ico{
    position:relative; width:16px; height:14px; margin-left:-2px;
    transition:transform .15s ease;
  }
  .btn-volver .ico::before{ /* eje */
    content:""; position:absolute; left:2px; top:50%;
    width:14px; height:2px; background:currentColor; border-radius:2px;
    transform:translateY(-50%);
  }
  .btn-volver .ico::after{ /* punta */
    content:""; position:absolute; left:0; top:50%;
    width:8px; height:8px; border-left:2px solid currentColor; border-bottom:2px solid currentColor;
    transform:translateY(-50%) rotate(45deg);
  }
  .btn-volver:hover .ico{ transform:translateX(-3px); }

  /* Variante “ghost” clara (opcional) */
  .btn-volver.ghost{
    background:#fff; color:#0088BC; border:1px solid #00B0F0;
    box-shadow:none;
  }
  .btn-volver.ghost:hover{ background:#E6F6FF; }
  .btn-volver:disabled{ opacity:.6; cursor:not-allowed; }
</style>

@if ($errors->any())
  <div style="margin:10px 0; padding:8px; border:1px solid #c00; color:#b00020;">
    @foreach ($errors->all() as $e)
      <div>• {{ $e }}</div>
    @endforeach
  </div>
@endif

@if (session('success'))
  <div style="margin:10px 0; padding:8px; border:1px solid #0a0; color:#0a0;">
    {{ session('success') }}
  </div>
@endif

  <form id="riesgos-form"
      action="{{ route('identificacion.store') }}"
      method="POST"
      data-fetch-base="{{ url('/identificacion') }}">
  @csrf
  <div class="a4">
    <!-- ENCABEZADO -->
    <div class="head">
      <img src="{{ asset('img/logo.PNG') }}" alt="Logo">
      <div class="head-text">
        <div class="line1">Service and Trading Business S.A. de C.V.</div>
        <div>Proceso Salud y Seguridad Ocupacional / Health and Occupational Safety Process</div>
        <div>Identificación de riesgo por puesto de trabajo / Identification of risk by work position</div>
      </div>
    </div>
    <hr class="head-hr">

    <div class="puesto-layout">
      <div class="puesto-main">
          <div class="bar">Datos Generales del Puesto</div>
      <table class="sheet">
        <colgroup><col style="width:36%"><col style="width:64%"></colgroup>
        <tr>
          <td class="label">Puesto de Trabajo Analizado</td>
          <td class="field">
            <input type="hidden" name="id_puesto_trabajo_matriz" id="puesto_id" value="{{ $id_puesto_trabajo_matriz ?? '' }}">
            <input type="text" id="puesto-input" list="puestos-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" value="{{ old('puesto_nombre') }}">
            <datalist id="puestos-list">
              @foreach ($puestos as $p)
                <option data-id="{{ $p->id_puesto_trabajo_matriz }}" value="{{ $p->puesto_trabajo_matriz }}"></option>
              @endforeach
            </datalist>
          </td>
        </tr>

        <tr>
          <td class="label">Departamento</td>
          <td class="field"><input class="cell-input" type="text" name="departamento" disabled></td>
        </tr>

        <tr>
          <td class="label">N° de empleados por puesto de trabajo</td>
          <td class="field">
            <input class="cell-input" type="number" min="0" name="ptm_num_empleados" value="{{ old('ptm_num_empleados') }}">
          </td>
        </tr>

        <tr>
          <td class="label">Descripción general del puesto</td>
          <td class="field">
            <textarea class="cell-input" name="ptm_descripcion_general" rows="3">{{ old('ptm_descripcion_general') }}</textarea>
          </td>
        </tr>

        <tr>
          <td class="label">Actividades diarias</td>
          <td class="field">
            <textarea class="cell-input" name="ptm_actividades_diarias" rows="3">{{ old('ptm_actividades_diarias') }}</textarea>
          </td>
        </tr>

        <tr>
          <td class="label">Objetivo del puesto</td>
          <td class="field">
            <textarea class="cell-input objetivo" name="ptm_objetivo_puesto" rows="3">{{ old('ptm_objetivo_puesto') }}</textarea>
          </td>
        </tr>
      </table>
      </div>
      <aside class="puestos-sidebar">
        <div class="puestos-sidebar__header">
          <span class="puestos-sidebar__title">Puestos</span>
          <div class="puestos-sidebar__legend">
            <span><span class="status-dot status-dot--con"></span> Listo</span>
            <span><span class="status-dot status-dot--sin"></span> Pendiente</span>
          </div>
          <div class="puestos-sidebar__search">
            <input type="text" id="puestos-search" class="puestos-sidebar__input" placeholder="Buscar puesto...">
          </div>
        </div>
        <div class="puestos-sidebar__scroll">
          <table class="puestos-sidebar__table">
            <tbody>
              @forelse ($puestos as $p)
              <tr class="puestos-sidebar__row {{ $p->tiene_ident ? 'puestos-sidebar__row--with' : 'puestos-sidebar__row--pending' }}" data-puesto-id="{{ $p->id_puesto_trabajo_matriz }}" data-nombre="{{ $p->puesto_trabajo_matriz }}">
                <td class="puestos-sidebar__name">{{ $p->puesto_trabajo_matriz }}</td>
                <td class="puestos-sidebar__status">
                  <span class="status-dot {{ $p->tiene_ident ? 'status-dot--con' : 'status-dot--sin' }}"></span>
                  {{ $p->tiene_ident ? 'Listo' : 'Pendiente' }}
                </td>
              </tr>
              @empty
              <tr class="puestos-sidebar__row puestos-sidebar__row--empty">
                <td colspan="2">No hay puestos disponibles.</td>
              </tr>
              @endforelse
              <tr class="puestos-sidebar__row puestos-sidebar__row--no-results" style="display:none;">
                <td colspan="2">Sin resultados.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </aside>
    </div>


    <!-- === estilos mínimos extra (si aún no los tienes) === -->
    <style>
      .sheet thead th{ background:#bfbfbf; text-align:center; font-weight:700; padding:6px; }
      .sheet .label-cell{ padding:8px; font-size:12px; }
      .cell-input{ width:100%; border:0; padding:8px; font:inherit; background:transparent; outline:none; }
    </style>

<datalist id="fisico-desc-list">
  <option value="N/A"></option>
  <option value="Bomba de Aspersion"></option>
  <option value="Fardos de Bolsas Plasticas"></option>
  <option value="Master Vacíos"></option>
  <option value="Herramientas"></option>
  <option value="Lockers"></option>
  <option value="Estantes"></option>
  <option value="Escaleras"></option>
  <option value="Bolsas de Hidroxido de Calcio"></option>
  <option value="Master"></option>
  <option value="Extintores"></option>
  <option value="Tuberia PVC"></option>
  <option value="Cloro Liquido"></option>
  <option value="Bines, Utensilios"></option>
  <option value="Paletas con Master"></option>
  <option value="Bines"></option>
  <option value="Paleo de Hielo"></option>
  <option value="Basura"></option>
  <option value="Canasta con Camaron"></option>
</datalist>
<datalist id="fisico-capacitacion-list">
  @foreach (($capacitacionesCatalogo ?? []) as $cap)
    @if (!empty($cap->capacitacion))
      <option value="{{ $cap->capacitacion }}"></option>
    @endif
  @endforeach
  <option value="N/A"></option>
</datalist>
<datalist id="fisico-epp-list">
  @foreach (($equiposEpp ?? []) as $epp)
    @if (!empty($epp->equipo))
      <option value="{{ $epp->equipo }}"></option>
    @endif
  @endforeach
  <option value="N/A"></option>
</datalist>

    <!-- ===================== ESFUERZO FISICO ===================== -->
    <div class="bar" style="margin-top:8px;">ESFUERZO FISICO</div>
    <table class="sheet">
      <colgroup>
        <col style="width:10%">
        <col style="width:15%">
        <col style="width:15%">
        <col style="width:10%">
        <col style="width:10%">
        <col style="width:10%">
        <col style="width:10%">
        <col style="width:10%">
        <col style="width:10%">
      </colgroup>
      <thead>
        <tr>
          <th>Tipos de esfuerzo</th>
          <th>Descripción de Carga</th>
          <th>Equipo de apoyo</th>
          <th>Duración de la actividad</th>
          <th>Distancia de traslado</th>
          <th>Frecuencia de Carga</th>
          <th>EPP utilizado</th>
          <th>Peso aproximado</th>
          <th>Capacitación</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td class="label-cell">Cargar</td>
          <td><input class="cell-input" name="fisico_cargar_desc"   list="fisico-desc-list"></td>
          <td><input class="cell-input" name="fisico_cargar_equipo"></td>
          <td><input class="cell-input" name="fisico_cargar_duracion"></td>
          <td><input class="cell-input" name="fisico_cargar_distancia"></td>
          <td><input class="cell-input" name="fisico_cargar_frecuencia"></td>
          <td><input class="cell-input" name="fisico_cargar_epp" list="fisico-epp-list"></td>
          <td><input class="cell-input" name="fisico_cargar_peso"></td>
          <td><input class="cell-input" name="fisico_cargar_capacitacion" list="fisico-capacitacion-list"></td>
        </tr>
        <tr>
          <td class="label-cell">Halar</td>
          <td><input class="cell-input" name="fisico_halar_desc"    list="fisico-desc-list"></td>
          <td><input class="cell-input" name="fisico_halar_equipo"></td>
          <td><input class="cell-input" name="fisico_halar_duracion"></td>
          <td><input class="cell-input" name="fisico_halar_distancia"></td>
          <td><input class="cell-input" name="fisico_halar_frecuencia"></td>
          <td><input class="cell-input" name="fisico_halar_epp" list="fisico-epp-list"></td>
          <td><input class="cell-input" name="fisico_halar_peso"></td>
          <td><input class="cell-input" name="fisico_halar_capacitacion" list="fisico-capacitacion-list"></td>
        </tr>
        <tr>
          <td class="label-cell">Empujar</td>
          <td><input class="cell-input" name="fisico_empujar_desc"  list="fisico-desc-list"></td>
          <td><input class="cell-input" name="fisico_empujar_equipo"></td>
          <td><input class="cell-input" name="fisico_empujar_duracion"></td>
          <td><input class="cell-input" name="fisico_empujar_distancia"></td>
          <td><input class="cell-input" name="fisico_empujar_frecuencia"></td>
          <td><input class="cell-input" name="fisico_empujar_epp" list="fisico-epp-list"></td>
          <td><input class="cell-input" name="fisico_empujar_peso"></td>
          <td><input class="cell-input" name="fisico_empujar_capacitacion" list="fisico-capacitacion-list"></td>
        </tr>
        <tr>
          <td class="label-cell">Sujetar</td>
          <td><input class="cell-input" name="fisico_sujetar_desc"  list="fisico-desc-list"></td>
          <td><input class="cell-input" name="fisico_sujetar_equipo"></td>
          <td><input class="cell-input" name="fisico_sujetar_duracion"></td>
          <td><input class="cell-input" name="fisico_sujetar_distancia"></td>
          <td><input class="cell-input" name="fisico_sujetar_frecuencia"></td>
          <td><input class="cell-input" name="fisico_sujetar_epp" list="fisico-epp-list"></td>
          <td><input class="cell-input" name="fisico_sujetar_peso"></td>
          <td><input class="cell-input" name="fisico_sujetar_capacitacion" list="fisico-capacitacion-list"></td>
        </tr>
      </tbody>
    </table>

    <script>
  // --- ESFUERZO FÍSICO: si desc = N/A => resto = NA ---
  (function(){
    const q  = sel => document.querySelector(sel);
    const setVal = (name, val) => { const el = q(`[name="${name}"]`); if (el) el.value = val ?? ''; };
    const isNA = v => {
      const t = (v || '').toString().trim().toUpperCase();
      return t === 'NA' || t === 'N/A';
    };

    const filas = [
      {k:'cargar'},
      {k:'halar'},
      {k:'empujar'},
      {k:'sujetar'}
    ];

    filas.forEach(({k}) => {
      const desc = q(`[name="fisico_${k}_desc"]`);
      if (!desc) return;

      // me aseguro que tenga el datalist
      desc.setAttribute('list', 'fisico-desc-list');

      const resto = [
        `fisico_${k}_equipo`,
        `fisico_${k}_duracion`,
        `fisico_${k}_distancia`,
        `fisico_${k}_epp`,
        `fisico_${k}_frecuencia`,
        `fisico_${k}_peso`,
        `fisico_${k}_capacitacion`
      ];

      const applyNA = () => {
        if (isNA(desc.value)) {
          resto.forEach(n => setVal(n, 'N/A'));
        }
        // Si quieres que se limpien cuando deje de ser NA, descomenta:
        // else { resto.forEach(n => { if (q(`[name="${n}"]`)?.value === 'NA') setVal(n, ''); }); }
      };

      desc.addEventListener('change', applyNA);
      desc.addEventListener('blur', applyNA);
    });
  })();
</script>

<!-- ===================== ESFUERZO VISUAL ===================== -->
<div class="bar" style="margin-top:8px;">ESFUERZO VISUAL</div>
<table class="sheet" id="tabla-visual">
  <colgroup>
    <col style="width:65%">
    <col style="width:30%">
    <col style="width:5%">
  </colgroup>
  <thead>
    <tr>
      <th>Tipo de esfuerzo</th>
      <th>Tiempo de exposición</th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr class="row-visual">
      <td>
        <input class="cell-input visual-tipo" name="visual[0][tipo]" list="visual-list" placeholder="Escribe para buscar…" autocomplete="off">
        <input type="hidden" class="visual-id" name="visual[0][id_tipo]">
        <datalist id="visual-list">
          <!-- Opcional: llena desde catálogo si lo tienes
          @foreach (($visualOpciones ?? []) as $v)
            <option data-id="{{ $v->id }}" value="{{ $v->nombre }}"></option>
          @endforeach
          -->
          <!-- Ejemplos (puedes borrar) -->
          <option value="Natural"></option>
          <option value="Brillo Computadora"></option>
          <option value="Brillo Monitor"></option>
          <option value="Lectura Prolongada de Documentos"></option>
          <option value="Soldadura"></option>
          <option value="Ninguno"></option>
        </datalist>
      </td>
      <td><input class="cell-input" name="visual[0][tiempo]" placeholder="ej. 3 h/día"></td>
      <td style="text-align:center;vertical-align:middle;">
        <button type="button" class="icon-btn add" aria-label="Añadir fila">+</button>
      </td>
    </tr>
  </tbody>
</table>


<!-- ===================== EXPOSICIÓN A RUIDO ===================== -->
<div class="bar" style="margin-top:8px;">EXPOSICIÓN A RUIDO</div>
<table class="sheet" id="tabla-ruido">
  <colgroup>
    <col style="width:55%">
    <col style="width:20%">
    <col style="width:20%">
    <col style="width:5%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción de ruido</th>
      <th>Duración de exposición</th>
      <th>EPP utilizado</th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr class="row-ruido">
      <td>
        <input class="cell-input ruido-desc" name="ruido[0][desc]" list="ruido-list" placeholder="Escribe para buscar…" autocomplete="off">
        <input type="hidden" class="ruido-id" name="ruido[0][id_tipo]">
        <datalist id="ruido-list">
          <!-- @foreach (($ruidoOpciones ?? []) as $r)
            <option data-id="{{ $r->id }}" value="{{ $r->nombre }}"></option>
          @endforeach -->
          <option value="Calderas"></option>
          <option value="Compresores de Amoniaco"></option>
          <option value="Fabrica de Hielo"></option>
          <option value="Generadores"></option>
          <option value="Maquina de IQF"></option>
          <option value="Maquina de Clasificado"></option>
          <option value="Ninguno"></option>
        </datalist>
      </td>
      <td><input class="cell-input" name="ruido[0][duracion]"></td>
      <td><input class="cell-input" name="ruido[0][epp]" list="fisico-epp-list"></td>
      <td style="text-align:center;vertical-align:middle;">
        <button type="button" class="icon-btn add" aria-label="Añadir fila">+</button>
      </td>
    </tr>
  </tbody>
</table>


<!-- ===================== EXPOSICIÓN STRESS TÉRMICO ===================== -->
<div class="bar" style="margin-top:8px;">EXPOSICIÓN STRESS TÉRMICO</div>
<table class="sheet" id="tabla-termico">
  <colgroup>
    <col style="width:55%">
    <col style="width:20%">
    <col style="width:20%">
    <col style="width:5%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción de stress térmico</th>
      <th>Duración de exposición</th>
      <th>EPP utilizado</th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr class="row-termico">
      <td>
        <input class="cell-input termico-desc" name="termico[0][desc]" list="termico-list" placeholder="Escribe para buscar…" autocomplete="off">
        <input type="hidden" class="termico-id" name="termico[0][id_tipo]">
        <datalist id="termico-list">
          <!-- @foreach (($termicoOpciones ?? []) as $t)
            <option data-id="{{ $t->id }}" value="{{ $t->nombre }}"></option>
          @endforeach -->
          <option value="Temperatura de Sala"></option>
          <option value="Temperatura de Oficina"></option>
          <option value="Temperatura de Cocina"></option>
          <option value="Temperatura Ambiente"></option>
          <option value="Baja Temperatura (Holding)"></option>
          <option value="Baja Temperatura (IQF)"></option>
          <option value="Baja Temperatura (Sala de Hielo)"></option>
        </datalist>
      </td>
      <td><input class="cell-input" name="termico[0][duracion]"></td>
      <td><input class="cell-input" name="termico[0][epp]" list="fisico-epp-list"></td>
      <td style="text-align:center;vertical-align:middle;">
        <button type="button" class="icon-btn add" aria-label="Añadir fila">+</button>
      </td>
    </tr>
  </tbody>
</table>

<script>
(() => {
  function idFromDatalist(input){
    const listId = input.getAttribute('list');
    const dl = document.getElementById(listId);
    const val = (input.value || '').trim().toLowerCase();
    if (!dl) return '';
    const opt = Array.from(dl.options || []).find(o => (o.value || '').trim().toLowerCase() === val);
    return opt ? (opt.getAttribute('data-id') || '') : '';
  }

  function setupDynTable(cfg){
    const tabla = document.getElementById(cfg.tableId);
    if (!tabla) return;
    const tbody = tabla.querySelector('tbody');

    function renumerar(){
      [...tbody.querySelectorAll('tr')].forEach((tr, i) => {
        tr.querySelectorAll('input, select, textarea').forEach(el => {
          if (!el.name) return;
          el.name = el.name.replace(/\[\d+\]/, `[${i}]`);
        });
      });
    }

    function actualizarBotones(){
      const filas = tbody.querySelectorAll('tr');
      filas.forEach((tr, i) => {
        const btn = tr.querySelector('.icon-btn');
        if (!btn) return;
        if (i === filas.length - 1){
          btn.textContent = '+';
          btn.classList.add('add'); btn.classList.remove('remove');
          btn.setAttribute('aria-label','Añadir fila');
        }else{
          btn.textContent = '×';
          btn.classList.remove('add'); btn.classList.add('remove');
          btn.setAttribute('aria-label','Eliminar fila');
        }
      });
    }

    function wireRow(tr){
      const visible = tr.querySelector(cfg.visibleSelector);
      const hidden  = tr.querySelector(cfg.hiddenSelector);
      if (visible && hidden){
        const upd = () => hidden.value = idFromDatalist(visible);
        visible.addEventListener('input', upd);
        visible.addEventListener('change', upd);
        upd();
      }
    }

    tabla.addEventListener('click', (e) => {
      const btn = e.target.closest('.icon-btn');
      if (!btn) return;
      const fila = btn.closest('tr');

      if (btn.classList.contains('add')){
        const clon = fila.cloneNode(true);
        clon.querySelectorAll('input, textarea, select').forEach(i => {
          if (i.type === 'hidden') i.value = '';
          else i.value = '';
          if (i.type === 'checkbox' || i.type === 'radio') i.checked = false;
        });
        tbody.appendChild(clon);
        renumerar();
        actualizarBotones();
        wireRow(clon);
      } else if (btn.classList.contains('remove')){
        fila.remove();
        renumerar();
        actualizarBotones();
      }
    });

    // init
    wireRow(tbody.querySelector('tr'));
    actualizarBotones();
  }

  // Visual
  setupDynTable({
    tableId: 'tabla-visual',
    visibleSelector: '.visual-tipo',
    hiddenSelector:  '.visual-id'
  });

  // Ruido
  setupDynTable({
    tableId: 'tabla-ruido',
    visibleSelector: '.ruido-desc',
    hiddenSelector:  '.ruido-id'
  });

  // Térmico
  setupDynTable({
    tableId: 'tabla-termico',
    visibleSelector: '.termico-desc',
    hiddenSelector:  '.termico-id'
  });
})();
</script>


    <!-- ===== EXPOSICIÓN A QUÍMICOS ===== -->
<div class="bar" style="margin-top:8px;">EXPOSICIÓN A QUÍMICOS</div>
<table class="sheet" id="tabla-quimicos">
  <colgroup>
    <col style="width:20%">
    <col style="width:20%">
    <col style="width:20%">
    <col style="width:20%">
    <col style="width:20%">
    <col style="width:28px">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción del químico</th>
      <th>Capacitación</th>
      <th>Duración de exposición</th>
      <th>Frecuencia</th>
      <th>EPP utilizado</th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr class="row-quimico">
    <td>
      <!-- visible: nombre; hidden: id_quimico -->
      <input class="cell-input quimico-nombre" name="quimicos[0][nombre]" list="quimico-list"
            placeholder="Escribe para buscar…" autocomplete="off">
      <input type="hidden" class="quimico-id" name="quimicos[0][id_quimico]">
      <datalist id="quimico-list">
        @foreach ($quimicos as $q)
          <option data-id="{{ $q->id_quimico }}" value="{{ $q->nombre_comercial }}"></option>
        @endforeach
      </datalist>
    </td>
    <td><input class="cell-input" name="quimicos[0][capacitacion]" list="fisico-capacitacion-list"></td>
    <td><input class="cell-input" name="quimicos[0][duracion]"></td>
    <td><input class="cell-input" name="quimicos[0][frecuencia]"></td>
        <td><input class="cell-input" name="quimicos[0][epp]" list="fisico-epp-list"></td>
    <td style="text-align:center; vertical-align:middle;">
      <button type="button" class="icon-btn add" aria-label="Añadir químico">+</button>
    </td>
  </tr>
  </tbody>
</table>

<style>
  /* Botones de + / × dentro de la celda angosta */
  .icon-btn{
    width:100%; height:100%;
    min-height:34px;
    display:flex; align-items:center; justify-content:center;
    font-size:20px; line-height:1; background:transparent; border:0; cursor:pointer;
  }
  .icon-btn:focus{ outline:2px solid #666; outline-offset:-2px; }
</style>

<script>
(() => {
  const tabla = document.getElementById('tabla-quimicos');
  const tbody = tabla.querySelector('tbody');

  function idFromDatalist(input) {
    const listId = input.getAttribute('list');
    const dl = document.getElementById(listId);
    const val = (input.value || '').trim().toLowerCase();
    if (!dl) return '';
    const opt = Array.from(dl.options || []).find(o =>
      (o.value || '').trim().toLowerCase() === val
    );
    return opt ? (opt.getAttribute('data-id') || '') : '';
  }

  function wireRow(tr) {
    const nombre = tr.querySelector('input.quimico-nombre');
    const hidden = tr.querySelector('input.quimico-id');
    if (!nombre || !hidden) return;
    const upd = () => { hidden.value = idFromDatalist(nombre) };
    nombre.addEventListener('input', upd);
    nombre.addEventListener('change', upd);
    upd(); // inicializar
  }

  function renumerar() {
    [...tbody.querySelectorAll('tr')].forEach((tr, i) => {
      tr.querySelectorAll('select, input').forEach(el => {
        if (el.name) el.name = el.name.replace(/\[\d+\]/, `[${i}]`);
      });
    });
  }

  function actualizarBotones() {
    const filas = tbody.querySelectorAll('tr');
    filas.forEach((tr, i) => {
      const btn = tr.querySelector('.icon-btn');
      if (i === filas.length - 1) {
        btn.textContent = '+';
        btn.classList.add('add');
        btn.classList.remove('remove');
        btn.setAttribute('aria-label','Añadir químico');
      } else {
        btn.textContent = '×';
        btn.classList.remove('add');
        btn.classList.add('remove');
        btn.setAttribute('aria-label','Eliminar fila');
      }
    });
  }

  tabla.addEventListener('click', (e) => {
    const btn = e.target.closest('.icon-btn');
    if (!btn) return;
    const fila = btn.closest('tr');

    if (btn.classList.contains('add')) {
      const clon = fila.cloneNode(true);
      clon.querySelectorAll('input').forEach(i => i.value = '');
      tbody.appendChild(clon);
      renumerar();
      actualizarBotones();
      wireRow(clon); // <-- importante
    } else if (btn.classList.contains('remove')) {
      fila.remove();
      renumerar();
      actualizarBotones();
    }
  });

  // inicializar primera fila
  wireRow(tbody.querySelector('tr'));
  actualizarBotones();
})();
</script>


<!-- ===== CONDICIONES DE INSTALACIONES ===== -->
<div class="bar" style="margin-top:8px;">CONDICIONES DE INSTALACIONES</div>
<table class="sheet">
  <colgroup>
    <col style="width:40%">
    <col style="width:10%">
    <col style="width:10%">
    <col style="width:10%">
    <col style="width:30%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción del elemento</th>
      <th>Adecuado</th>
      <th>No adecuado</th>
      <th>N/A</th>
      <th>Observaciones</th>
    </tr>
  </thead>
  <tbody>
    <!-- Fila 1 -->
    <tr>
      <td class="label-cell">Paredes, muros, losas y trabes</td>
      <td class="center"><input type="radio" name="instalaciones[0][estado]" value="A"></td>
      <td class="center"><input type="radio" name="instalaciones[0][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="instalaciones[0][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="instalaciones[0][obs]"></td>
    </tr>
    <!-- Fila 2 -->
    <tr>
      <td class="label-cell">Pisos</td>
      <td class="center"><input type="radio" name="instalaciones[1][estado]" value="A"></td>
      <td class="center"><input type="radio" name="instalaciones[1][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="instalaciones[1][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="instalaciones[1][obs]"></td>
    </tr>
    <!-- Fila 3 -->
    <tr>
      <td class="label-cell">Techos</td>
      <td class="center"><input type="radio" name="instalaciones[2][estado]" value="A"></td>
      <td class="center"><input type="radio" name="instalaciones[2][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="instalaciones[2][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="instalaciones[2][obs]"></td>
    </tr>
    <!-- Fila 4 -->
    <tr>
      <td class="label-cell">Puertas y Ventanas</td>
      <td class="center"><input type="radio" name="instalaciones[3][estado]" value="A"></td>
      <td class="center"><input type="radio" name="instalaciones[3][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="instalaciones[3][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="instalaciones[3][obs]"></td>
    </tr>
    <!-- Fila 5 -->
    <tr>
      <td class="label-cell">Escaleras y rampas</td>
      <td class="center"><input type="radio" name="instalaciones[4][estado]" value="A"></td>
      <td class="center"><input type="radio" name="instalaciones[4][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="instalaciones[4][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="instalaciones[4][obs]"></td>
    </tr>
    <!-- Fila 6 -->
    <tr>
      <td class="label-cell">Anaqueles y estantería</td>
      <td class="center"><input type="radio" name="instalaciones[5][estado]" value="A"></td>
      <td class="center"><input type="radio" name="instalaciones[5][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="instalaciones[5][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="instalaciones[5][obs]"></td>
    </tr>
  </tbody>
</table>

<!-- ===== MAQUINARIA, EQUIPO Y HERRAMIENTAS ===== -->
<div class="bar" style="margin-top:8px;">MAQUINARIA, EQUIPO Y HERRAMIENTAS</div>
<table class="sheet">
  <colgroup>
    <col style="width:40%">
    <col style="width:10%">
    <col style="width:10%">
    <col style="width:10%">
    <col style="width:30%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción del elemento</th>
      <th>Adecuado</th>
      <th>No adecuado</th>
      <th>N/A</th>
      <th>Observaciones</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="label-cell">Estado de Maquinaria y Equipo</td>
      <td class="center"><input type="radio" name="maq[0][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[0][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[0][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[0][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Se ejecuta mantenimiento preventivo</td>
      <td class="center"><input type="radio" name="maq[1][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[1][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[1][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[1][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Se ejecuta mantenimiento correctivo</td>
      <td class="center"><input type="radio" name="maq[2][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[2][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[2][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[2][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Estado resguardos y guardas</td>
      <td class="center"><input type="radio" name="maq[3][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[3][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[3][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[3][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Estado de conexiones eléctricas de maquinaria</td>
      <td class="center"><input type="radio" name="maq[4][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[4][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[4][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[4][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Se realizan inspecciones de Maquinaria y Equipos</td>
      <td class="center"><input type="radio" name="maq[5][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[5][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[5][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[5][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Estado de paros de emergencia</td>
      <td class="center"><input type="radio" name="maq[6][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[6][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[6][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[6][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Tiene el entrenamiento correspondiente para uso de maquinaria y equipo</td>
      <td class="center"><input type="radio" name="maq[7][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[7][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[7][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[7][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Utiliza el EPP correspondiente, especifique cuál</td>
      <td class="center"><input type="radio" name="maq[8][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[8][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[8][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[8][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Estado de herramientas</td>
      <td class="center"><input type="radio" name="maq[9][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[9][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[9][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[9][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Se realizan inspecciones de Herramientas</td>
      <td class="center"><input type="radio" name="maq[10][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[10][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[10][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[10][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Almacenamiento Correcto de Herramientas</td>
      <td class="center"><input type="radio" name="maq[11][estado]" value="A"></td>
      <td class="center"><input type="radio" name="maq[11][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="maq[11][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="maq[11][obs]"></td>
    </tr>
  </tbody>
</table>

<!-- ===== EQUIPOS Y SERVICIOS DE EMERGENCIA ===== -->
<div class="bar" style="margin-top:8px;">EQUIPOS Y SERVICIOS DE EMERGENCIA</div>
<table class="sheet">
  <colgroup>
    <col style="width:40%">
    <col style="width:10%">
    <col style="width:10%">
    <col style="width:10%">
    <col style="width:30%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción de elemento</th>
      <th>Adecuado</th>
      <th>No adecuado</th>
      <th>N/A</th>
      <th>Observaciones</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="label-cell">Señalización rutas de evacuación y salidas de emergencia y punto reunión</td>
      <td class="center"><input type="radio" name="emer[0][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[0][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[0][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[0][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Ubicación de Extintores o mangueras de incendios</td>
      <td class="center"><input type="radio" name="emer[1][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[1][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[1][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[1][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Ubicación de Camillas y elementos de primeros auxilios</td>
      <td class="center"><input type="radio" name="emer[2][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[2][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[2][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[2][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Ubicación de Botiquín</td>
      <td class="center"><input type="radio" name="emer[3][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[3][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[3][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[3][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Realización de Simulacros</td>
      <td class="center"><input type="radio" name="emer[4][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[4][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[4][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[4][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Socialización Plan de evacuación</td>
      <td class="center"><input type="radio" name="emer[5][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[5][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[5][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[5][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Capacitación sobre actuación en caso de emergencia y uso de extintor.</td>
      <td class="center"><input type="radio" name="emer[6][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[6][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[6][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[6][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Ubicación de alarmas aviso emergencia</td>
      <td class="center"><input type="radio" name="emer[7][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[7][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[7][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[7][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Ubicación de alarmas de humo</td>
      <td class="center"><input type="radio" name="emer[8][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[8][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[8][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[8][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Ubicación de lámparas de emergencia</td>
      <td class="center"><input type="radio" name="emer[9][estado]" value="A"></td>
      <td class="center"><input type="radio" name="emer[9][estado]" value="NA"></td>
      <td class="center"><input type="radio" name="emer[9][estado]" value="N/A"></td>
      <td><input class="cell-input obs-input" name="emer[9][obs]"></td>
    </tr>
  </tbody>
</table>

<!-- ===== RIESGO DE FUEGO O EXPLOSIÓN ===== -->
<div class="bar" style="margin-top:8px;">RIESGO DE FUEGO O EXPLOSIÓN</div>
<table class="sheet">
  <colgroup>
    <col style="width:33.33%">
    <col style="width:33.33%">
    <col style="width:33.34%">
  </colgroup>
  <tbody>
    <tr>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[inflamables_area]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Hay sustancias inflamables o combustible o material explosivo en su área de trabajo</span>
        </label>
      </td>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[ventilacion_extraccion]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Hay ventilación natural suficiente o sistema de extracción</span>
        </label>
      </td>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[limpieza_regulares]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Se realiza limpiezas regulares en el área de trabajo</span>
        </label>
      </td>
    </tr>

    <tr>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[senalizacion_riesgo]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Está la debida señalización de riesgo inflamable o de explosión si se requiere</span>
        </label>
      </td>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[focos_ignicion]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Hay fuentes de calor o focos de ignición cercanos</span>
        </label>
      </td>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[riesgo_electrico_friccion]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Hay maquinaria o equipo que represente un riesgo eléctrico o de fricción</span>
        </label>
      </td>
    </tr>

    <tr>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[trasiego_combustibles]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Realizan trasiego de líquidos combustibles o inflamables</span>
        </label>
      </td>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[cilindros_alta_presion]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Trabaja con cilindros de alta presión, están ubicados en su estructura y con cadena seguridad</span>
        </label>
      </td>
      <td>
        <label class="checkcell">
          <select name="riesgo_fuego[derrames_combustibles]">
            <option value="" selected>—</option>
            <option value="si">Sí</option>
            <option value="no">No</option>
            <option value="na">N/A</option>
          </select>
          <span>Hay derrames de sustancia combustibles o inflamables</span>
        </label>
      </td>
    </tr>
  </tbody>
</table>

<!-- estilos de apoyo para las celdas con checkbox -->
<style>
  .checkcell{ display:flex; align-items:flex-start; gap:8px; padding:8px; }
  .checkcell input[type="checkbox"]{ transform: scale(1.15); margin-top:2px; }
</style>

<!-- ===== POSICIONES O MOVIMIENTOS DE TRABAJO ERGONÓMICO ===== -->
<div class="bar" style="margin-top:8px;">POSICIONES O MOVIMIENTOS DE TRABAJO ERGONÓMICO</div>
<table class="sheet">
  <colgroup>
    <col style="width:50%">
    <col style="width:7%">
    <col style="width:7%">
    <col style="width:7%">
    <col style="width:29%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción de elemento</th>
      <th>SI</th>
      <th>NO</th>
      <th>N/A</th>
      <th>Observaciones</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="label-cell">Realiza movimientos repetitivos, ¿con qué frecuencia y duración?</td>
      <td class="center"><input type="radio" name="ergo[0][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[0][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[0][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[0][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Realizo posturas forzadas, ¿durante cuánto tiempo?</td>
      <td class="center"><input type="radio" name="ergo[1][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[1][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[1][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[1][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Tengo suficiente espacio para realizar la tarea</td>
      <td class="center"><input type="radio" name="ergo[2][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[2][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[2][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[2][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Realizo elevación de brazos, ¿a qué grados?</td>
      <td class="center"><input type="radio" name="ergo[3][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[3][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[3][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[3][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Realiza movimiento o giros de la muñeca</td>
      <td class="center"><input type="radio" name="ergo[4][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[4][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[4][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[4][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Realizo inclinación de espalda o cuello, ¿a qué grados?</td>
      <td class="center"><input type="radio" name="ergo[5][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[5][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[5][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[5][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Manipula alguna herramienta o utensilio constantemente</td>
      <td class="center"><input type="radio" name="ergo[6][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[6][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[6][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[6][obs]"></td>
    </tr>
    <tr>
      <td class="label-cell">Manipula herramienta que genere vibración, ¿cuál y por cuánto tiempo?</td>
      <td class="center"><input type="radio" name="ergo[7][resp]" value="SI"></td>
      <td class="center"><input type="radio" name="ergo[7][resp]" value="NO"></td>
      <td class="center"><input type="radio" name="ergo[7][resp]" value="NA"></td>
      <td><input class="cell-input" name="ergo[7][obs]"></td>
    </tr>
  </tbody>
</table>

<!-- ===== POSTURAS (matriz de checkboxes) ===== -->
<div class="bar" style="margin-top:8px;">POSTURAS</div>
<table class="sheet">
  <colgroup>
    <col style="width:16.66%">
    <col style="width:16.66%">
    <col style="width:16.66%">
    <col style="width:16.66%">
    <col style="width:16.66%">
    <col style="width:16.70%">
  </colgroup>
  <tbody>
    <tr>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="agachado"><span>Agachado</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="de_rodillas"><span>De rodillas</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="volteado"><span>Volteado</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="parado"><span>Parado</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="sentado"><span>Sentado</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="arrastrandose"><span>Arrastrándose</span>
        </label>
      </td>
    </tr>
    <tr>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="subiendo"><span>Subiendo</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="balanceandose"><span>Balanceándose</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="corriendo"><span>Corriendo</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="empujando"><span>Empujando</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="halando"><span>Halando</span>
        </label>
      </td>
      <td class="center">
        <label class="checkcell">
          <input type="checkbox" name="posturas[]" value="girando"><span>Girando</span>
        </label>
      </td>
    </tr>
  </tbody>
</table>

<!-- ===== RIESGO CONDICIONES DE TRABAJO ===== -->
<div class="bar" style="margin-top:8px;">RIESGO CONDICIONES DE TRABAJO</div>

<!-- Subtítulo centrado como en el formato -->
<table class="sheet">
  <thead>
    <tr>
      <th colspan="4" class="subhead">TRABAJO EN ALTURAS</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="label-cell" style="width:25%;">Altura</td>
      <td style="width:25%;"><input class="cell-input" name="alturas[altura]"></td>
      <td class="label-cell" style="width:25%;">Cuenta con medios de anclaje y acceso seguro</td>
      <td style="width:25%;"><input class="cell-input" name="alturas[anclaje_seguro]"></td>
    </tr>
    <tr>
      <td class="label-cell" style="width:25%;">Inspección estado de EPP</td>
      <td style="width:25%;"><input class="cell-input" name="alturas[inspeccion]"></td>
      <td class="label-cell" style="width:25%;">EPP utilizado</td>
      <td style="width:25%;"><input class="cell-input" name="alturas[epp]" list="fisico-epp-list"></td>
    </tr>
    <tr>
      <td class="label-cell" style="width:25%;">Señalización</td>
      <td style="width:25%;"><input class="cell-input" name="alturas[senalizacion]"></td>
      <td class="label-cell" style="width:25%;">Capacitación Recibida</td>
      <td style="width:25%;"><input class="cell-input" name="alturas[capacitacion]" list="fisico-capacitacion-list"></td>
    </tr>
    <tr>
      <td class="label-cell">Se da Aviso del trabajo en altura</td>
      <td><input class="cell-input" name="alturas[aviso_trabajo_altura]"></td>
      <td class="label-cell">Firma hoja de trabajo seguro</td>
      <td><input class="cell-input" name="alturas[firma_trabajo_seguro]"></td>
    </tr>
  </tbody>
</table>

<!-- estilos de apoyo (si aún no estaban) -->
<style>
  td.center input[type="radio"]{ transform: scale(1.2); }
  .checkcell{ display:flex; align-items:center; gap:6px; padding:6px; }
  .checkcell input[type="checkbox"]{ transform: scale(1.1); }
  .subhead{ background:transparent !important; font-weight:700; text-transform:uppercase; text-align:center; }
</style>

<!-- ===== TRABAJO CON ELECTRICIDAD / RIESGO ELÉCTRICO ===== -->
<div class="bar" style="margin-top:8px;">TRABAJO CON ELECTRICIDAD / RIESGO ELÉCTRICO</div>
<table class="sheet">
  <colgroup>
    <col style="width:16.66%"><col style="width:16.66%">
    <col style="width:16.66%"><col style="width:16.66%">
    <col style="width:16.66%"><col style="width:16.70%">
  </colgroup>
  <tbody>
  <tr>
    <td class="label-cell">Señalización y delimitación</td>
    <td>
      <select class="cell-input" name="elec_senalizacion">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>

    <td class="label-cell">Capacitación o certificación recibida</td>
    <td>
      <select class="cell-input" name="elec_capacitacion">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>

    <td class="label-cell">Hay alta tensión en su área de trabajo y está señalada y con acceso restringido</td>
    <td>
      <select class="cell-input" name="elec_alta_tension">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>
  </tr>

  <tr>
    <td class="label-cell">Firma de Hoja de Trabajo Seguro</td>
    <td>
      <select class="cell-input" name="elec_hoja_trabajo_seguro">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>

    <td class="label-cell">EPP utilizado</td>
    <td>
      <select class="cell-input" name="elec_epp_utilizado">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>

    <td class="label-cell">Hay zonas de electricidad estática</td>
    <td>
      <select class="cell-input" name="elec_estatica">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>
  </tr>

  <tr>
    <td class="label-cell">Sistema de bloqueo con tarjeta y candado</td>
    <td>
      <select class="cell-input" name="elec_bloqueo">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>

  <td class="label-cell">Verificación de ausencia de tensión antes de trabajar</td>
    <td>
      <select class="cell-input" name="elec_ausencia_tension">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>

    <td class="label-cell">Se da aviso de la realización trabajo eléctrico</td>
    <td>
      <select class="cell-input" name="elec_aviso">
        <option value="" selected>—</option>
        <option value="si">Sí</option>
        <option value="no">No</option>
        <option value="na">N/A</option>
      </select>
    </td>
  </tr>

   <!-- Instrucciones -->
<tr>
  <td class="instrucciones" colspan="6">
    Instrucciones: Responder sí o no en la casilla que le corresponde y cualquier observación colocarla en el espacio de observaciones
  </td>
</tr>

<style>
  .checkcell{ display:flex; align-items:center; justify-content:space-between; gap:8px; padding:8px; }
  .checkcell .cell-input[list="sino-na"]{ max-width:90px; text-align:center; }
</style>

<!-- Matriz de verificación (3 celdas por fila, 100% del ancho) -->
<tr>
  <td colspan="2">
    <label class="checkcell">
      <span>Los cables se encuentran ordenados sin obstaculizar el paso.</span>
      <input class="cell-input" name="elec_verif[cables_ordenados]" list="sino-na" autocomplete="off">
    </label>
  </td>
  <td colspan="2">
    <label class="checkcell">
      <span>Las tomacorriente e interruptores están alejadas de fuentes de humedad o salpicadura de agua.</span>
      <input class="cell-input" name="elec_verif[toma_lejos_humedad]" list="sino-na" autocomplete="off">
    </label>
  </td>
  <td colspan="2">
    <label class="checkcell">
      <span>Las cajas e interruptores se encuentran rotulados y cerrados</span>
      <input class="cell-input" name="elec_verif[cajas_rotuladas_cerradas]" list="sino-na" autocomplete="off">
    </label>
  </td>
</tr>

<tr>
  <td colspan="2">
    <label class="checkcell">
      <span>Interruptores, extensiones, se encuentran en buenas condiciones</span>
      <input class="cell-input" name="elec_verif[buenas_condiciones]" list="sino-na" autocomplete="off">
    </label>
  </td>
  <td colspan="2">
    <label class="checkcell">
      <span>Los cables mantienen su aislamiento en todo el recorrido (sin peladuras)</span>
      <input class="cell-input" name="elec_verif[aislamiento_ok]" list="sino-na" autocomplete="off">
    </label>
  </td>
  <td colspan="2">
    <label class="checkcell">
      <span>Hay señalización de riesgo eléctrico</span>
      <input class="cell-input" name="elec_verif[senalizacion_riesgo]" list="sino-na" autocomplete="off">
    </label>
  </td>
</tr>

    <!-- Observaciones -->
    <tr>
      <td class="label-cell">Observaciones:</td>
      <td colspan="5"><textarea class="cell-input" style="height:70px" name="elec_observaciones"></textarea></td>
    </tr>
  </tbody>
</table>

<datalist id="sino-na">
  <option value="SI"></option>
  <option value="NO"></option>
  <option value="NA"></option>
</datalist>

<!-- ===== RIESGO DE CAÍDA MISMO NIVEL ===== -->
<div class="bar" style="margin-top:8px;">RIESGO DE CAÍDA MISMO NIVEL</div>
<table class="sheet" id="tabla-caida">
  <colgroup>
    <col style="width:64%">
    <col style="width:36%">
  </colgroup>
  <thead>
    <tr>
      <th>Descripción de elemento</th>
      <th>Respuesta (SI/NO/NA)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="label-cell">Estado de pisos es adecuado</td>
      <td><input class="cell-input" name="caida[pisos_adecuados]" list="sino-na" autocomplete="off"></td>
    </tr>
    <tr>
      <td class="label-cell">Las vías de circulación están libres</td>
      <td><input class="cell-input" name="caida[vias_libres]" list="sino-na" autocomplete="off"></td>
    </tr>
    <tr>
      <td class="label-cell">Las rampas o escalones están identificados</td>
      <td><input class="cell-input" name="caida[rampas_identificadas]" list="sino-na" autocomplete="off"></td>
    </tr>
    <tr>
      <td class="label-cell">Las gradas tienen barandas o pasamanos</td>
      <td><input class="cell-input" name="caida[gradas_barandas]" list="sino-na" autocomplete="off"></td>
    </tr>
    <tr>
      <td class="label-cell">Tienen sistema antideslizante en rampas o gradas</td>
      <td><input class="cell-input" name="caida[antideslizantes]" list="sino-na" autocomplete="off"></td>
    </tr>
    <tr>
      <td class="label-cell">Señalización de prevención de riesgo piso resbaloso</td>
      <td><input class="cell-input" name="caida[senalizacion_piso_resbaloso]" list="sino-na" autocomplete="off"></td>
    </tr>
    <tr>
      <td class="label-cell">Observaciones:</td>
      <td><textarea class="cell-input" style="height:70px" name="caida_observaciones"></textarea></td>
    </tr>
  </tbody>
</table>

<div class="bar" style="margin-top:8px;">OTROS RIESGOS</div>
<table class="sheet" id="tabla-caida">

  <thead>
    <tr>
      <th>NOMBRE</th>
      <th>Respuesta (SI/NO)</th>
    </tr>
  </thead>
  <tbody>
        <tr>
      <td class="label-cell">Biológico</td>
      <td><input class="cell-input" name="otros_biologico" list="sino-na"></input></td>
    </tr>
        <tr>
      <td class="label-cell">Psicosocial</td>
      <td><input class="cell-input" name="otros_psicosocial" list="sino-na"></input></td>
    </tr>
        <tr>
      <td class="label-cell">Naturales</td>
      <td><input class="cell-input" name="otros_naturales" list="sino-na"></input></td>
    </tr>
  </tbody>
</table>


<!-- estilos de apoyo -->
<style>
  .instrucciones{ text-align:center; font-size:12px; padding:6px; }
  .checkcell{ display:flex; align-items:flex-start; gap:8px; padding:8px; }
  .checkcell input[type="checkbox"]{ transform: scale(1.1); margin-top:2px; }
</style>

<!-- ===== TABLA DE IDENTIFICACIÓN DE RIESGO ===== -->
<div class="bar" style="margin-top:8px;">TABLA DE IDENTIFICACIÓN DE RIESGO</div>
<table class="sheet risk-table">
  <thead>
    <tr>
      <th rowspan="2" class="center">TIPO</th>
      <th rowspan="2" class="center">PELIGRO</th>
      <th rowspan="2">RIESGO</th>
      <th rowspan="2">CONSECUENCIA</th>
      <th colspan="2" class="center">APLICA</th>
      <th rowspan="2">OBSERVACIONES</th>
      <th colspan="3" class="center">EVALUACIÓN DEL RIESGO</th>
    </tr>
    <tr>
      <th>SI</th>
      <th>NO</th>
      <th>Probabilidad</th>
      <th>Consecuencia</th>
      <th>Nivel de Riesgo</th>
    </tr>
  </thead>
  <tbody>
    <!-- ========== MECÁNICO (7 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="8">MECANICO</td>
      <td class="label-cell">CAÍDA A DESNIVEL</td>
      <td class="label-cell">Caída a mismo nivel</td>
      <td class="label-cell">Fractura/contusiones</td>
      <td class="center"><input type="radio" name="r0_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r0_aplica" value="no"></td>
      <td><input class="cell-input" name="r0_obs"></td>
      <td>
            <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r0_nivel" placeholder=""></td>
    </tr>
    <tr>
      <td class="label-cell">TRABAJO EN ALTURAS</td>
      <td class="label-cell">Caída a distinto nivel</td>
      <td class="label-cell">Muerte/Fractura/Contusiones</td>
      <td class="center"><input type="radio" name="r1_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r1_aplica" value="no"></td>
      <td><input class="cell-input" name="r1_obs"></td>
      <td>
          <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist></td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r1_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">OBJETOS SUSPENDIDOS</td>
      <td class="label-cell">Caída de Objetos</td>
      <td class="label-cell">Muerte/Fractura/Contusiones</td>
      <td class="center"><input type="radio" name="r2_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r2_aplica" value="no"></td>
      <td><input class="cell-input" name="r2_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r2_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">OBJETOS EN MOVIMIENTO Y FIJOS</td>
      <td class="label-cell">Choque contra objetos</td>
      <td class="label-cell">Traumatismo / fractura / contusiones</td>
      <td class="center"><input type="radio" name="r3_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r3_aplica" value="no"></td>
      <td><input class="cell-input" name="r3_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r3_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">EQUIPO. HERRAMIENTAS U OBJETOS PUNZOCORTANTES</td>
      <td class="label-cell">Golpes o cortes</td>
      <td class="label-cell">Heridas/cortes/amputaciones</td>
      <td class="center"><input type="radio" name="r4_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r4_aplica" value="no"></td>
      <td><input class="cell-input" name="r4_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r4_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">PROYECCIÓN DE FRAGMENTO O PARTÍCULAS</td>
      <td class="label-cell">Impacto de fragmentos sobre las personas</td>
      <td class="label-cell">Golpes/ Fracturas/ Contusiones</td>
      <td class="center"><input type="radio" name="r5_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r5_aplica" value="no"></td>
      <td><input class="cell-input" name="r5_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
          </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r5_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">MAQUINARIA, EQUIPO O HERRAMIENTAS EN MOVIMIENTO</td>
      <td class="label-cell">Atrapamiento por o entre objetos</td>
      <td class="label-cell">Amputaciones, contusiones/heridas/contusiones</td>
      <td class="center"><input type="radio" name="r6_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r6_aplica" value="no"></td>
      <td><input class="cell-input" name="r6_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r6_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">MAQUINARIA, EQUIPO O HERRAMIENTAS EN MOVIMIENTO (MONTACARGAS)</td>
      <td class="label-cell">Golpes, atrapamientos o colisiones</td>
      <td class="label-cell">Amputaciones, contusiones/heridas/contusiones</td>
      <td class="center"><input type="radio" name="r33_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r33_aplica" value="no"></td>
      <td><input class="cell-input" name="r33_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r33_nivel"></td>
    </tr>

    <!-- ========== ELÉCTRICO (3 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="3">ELECTRICO</td>
      <td class="label-cell">ALTA O MEDIA TENSION</td>
      <td class="label-cell">Contacto eléctrico directo</td>
      <td class="label-cell">Muerte</td>
      <td class="center"><input type="radio" name="r7_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r7_aplica" value="no"></td>
      <td><input class="cell-input" name="r7_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r7_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">BAJA TENSION</td>
      <td class="label-cell">Contacto eléctrico indirecto</td>
      <td class="label-cell">Quemaduras/Muerte</td>
      <td class="center"><input type="radio" name="r8_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r8_aplica" value="no"></td>
      <td><input class="cell-input" name="r8_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r8_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">ELECTRICIDAD ESTATICA</td>
      <td class="label-cell">Descarga eléctrica estática—incendio, explosión</td>
      <td class="label-cell">Quemaduras/Muerte</td>
      <td class="center"><input type="radio" name="r9_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r9_aplica" value="no"></td>
      <td><input class="cell-input" name="r9_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r9_nivel"></td>
    </tr>

    <!-- ========== FUEGO Y EXPLOSIÓN (2 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="2">FUEGO Y<br>EXPLOSION</td>
      <td class="label-cell">LIQUIDOS, GASES O MATERIALES COMBUSTIBLES O INFLAMABLES</td>
      <td class="label-cell">Explosión/Incendio</td>
      <td class="label-cell">Quemaduras/Muerte</td>
      <td class="center"><input type="radio" name="r10_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r10_aplica" value="no"></td>
      <td><input class="cell-input" name="r10_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r10_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">CILINDROS ALTA PRESION</td>
      <td class="label-cell">Explosión</td>
      <td class="label-cell">Quemaduras/Muerte</td>
      <td class="center"><input type="radio" name="r11_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r11_aplica" value="no"></td>
      <td><input class="cell-input" name="r11_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r11_nivel"></td>
    </tr>
    <tr>
      <td class="tipo" rowspan="4">FISICO</td>
      <td class="label-cell">RUIDO</td>
      <td class="label-cell">Exposición a Ruido</td>
      <td class="label-cell">Hipoacusia</td>
      <td class="center"><input type="radio" name="r15_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r15_aplica" value="no"></td>
      <td><input class="cell-input" name="r15_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r15_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">ESFUERZO VISUAL</td>
      <td class="label-cell">Exposición a radiación luminosa</td>
      <td class="label-cell">Daño a la vista / cansancio visual</td>
      <td class="center"><input type="radio" name="r32_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r32_aplica" value="no"></td>
      <td><input class="cell-input" name="r32_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r32_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">VIBRACION</td>
      <td class="label-cell">Exposición a vibraciones</td>
      <td class="label-cell">Trastornos musculoesqueléticos</td>
      <td class="center"><input type="radio" name="r17_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r17_aplica" value="no"></td>
      <td><input class="cell-input" name="r17_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r17_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">TEMPERATURAS EXTREMAS (FRIO, CALOR)</td>
      <td class="label-cell">Exposición a temperaturas extremas</td>
      <td class="label-cell">Estrés térmico, quemaduras</td>
      <td class="center"><input type="radio" name="r18_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r18_aplica" value="no"></td>
      <td><input class="cell-input" name="r18_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r18_nivel"></td>
    </tr>

    <!-- ========== BIOLOGICO (3 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="3">BIOLOGICO</td>
      <td class="label-cell">BACTERIAS</td>
      <td class="label-cell">Contacto con ambientes o superficies contaminadas</td>
      <td class="label-cell">Intoxicación / Enfermedades</td>
      <td class="center"><input type="radio" name="r19_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r19_aplica" value="no"></td>
      <td><input class="cell-input" name="r19_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r19_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">VIRUS</td>
      <td class="label-cell">Contacto con ambientes o superficies contaminadas</td>
      <td class="label-cell">Intoxicación / Enfermedades</td>
      <td class="center"><input type="radio" name="r20_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r20_aplica" value="no"></td>
      <td><input class="cell-input" name="r20_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r20_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">HONGOS</td>
      <td class="label-cell">Contacto o exposición</td>
      <td class="label-cell">Intoxicación / Enfermedades</td>
      <td class="center"><input type="radio" name="r21_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r21_aplica" value="no"></td>
      <td><input class="cell-input" name="r21_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r21_nivel"></td>
    </tr>

    <!-- ========== ERGONOMICO (4 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="4">ERGONOMICO</td>
      <td class="label-cell">CARGA FÍSICA CON POSTURA FORZADA/PARADO U SENTADO</td>
      <td class="label-cell">Sobreesfuerzo</td>
      <td class="label-cell">Trastornos musculoesqueléticos</td>
      <td class="center"><input type="radio" name="r22_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r22_aplica" value="no"></td>
      <td><input class="cell-input" name="r22_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r22_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">CARGA FÍSICA POR LEVANTAR/MANEJAR OBJETOS PESADOS O HACERLO INADECUADAMENTE</td>
      <td class="label-cell">Sobreesfuerzo</td>
      <td class="label-cell">Lumbalgia</td>
      <td class="center"><input type="radio" name="r23_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r23_aplica" value="no"></td>
      <td><input class="cell-input" name="r23_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r23_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">TAREAS REPETITIVAS</td>
      <td class="label-cell">Probabilidad del daño</td>
      <td class="label-cell">Trastornos musculoesqueléticos</td>
      <td class="center"><input type="radio" name="r24_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r24_aplica" value="no"></td>
      <td><input class="cell-input" name="r24_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r24_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">PROBLEMAS DEL DISEÑO DE LUGAR DE TRABAJO</td>
      <td class="label-cell">Probabilidad del daño</td>
      <td class="label-cell">Trastornos musculoesqueléticos</td>
      <td class="center"><input type="radio" name="r25_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r25_aplica" value="no"></td>
      <td><input class="cell-input" name="r25_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r25_nivel"></td>
    </tr>

    <!-- ========== LOCATIVO (3 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="8">LOCATIVO</td>
      <td class="label-cell">ESCALERAS Y RAMPLAS EN MAL ESTADO/MAL DISEÑADAS</td>
      <td class="label-cell">Caída y golpes</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r26_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r26_aplica" value="no"></td>
      <td><input class="cell-input" name="r26_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r26_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">MAL DISEÑO DE VÍAS DE EVACUACIÓN (ANCHO, PENDIENTE, ALTURA)</td>
      <td class="label-cell">Bloqueo, congestión o caídas durante emergencias</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r27_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r27_aplica" value="no"></td>
      <td><input class="cell-input" name="r27_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r27_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">INFRAESTRUCTURA INADECUADA (TECHOS BAJOS, AREA REDUCIDA, SUPERFICIE DEFECTUOSA)</td>
      <td class="label-cell">Espacios confinados, obstáculos físicos y limitaciones estructurales que dificultan el trabajo seguro</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r28_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r28_aplica" value="no"></td>
      <td><input class="cell-input" name="r28_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r28_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">PAREDES, MUROS Y LOSAS EN MAL ESTADO</td>
      <td class="label-cell">Desprendimiento de estructuras</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r34_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r34_aplica" value="no"></td>
      <td><input class="cell-input" name="r34_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r34_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">TECHOS EN MAL ESTADO</td>
      <td class="label-cell">Filtraciones, desprendimiento de materiales o colapso</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r35_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r35_aplica" value="no"></td>
      <td><input class="cell-input" name="r35_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r35_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">PISOS EN MAL ESTADO</td>
      <td class="label-cell">Caída y golpes</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r36_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r36_aplica" value="no"></td>
      <td><input class="cell-input" name="r36_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r36_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">ANAQUELES Y ESTANTERIAS EN MAL ESTADO</td>
      <td class="label-cell">Colapso de Anaqueles y estanterías</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r37_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r37_aplica" value="no"></td>
      <td><input class="cell-input" name="r37_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r37_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">PUERTAS Y VENTANAS EN MAL ESTADO</td>
      <td class="label-cell">Posibles accidentes por fallas estructurales de puertas y ventanas</td>
      <td class="label-cell">Fractura / contusiones</td>
      <td class="center"><input type="radio" name="r38_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r38_aplica" value="no"></td>
      <td><input class="cell-input" name="r38_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r38_nivel"></td>
    </tr>

    <!-- ========== PSICOSOCIALES (2 filas) ========== -->
    <tr>
      <td class="tipo" rowspan="2">PSICOSOCIALES</td>
      <td class="label-cell">CARGA DE TRABAJO</td>
      <td class="label-cell">Estrés laboral</td>
      <td class="label-cell">Afecciones al sistema de respuesta fisiológica, cognitiva y motora</td>
      <td class="center"><input type="radio" name="r29_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r29_aplica" value="no"></td>
      <td><input class="cell-input" name="r29_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r29_nivel"></td>
    </tr>
    <tr>
      <td class="label-cell">HOSTIGAMIENTO LABORAL</td>
      <td class="label-cell">Estrés laboral</td>
      <td class="label-cell">Afecciones al sistema de respuesta fisiológica, cognitiva y motora</td>
      <td class="center"><input type="radio" name="r30_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r30_aplica" value="no"></td>
      <td><input class="cell-input" name="r30_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r30_nivel"></td>
    </tr>

    <!-- ========== FENOMENOS AMBIENTALES (1 fila) ========== -->
    <tr>
      <td class="tipo" rowspan="1">FENOMENOS<br>AMBIENTALES</td>
      <td class="label-cell">LLUVIA, TERREMOTOS, RAYOS, DESBORDE DE RÍOS</td>
      <td class="label-cell">Inundaciones, caída de objetos por derrumbe, descarga eléctrica</td>
      <td class="label-cell">Ahogamiento / Politraumatismos / Quemaduras / Muerte</td>
      <td class="center"><input type="radio" name="r31_aplica" value="si"></td>
      <td class="center"><input type="radio" name="r31_aplica" value="no"></td>
      <td><input class="cell-input" name="r31_obs"></td>
      <td>
        <input type="text" id="probabilidad-input" list="probabilidad-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="probabilidad-list">
                @foreach ($probabilidad as $p)
                    <option data-id="{{ $p->id_probabilidad }}" value="{{ $p->probabilidad }}"></option>
                @endforeach
            </datalist>
      </td>
      <td>
            <input type="text" id="consecuencia-input" list="consecuencia-list" class="cell-input" placeholder="Escribe para buscar…" autocomplete="off" style="font-size:10px;">
            <datalist id="consecuencia-list">
                @foreach ($consecuencia as $p)
                    <option data-id="{{ $p->id_consecuencia }}" value="{{ $p->consecuencia }}"></option>
                @endforeach
            </datalist>
      </td>
      <td><input class="cell-input" name="r31_nivel" style="font-size:10px;"></td>
    </tr>
  </tbody>
</table>

<style>
  /* Encabezados grises como el resto del formato */
  .risk-table thead th{
    background:#bfbfbf; font-weight:700; text-align:center; padding:6px;
  }
  .risk-table .tipo{
    text-transform:uppercase; font-weight:700; text-align:center; vertical-align:middle;
    background:#eee;
  }
  .risk-table input[type="radio"]{ transform:scale(1.1); }
</style>


<!-- ===== TABLA DE IDENTIFICACIÓN DE RIESGO (continuación) ===== -->
<style>
  .cell-select{ width:100%; border:0; padding:8px; font:inherit; background:transparent; outline:none; appearance:none; }
</style>

<!-- ===== FIRMAS Y FECHAS ===== -->
<table class="sheet" style="margin-top:8px;">
  <colgroup>
    <col style="width:20%">
    <col style="width:40%">
    <col style="width:10%">
    <col style="width:30%">
  </colgroup>
  <tbody>
    <tr>
      <td class="label-cell">Evaluación realizada por:</td>
      <td><input class="cell-input" name="eval_realizada_por"></td>
      <td class="label-cell center">Fecha</td>
      <td><input class="cell-input" type="date" name="eval_realizada_fecha"></td>
    </tr>
    <tr>
      <td class="label-cell">Evaluación revisada por:</td>
      <td><input class="cell-input" name="eval_revisada_por"></td>
      <td class="label-cell center">Fecha</td>
      <td><input class="cell-input" type="date" name="eval_revisada_fecha"></td>
    </tr>
    <tr>
      <td class="label-cell">Fecha próxima evaluación:</td>
      <td colspan="3"><input class="cell-input" type="date" name="fecha_proxima_evaluacion"></td>
    </tr>
  </tbody>
</table>

  <script>
  document.addEventListener('DOMContentLoaded', () => {
    // Reusa ensureHidden(tr, name) si ya la definiste arriba
    function _ensureHidden(container, name){
      let el = container.querySelector('input[type="hidden"][name="'+name+'"]');
      if (!el) { el = document.createElement('input'); el.type='hidden'; el.name=name; container.appendChild(el); }
      return el;
    }
    document.querySelectorAll('.risk-table tbody tr').forEach(tr => {
      const aplica = tr.querySelector('input[type="radio"][name^="r"][name$="_aplica"]');
      const obs    = tr.querySelector('input[name^="r"][name$="_obs"]');
      if (!aplica || !obs) return; // no es una fila de riesgo

      const m = aplica.name.match(/^r(\d+)_/);
      if (!m) return;
      const idx = m[1];

      // La primera label-cell de la fila es la columna PELIGRO
      const peligroCell = tr.querySelector('td.label-cell');
      const peligro = (peligroCell ? peligroCell.textContent : '').trim();

      _ensureHidden(tr, `riesgos[${idx}][nombre]`).value = peligro;

      // Normaliza datalist SI/NO/NA
const ALLOWED_SINO = new Set(['SI','NO','NA','']);
form.querySelectorAll('input[list="sino-na"]').forEach(inp => {
  const norm = () => {
    const v = (inp.value || '').trim().toUpperCase();
    inp.value = ALLOWED_SINO.has(v) ? v : '';
  };
  inp.addEventListener('change', norm);
  inp.addEventListener('blur', norm);
});

const ALLOWED_SINO = new Set(['SI','NO','NA','']);
  document.querySelectorAll('input[list="sino-na"]').forEach(inp => {
    const norm = () => { const v=(inp.value||'').trim().toUpperCase(); inp.value = ALLOWED_SINO.has(v)?v:''; };
    inp.addEventListener('change', norm);
    inp.addEventListener('blur', norm);
  });

    });
  });
</script>

  <div class="actions">
    <button class="btn-guardar" type="submit">Guardar</button>
  </div>
</form>

  </div>

<script>
document.addEventListener('DOMContentLoaded', () => {
  /* ===================== Referencias base ===================== */
  const form        = document.getElementById('riesgos-form');
  const fetchBase   = form?.dataset.fetchBase || '';
  const puestoInput = document.getElementById('puesto-input');
  const hiddenId    = document.getElementById('puesto_id');
  const sidebarRows = Array.from(document.querySelectorAll('.puestos-sidebar__row[data-puesto-id]'));
  const sidebarSearch = document.getElementById('puestos-search');
  const noResultsRow = document.querySelector('.puestos-sidebar__row--no-results');
  const hasSidebarRows = sidebarRows.length > 0;

  const deptoInput  = document.querySelector('input[name="departamento"]');
  const numInput    = document.querySelector('input[name="ptm_num_empleados"]');
  const descInput   = document.querySelector('textarea[name="ptm_descripcion_general"]');
  const actsInput   = document.querySelector('textarea[name="ptm_actividades_diarias"]');
  const objInput    = document.querySelector('textarea[name="ptm_objetivo_puesto"]');

  const puestosData = @json($puestos);
  const VALORACIONES = @json($valoracionTabla ?? []);
  const RIESGOS_CAT  = @json($riesgos ?? []); // {id_riesgo, nombre_riesgo}

  /* ===================== Utilidades generales ===================== */
  const markActiveRow = (id, { scroll = false } = {}) => {
    const targetId = String(id ?? '').trim();
    sidebarRows.forEach(row => {
      if (targetId && String(row.dataset.puestoId) === targetId) {
        row.classList.add('puestos-sidebar__row--active');
        if (scroll) {
          row.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      } else {
        row.classList.remove('puestos-sidebar__row--active');
      }
    });
  };

  sidebarRows.forEach(row => {
    row.addEventListener('click', () => {
      const id = row.dataset.puestoId;
      const match = (puestosData || []).find(p => String(p.id_puesto_trabajo_matriz) === String(id));
      if (match && puestoInput) {
        markActiveRow(id, { scroll: false });
        puestoInput.value = match.puesto_trabajo_matriz;
        puestoInput.dispatchEvent(new Event('change', { bubbles: true }));
      }
    });
  });

  const q  = sel => form.querySelector(sel);
  const qa = sel => Array.from(form.querySelectorAll(sel));
  const setVal    = (name, val) => { const el = q(`[name="${name}"]`); if (el) el.value = val ?? ''; };
  const setRadio  = (name, val) => qa(`input[type="radio"][name="${name}"]`).forEach(r => r.checked = (String(r.value) === String(val)));
  const setSelect = (name, val) => setVal(name, val);

  function ensureHidden(container, name){
    let el = container.querySelector(`input[type="hidden"][name="${name}"]`);
    if (!el) { el = document.createElement('input'); el.type='hidden'; el.name=name; container.appendChild(el); }
    return el;
  }

  const normalize = s => (s||'')
    .toString()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu,'')
    .toLowerCase()
    .trim();

  const filterSidebar = (term = '') => {
    const needle = normalize(term);
    let visibles = 0;
    sidebarRows.forEach(row => {
      const nombre = normalize(row.dataset.nombre || row.textContent || '');
      const match = !needle || nombre.includes(needle);
      row.style.display = match ? '' : 'none';
      if (match) visibles += 1;
    });
    if (noResultsRow) noResultsRow.style.display = (hasSidebarRows && !visibles) ? '' : 'none';
  };

  sidebarSearch?.addEventListener('input', e => {
    filterSidebar(e.target.value || '');
  });

  filterSidebar(sidebarSearch?.value || '');

  /* ===================== Autollenado por selección de puesto ===================== */
  const norm = v => (v||'').toString().trim().toLowerCase();
  function actualizarDesdeSeleccion() {
    const valor = puestoInput?.value || '';
    const match = (puestosData || []).find(p => norm(p.puesto_trabajo_matriz) === norm(valor));
    if (match) {
      deptoInput.value = match.departamento || '';
      numInput.value   = match.num_empleados ?? '';
      descInput.value  = match.descripcion_general || '';
      actsInput.value  = match.actividades_diarias || '';
      objInput.value   = match.objetivo_puesto || '';
      hiddenId.value   = match.id_puesto_trabajo_matriz;
    } else {
      deptoInput.value = '';
      numInput.value   = '';
      descInput.value  = '';
      actsInput.value  = '';
      objInput.value   = '';
      hiddenId.value   = '';
    }
  }

  /* ===================== Probabilidad/Consecuencia → Nivel ===================== */
  const NIVEL_COLOR_BY_ID = { '5':'#ff0000', '4':'#be5014', '3':'#ffc000', '2':'#ffff00', '1':'#92d050' };
  const NIVEL_COLOR_BY_LABEL = {
    'RIESGO MUY ALTO':'#ff0000','MUY ALTO':'#ff0000',
    'RIESGO ALTO':'#be5014','ALTO':'#be5014',
    'RIESGO MEDIO':'#ffc000','MEDIO':'#ffc000',
    'RIESGO BAJO':'#ffff00','BAJO':'#ffff00',
    'RIESGO IRRELEVANTE':'#92d050','IRRELEVANTE':'#92d050'
  };

  function getIdFromDatalist(listId, value){
    const dl = document.getElementById(listId);
    if (!dl) return null;
    const v = (value||'').trim().toLowerCase();
    const opt = Array.from(dl.options || []).find(o => (o.value||'').trim().toLowerCase() === v);
    return opt ? (opt.getAttribute('data-id') || null) : null;
  }

  function findNivelRecord(probId, consId){
    return (VALORACIONES || []).find(v =>
      String(v.id_probabilidad)===String(probId) && String(v.id_consecuencia)===String(consId)
    );
  }

  function nivelColor(nivelLabel, nivelId){
    if (nivelId && NIVEL_COLOR_BY_ID[String(nivelId)]) return NIVEL_COLOR_BY_ID[String(nivelId)];
    const key = (nivelLabel||'').toString().trim().toUpperCase();
    return NIVEL_COLOR_BY_LABEL[key] || '';
  }

  function updateRow(tr){
    const probInput  = tr.querySelector('input[list="probabilidad-list"]');
    const consInput  = tr.querySelector('input[list="consecuencia-list"]');
    const nivelInput = tr.querySelector('input[name$="_nivel"]');
    if (!probInput || !consInput || !nivelInput) return;

    const prefix = (nivelInput.name || '').replace(/_nivel$/, '');

    const probId = getIdFromDatalist('probabilidad-list', probInput.value);
    const consId = getIdFromDatalist('consecuencia-list', consInput.value);

    ensureHidden(tr, `${prefix}_id_probabilidad`).value = probId || '';
    ensureHidden(tr, `${prefix}_id_consecuencia`).value = consId || '';

    let nivelLabel = '', nivelId = '';
    if (probId && consId) {
      const rec = findNivelRecord(probId, consId);
      if (rec) { nivelLabel = rec.nivel_riesgo || ''; nivelId = rec.id_nivel_riesgo || ''; }
    }

    nivelInput.value = nivelLabel || (nivelId ? String(nivelId) : '');
    const color = nivelColor(nivelLabel, nivelId);
    if (color) {
      Object.assign(nivelInput.style, { backgroundColor: color, color:'#000', fontWeight:'700', textAlign:'center' });
    } else {
      Object.assign(nivelInput.style, { backgroundColor: '', color:'', fontWeight:'', textAlign:'' });
    }

    ensureHidden(tr, `${prefix}_id_nivel_riesgo`).value = nivelId || '';
  }

  // Enlazar cálculo a todas las filas con prob/cons/nivel
  qa('table.sheet tbody tr').forEach(tr => {
    const p = tr.querySelector('input[list="probabilidad-list"]');
    const c = tr.querySelector('input[list="consecuencia-list"]');
    const n = tr.querySelector('input[name$="_nivel"]');
    if (p && c && n) {
      ['input','change'].forEach(evt => {
        p.addEventListener(evt, () => updateRow(tr));
        c.addEventListener(evt, () => updateRow(tr));
      });
      updateRow(tr); // inicial
    }
  });

  /* ===================== Mapear ID de riesgo por catálogo ===================== */
  const ID_BY_NOMBRE = Object.fromEntries(
    (RIESGOS_CAT || []).map(r => [ normalize(r.nombre_riesgo || r.peligro || ''), r.id_riesgo ])
  );

  function generateRiesgoHiddenInputs(){
    document.querySelectorAll('table.risk-table tbody tr').forEach(tr => {
      const any = tr.querySelector('input[name^="r"][name$="_obs"], input[name^="r"][name$="_nivel"], input[name^="r"][name$="_aplica"]');
      if (!any) return;
      const m = any.name.match(/^r(\d+)_/); if (!m) return;
      const idx = m[1];

      // 1ª label-cell de la fila = PELIGRO (en tu tabla)
      const peligroCell = tr.querySelector('td.label-cell');
      const peligroTxt  = (peligroCell?.textContent || '').trim();
      const idRiesgo    = ID_BY_NOMBRE[ normalize(peligroTxt) ] || '';

      ensureHidden(tr, `riesgos[${idx}][id_riesgo]`).value = idRiesgo;
      ensureHidden(tr, `riesgos[${idx}][nombre]`).value    = peligroTxt;
    });
  }
  generateRiesgoHiddenInputs();

  /* ===================== Relleno/Hidratación de secciones ===================== */
  function valueFromDataId(listId, dataId){
    const dl = document.getElementById(listId);
    if (!dl) return '';
    const opt = Array.from(dl.options||[]).find(o => String(o.getAttribute('data-id')) === String(dataId));
    return opt ? (opt.value || '') : '';
  }

  function fillQuimicos(rows){
    const tabla = document.getElementById('tabla-quimicos');
    const tbody = tabla?.querySelector('tbody');
    if (!tbody) return;
    // limpia dejando 1 fila
    qa('#tabla-quimicos tbody tr').slice(1).forEach(tr => tr.remove());
    const base = tbody.querySelector('tr') || null;
    if (!base) return;
    base.querySelectorAll('input').forEach(i => i.value = '');

    if (!rows || !rows.length) return;

    // primera
    base.querySelector('.quimico-id').value = rows[0].id_quimico ?? '';
    base.querySelector('.quimico-nombre').value = rows[0].nombre ?? '';
    q(`[name="quimicos[0][capacitacion]"]`).value = rows[0].capacitacion ?? '';
    q(`[name="quimicos[0][duracion]"]`).value = rows[0].duracion ?? '';
    q(`[name="quimicos[0][frecuencia]"]`).value = rows[0].frecuencia ?? '';
    q(`[name="quimicos[0][epp]"]`).value = rows[0].epp ?? '';

    // resto
    for (let i=1;i<rows.length;i++){
      const clon = base.cloneNode(true);
      tbody.appendChild(clon);
      clon.querySelectorAll('input').forEach(inp => inp.value = '');
      clon.querySelector('.quimico-id').name = `quimicos[${i}][id_quimico]`;
      clon.querySelector('.quimico-nombre').name = `quimicos[${i}][nombre]`;
      clon.querySelector('.quimico-id').value = rows[i].id_quimico ?? '';
      clon.querySelector('.quimico-nombre').value = rows[i].nombre ?? '';
      clon.querySelector('[name^="quimicos["][name$="[capacitacion]"]').name = `quimicos[${i}][capacitacion]`;
      clon.querySelector('[name^="quimicos["][name$="[duracion]"]').name = `quimicos[${i}][duracion]`;
      clon.querySelector('[name^="quimicos["][name$="[frecuencia]"]').name = `quimicos[${i}][frecuencia]`;
      clon.querySelector('[name^="quimicos["][name$="[epp]"]').name = `quimicos[${i}][epp]`;
      clon.querySelector(`[name="quimicos[${i}][capacitacion]"]`).value = rows[i].capacitacion ?? '';
      clon.querySelector(`[name="quimicos[${i}][duracion]"]`).value = rows[i].duracion ?? '';
      clon.querySelector(`[name="quimicos[${i}][frecuencia]"]`).value = rows[i].frecuencia ?? '';
      clon.querySelector(`[name="quimicos[${i}][epp]"]`).value = rows[i].epp ?? '';
    }
  }

  const DYN_TABLE_CONFIG = {
    visual: {
      tableId: 'tabla-visual',
      fields: [
        { selector: '.visual-tipo', key: 'tipo', trigger: 'input change' },
        { selector: 'input[name$="[tiempo]"]', key: 'tiempo' }
      ]
    },
    ruido: {
      tableId: 'tabla-ruido',
      fields: [
        { selector: '.ruido-desc', key: 'desc', trigger: 'input change' },
        { selector: 'input[name$="[duracion]"]', key: 'duracion' },
        { selector: 'input[name$="[epp]"]', key: 'epp' }
      ]
    },
    termico: {
      tableId: 'tabla-termico',
      fields: [
        { selector: '.termico-desc', key: 'desc', trigger: 'input change' },
        { selector: 'input[name$="[duracion]"]', key: 'duracion' },
        { selector: 'input[name$="[epp]"]', key: 'epp' }
      ]
    }
  };

  function fillDynamicRows(config, rows){
    const tabla = document.getElementById(config.tableId);
    if (!tabla) return;
    const tbody = tabla.querySelector('tbody');
    if (!tbody) return;
    const base = tbody.querySelector('tr');
    if (!base) return;

    const cleanupRow = (tr) => {
      tr.querySelectorAll('input, textarea, select').forEach(el => {
        if (el.type === 'checkbox' || el.type === 'radio') el.checked = false;
        else el.value = '';
      });
    };

    const ensureIndex = (tr, idx) => {
      tr.querySelectorAll('[name]').forEach(el => {
        const name = el.getAttribute('name');
        if (!name) return;
        el.setAttribute('name', name.replace(/\[\d+\]/, '[' + idx + ']'));
      });
    };

    [...tbody.querySelectorAll('tr')].slice(1).forEach(tr => tr.remove());
    cleanupRow(base);
    ensureIndex(base, 0);

    const baseBtn = base.querySelector('.icon-btn');
    if (baseBtn){
      baseBtn.textContent = '+';
      baseBtn.classList.add('add');
      baseBtn.classList.remove('remove');
    }

    if (!rows || !rows.length) return;

    rows.forEach((row, idx) => {
      let tr = idx === 0 ? base : null;
      if (idx > 0){
        const addBtn = tbody.querySelector('tr:last-child .icon-btn.add');
        if (addBtn) addBtn.click();
        tr = tbody.querySelector('tr:last-child');
        if (!tr) return;
      }

      cleanupRow(tr);
      ensureIndex(tr, idx);

      (config.fields || []).forEach(field => {
        const el = tr.querySelector(field.selector);
        if (!el) return;
        const val = field.value ? field.value(row, idx) : (row[field.key] ?? '');
        if (el.type === 'checkbox' || el.type === 'radio') {
          el.checked = !!val;
        } else {
          el.value = val ?? '';
        }
        if (field.trigger){
          field.trigger.split(' ').forEach(evt => {
            el.dispatchEvent(new Event(evt, { bubbles: true }));
          });
        }
      });
    });
  }

  const fillVisualRows  = rows => fillDynamicRows(DYN_TABLE_CONFIG.visual, rows);
  const fillRuidoRows   = rows => fillDynamicRows(DYN_TABLE_CONFIG.ruido, rows);
  const fillTermicoRows = rows => fillDynamicRows(DYN_TABLE_CONFIG.termico, rows);
  function fillEstadoObs(prefix, arr){
    if (!arr) return;
    arr.forEach((row, i) => {
      if (row?.estado) setRadio(`${prefix}[${i}][estado]`, row.estado);
      setVal(`${prefix}[${i}][obs]`, row?.obs ?? '');
    });
  }

  function fillErgo(ergo){
    if (!ergo) return;
    ergo.forEach((row, i) => {
      if (row?.resp) setRadio(`ergo[${i}][resp]`, row.resp);
      setVal(`ergo[${i}][obs]`, row?.obs ?? '');
    });
  }

  function fillCheckMatrix(map, groupSelectorName){
    if (!map) return;
    Object.entries(map).forEach(([k,v]) => {
      const el = q(`input[type="checkbox"][name="${groupSelectorName}"][value="${k}"]`);
      if (el) el.checked = !!v;
    });
  }

  function fillFuego(f){
    if (!f) return;
    Object.entries(f).forEach(([k,v]) => setSelect(`riesgo_fuego[${k}]`, v || ''));
  }


  // Normaliza respuestas triestado (SI/NO/NA) provenientes de la BD
  const normalizeTri = (v) => {
    if (typeof v === 'string') {
      const upper = v.trim().toUpperCase();
      const withoutAccents = upper.replace(/\u00cd/g, 'I');
      const collapsed = withoutAccents.replace(/[.\-\/\s]/g, '');
      if (collapsed === 'SI' || collapsed === 'TRUE' || collapsed === 'YES' || collapsed === 'S') return 'SI';
      if (collapsed === 'NO' || collapsed === 'FALSE' || collapsed === 'N') return 'NO';
      if (collapsed === '1') return 'SI';
      if (collapsed === '0') return 'NO';
      if (collapsed === 'NA' || collapsed === 'NOAPLICA') return 'NA';
      return '';
    }
    if (v === 1 || v === true)  return 'SI';
    if (v === 0 || v === false) return 'NO';
    return ''; // << dejar vacío cuando no haya dato
  };

function fillElec(sel, verif){
  if (sel) {
    Object.entries(sel).forEach(([k,v]) => setSelect(k, v || '')); // mantiene los <select> SI/NO/NA ya existentes
  }
  if (verif) {
    Object.entries(verif).forEach(([k,v]) => {
      setVal(`elec_verif[${k}]`, normalizeTri(v));
    });
  }
}

  function fillSimpleBlocks(payload){
    if (payload.visual){
      setVal('visual_tipo', payload.visual.tipo || '');
      setVal('visual_tiempo', payload.visual.tiempo || '');
    }
    if (payload.ruido){
      setVal('ruido_desc', payload.ruido.desc || '');
      setVal('ruido_duracion', payload.ruido.tiempo || '');
    }
    if (payload.termico){
      setVal('termico_desc', payload.termico.desc || '');
      setVal('termico_duracion', payload.termico.tiempo || '');
    }
    const F = payload.fisico || {};
    const putFis = (k, obj) => {
      setVal(`fisico_${k}_desc`, obj?.desc||'');
      setVal(`fisico_${k}_equipo`, obj?.equipo||'');
      setVal(`fisico_${k}_duracion`, obj?.duracion||'');
      setVal(`fisico_${k}_distancia`, obj?.distancia||'');
      setVal(`fisico_${k}_epp`, obj?.epp||'');
      setVal(`fisico_${k}_frecuencia`, obj?.frecuencia||'');
      setVal(`fisico_${k}_peso`, obj?.peso||'');
      setVal(`fisico_${k}_capacitacion`, obj?.capacitacion||'');
    };
    ['cargar','halar','empujar','sujetar'].forEach(k => putFis(k, F[k]));

    if (payload.alturas){
      setVal('alturas[altura]', payload.alturas.altura || '');
      setVal('alturas[anclaje_seguro]', payload.alturas.anclaje_seguro || '');
      setVal('alturas[inspeccion]', payload.alturas.inspeccion || '');
      setVal('alturas[epp]', payload.alturas.epp || '');
      setVal('alturas[senalizacion]', payload.alturas.senalizacion || '');
      setVal('alturas[capacitacion]', payload.alturas.capacitacion || '');
      setVal('alturas[aviso_trabajo_altura]', payload.alturas.aviso_trabajo_altura || '');
      setVal('alturas[firma_trabajo_seguro]', payload.alturas.firma_trabajo_seguro || '');
    }

    if (payload.otros){
      setVal('otros_biologico', payload.otros.biologico || '');
      setVal('otros_psicosocial', payload.otros.psicosocial || '');
      setVal('otros_naturales', payload.otros.naturales || '');
    }

    if (payload.firmas){
      setVal('eval_realizada_por', payload.firmas.eval_realizada_por || '');
      setVal('eval_realizada_fecha', payload.firmas.eval_realizada_fecha || '');
      setVal('eval_revisada_por', payload.firmas.eval_revisada_por || '');
      setVal('eval_revisada_fecha', payload.firmas.eval_revisada_fecha || '');
      setVal('fecha_proxima_evaluacion', payload.firmas.fecha_proxima_evaluacion || '');
    }
  }

  function fillRiesgos(riesgoValor, evalMap){
    qa('table.risk-table tbody tr').forEach(tr => {
      const any = tr.querySelector('input[name^="r"][name$="_obs"], input[name^="r"][name$="_nivel"], input[name^="r"][name$="_aplica"]');
      if (!any) return;
      const m = any.name.match(/^r(\d+)_/); if (!m) return;
      const i = m[1];
      const hid = tr.querySelector(`input[type="hidden"][name="riesgos[${i}][id_riesgo]"]`);
      let rid = hid?.value ? parseInt(hid.value, 10) : null;

      if (!rid && riesgoValor){
        const peligroCell = tr.querySelector('td.label-cell');
        const peligroTxt  = (peligroCell?.textContent || '').trim();
        const peligroN    = normalize(peligroTxt);
        for (const [k,v] of Object.entries(riesgoValor)){
          if (normalize(v?.nombre || '') === peligroN) { rid = parseInt(k,10); break; }
        }
      }
      if (!rid) return;

      const rv = riesgoValor && riesgoValor[rid] ? riesgoValor[rid] : null;
      if (rv && rv.valor) setRadio(`r${i}_aplica`, rv.valor);
      if (rv && (rv.obs ?? '') !== '') setVal(`r${i}_obs`, rv.obs);

      const er = evalMap && evalMap[rid] ? evalMap[rid] : null;
      if (er){
        const probInp = tr.querySelector('input[list="probabilidad-list"]');
        const consInp = tr.querySelector('input[list="consecuencia-list"]');

        if (probInp) probInp.value = er.prob_label || valueFromDataId('probabilidad-list', er.id_probabilidad);
        if (consInp) consInp.value = er.cons_label || valueFromDataId('consecuencia-list', er.id_consecuencia);

        updateRow(tr); // recalcula ids ocultos + color de nivel
      }
    });
  }

  /* ===================== Limpiar dinámicos (nuevo) ===================== */
// --- Limpieza completa (reemplaza tu clearDynamic) ---
function clearDynamic(){
  // radios y checks
  qa('input[type="radio"]').forEach(r => r.checked = false);
  qa('input[type="checkbox"]').forEach(c => c.checked = false);

  // campos que NO se deben limpiar (maestros del puesto)
  const skipNames = new Set([
    'ptm_num_empleados',
    'ptm_descripcion_general',
    'ptm_actividades_diarias',
    'ptm_objetivo_puesto',
    'departamento'
  ]);
  const shouldSkip = (el) => {
    if (el === puestoInput || el === hiddenId) return true;
    const nm = el.getAttribute('name') || '';
    return skipNames.has(nm);
  };

  // === limpiar TODO lo que tenga class="cell-input" ===
  qa('.cell-input').forEach(el => {
    if (shouldSkip(el)) return;

    if (el.matches('select')) {
      // reset de selects: opción vacía si existe, o ninguna seleccionada
      if ([...el.options].some(o => o.value === '')) el.value = '';
      else el.selectedIndex = -1;
    } else {
      // inputs/textareas (incluye los que usan datalist)
      el.value = '';
    }
  });

  // limpiar estilos de los campos *_nivel
  qa('input[name$="_nivel"]').forEach(n => {
    n.style.backgroundColor = '';
    n.style.color = '';
    n.style.fontWeight = '';
    n.style.textAlign = '';
  });

  // tabla de químicos: deja una sola fila vacía
  const tbody = document.querySelector('#tabla-quimicos tbody');
  if (tbody){
    [...tbody.querySelectorAll('tr')].slice(1).forEach(tr => tr.remove());
    const base = tbody.querySelector('tr');
    if (base) base.querySelectorAll('input').forEach(i => i.value = '');
  }

  // tablas dinámicas: quita filas extra y limpia la base
['#tabla-quimicos', '#tabla-visual', '#tabla-ruido', '#tabla-termico'].forEach(sel => {
  const tbody = document.querySelector(`${sel} tbody`);
  if (!tbody) return;
  [...tbody.querySelectorAll('tr')].slice(1).forEach(tr => tr.remove());
  const base = tbody.querySelector('tr');
  if (base){
    base.querySelectorAll('input, textarea, select').forEach(i => {
      if (i.type === 'hidden') i.value = '';
      else i.value = '';
      if (i.type === 'checkbox' || i.type === 'radio') i.checked = false;
    });
  }
});
}



  /* ===================== Hidratar por puesto (GET /{id}/datos) ===================== */
// --- HIDRATACIÓN (reemplaza hydrateByPuestoId) ---
async function hydrateByPuestoId(id){
  if (!fetchBase) return;

  // convierte 0/1/true/false o texto -> "SI/NO/NA" o "" (vacío si no hay dato)
  const triFrom = normalizeTri;

  const fillCaidaLocal = (src) => {
    const c = src || {};
    ['pisos_adecuados','vias_libres','rampas_identificadas',
     'gradas_barandas','antideslizantes','senalizacion_piso_resbaloso']
    .forEach(k => setVal(`caida[${k}]`, triFrom(c[k])));
    setVal('caida_observaciones', c.obs || '');
  };

  const fillElecLocal = (sel, verif) => {
    if (sel) Object.entries(sel).forEach(([k,v]) => setSelect(k, v || ''));
    if (verif) Object.entries(verif).forEach(([k,v]) => {
      setVal(`elec_verif[${k}]`, triFrom(v));
    });
  };

  // Limpia primero para que no queden residuos visuales
  clearDynamic();

  if (!id) { fillQuimicos([]); return; }

  try{
    const res = await fetch(`${fetchBase}/${id}`, { headers: { 'Accept':'application/json' }});
    if (!res.ok) { fillQuimicos([]); return; }

    const data = await res.json();

    // Datos maestros del puesto (estos sí se rellenan aunque no haya ident)
    setVal('ptm_num_empleados',        data?.puesto?.ptm_num_empleados ?? '');
    setVal('ptm_descripcion_general',  data?.puesto?.ptm_descripcion_general ?? '');
    setVal('ptm_actividades_diarias',  data?.puesto?.ptm_actividades_diarias ?? '');
    setVal('ptm_objetivo_puesto',      data?.puesto?.ptm_objetivo_puesto ?? '');
    if (deptoInput) deptoInput.value = data?.puesto?.departamento ?? '';

    // Si no hay identificación guardada, dejamos todo limpio
    if (!data.ident){
      fillVisualRows([]);
      fillRuidoRows([]);
      fillTermicoRows([]);
      fillQuimicos([]);
      return;
    }

    // Rellenos
    fillSimpleBlocks(data.ident);
    fillEstadoObs('instalaciones', data.ident.instalaciones);
    fillEstadoObs('maq',           data.ident.maq);
    fillEstadoObs('emer',          data.ident.emer);
    fillErgo(data.ident.ergo);
    fillCheckMatrix(data.ident.posturas, 'posturas[]');
    fillFuego(data.ident.fuego);

    // Eléctrico verificación (nuevo) + compat con el viejo elec_chk
    fillElecLocal(data.ident.elec_select, data.ident.elec_verif || data.ident.elec_chk);

    // Caída (nuevo) + compat con el viejo caida_nivel
    fillCaidaLocal(data.ident.caida || data.ident.caida_nivel || null);

    fillVisualRows(data.visual_rows || []);
    fillRuidoRows(data.ruido_rows || []);
    fillTermicoRows(data.termico_rows || []);

    // Químicos / Riesgos
    fillQuimicos(data.quimicos || []);
    fillRiesgos(data.riesgo_valor || {}, data.evaluacion_riesgo || {});
  } catch(e){
    console.error(e);
    // queda limpio por el clearDynamic inicial
  }
}



  /* ===================== Cambio de puesto (nuevo) ===================== */
  function onPuestoSelected(){
    // Actualiza maestros (depto, desc, etc.) y el hiddenId
    if (typeof actualizarDesdeSeleccion === 'function') actualizarDesdeSeleccion();

    const id = (hiddenId?.value || '').trim();
    markActiveRow(id);
    if (id){
      hydrateByPuestoId(id); // trae datos si existen; si no, limpia
    } else {
      clearDynamic();        // nuevo/no en lista → empezar desde cero
    }
  }

  /* ===================== Listeners / Inicialización ===================== */
  puestoInput?.addEventListener('input',  onPuestoSelected);
  puestoInput?.addEventListener('change', onPuestoSelected);

  // precarga por ID si viene del servidor
  const initialId = "{{ $id_puesto_trabajo_matriz ?? '' }}";
  if (initialId && puestosData?.length) {
    const m = puestosData.find(p => String(p.id_puesto_trabajo_matriz) === String(initialId));
    if (m && puestoInput) puestoInput.value = m.puesto_trabajo_matriz;
  }
  actualizarDesdeSeleccion();
  onPuestoSelected(); // aplica limpieza/hidratación inicial

  // Validación submit: forzar elegir un puesto válido
  form?.addEventListener('submit', e => {
    if (!hiddenId?.value) {
      e.preventDefault();
      alert('Selecciona un puesto válido de la lista.');
      puestoInput?.focus();
    }
  });
});
</script>
</body>
</html>
