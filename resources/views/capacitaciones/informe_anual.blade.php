@extends('layouts.capacitacion')
@section('title', 'Informe anual de Capacitaciones')

@section('content')
<style>
  :root{ --brand:#00B0F0; }

  .btn{padding:.5rem .75rem;border-radius:12px;border:1px solid #e5e7eb;background:#f9fafb;cursor:pointer}
  .btn-primary{background:#00B0F0;color:#fff;border-color:#00B0F0}
  .btn-outline{background:#fff;color:#111827}
  .badge{display:inline-block;padding:.1rem .45rem;border-radius:.5rem;font-weight:700;font-size:.72rem}
  .b-ok{background:#dcfce7;color:#166534}
  .b-warn{background:#fee2e2;color:#991b1b}
  .b-info{background:#e0f2fe;color:#075985}

  .table{width:100%;border-collapse:separate;border-spacing:0}
  .table th,.table td{border-bottom:1px solid #e5e7eb;padding:.55rem .6rem;font-size:.875rem;vertical-align:top}
  .table th{position:sticky;top:0;background:#f8fafc;z-index:1}

  .chip{display:inline-block;padding:.15rem .45rem;border-radius:.5rem;background:#f1f5f9;margin:.1rem .15rem 0 0}

  /* Mapa de colores (pintar TD para forzar fill completo) */
  .row-prog-done > td{
    background:#ecfdf5 !important;  /* verde suave */
    border-bottom-color:#bbf7d0;
  }
  .row-prog-done:hover > td{ background:#d1fae5 !important; }

  .row-prog-pending > td{
    background:#fffbeb !important;  /* amarillo suave */
    border-bottom-color:#fde68a;
  }
  .row-prog-pending:hover > td{ background:#fef3c7 !important; }

  @media print{
    .row-prog-done > td,
    .row-prog-pending > td,
    .badge,.chip{
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
  }
</style>

<div class="p-6 space-y-4">
  <div class="flex items-center justify-between flex-wrap gap-3">
    <div>
      <h1 class="text-2xl font-bold">Informe anual de Capacitaciones</h1>
      <p class="text-gray-600">Año <span class="font-semibold">{{ $year }}</span></p>
    </div>
  </div>

  <form method="GET" action="{{ route('capacitaciones.informe-anual') }}"
        class="flex items-end gap-3 flex-wrap bg-white p-3 rounded-xl border border-gray-200">
    <div>
      <label class="block text-sm font-medium text-gray-700">Año</label>
      <select name="anio" class="mt-1 rounded-xl border-gray-300 focus:border-brand focus:ring-0">
        @foreach($years as $y)
          <option value="{{ $y }}" @selected((int)$y === (int)$year)>{{ $y }}</option>
        @endforeach
      </select>
    </div>
    <div class="grow min-w-[260px]">
      <label class="block text-sm font-medium text-gray-700">Buscar capacitación</label>
      <input type="text" name="q" value="{{ $q }}" placeholder="Ej: Inducción, EPP, Ergonomía…"
             class="mt-1 w-full rounded-xl border-gray-300 focus:border-brand focus:ring-0" />
    </div>
    <div class="flex gap-2">
      <button class="btn btn-primary">Filtrar</button>
      <a class="btn" style="background:#64748b;border-color:#64748b;color:#fff"
    href="{{ route('capacitaciones.informe-anual.export', request()->only('anio','q')) }}">
    Descargar Excel
    </a>
      <a class="btn btn-outline" href="{{ route('capacitaciones.informe-anual') }}">Limpiar</a>
    </div>
  </form>

  <!-- KPI: % Ejecutadas vs No Ejecutadas (solo programadas) -->
  <div class="grid sm:grid-cols-3 gap-3">
    <div class="p-4 rounded-xl border border-gray-200 bg-white">
      <div class="text-xs text-gray-500">Programadas</div>
      <div class="text-2xl font-bold">{{ $kpi['total'] }}</div>
      <div class="text-xs text-gray-500">Totales</div>
    </div>
    <div class="p-4 rounded-xl border border-gray-200 bg-white">
      <div class="text-xs text-gray-500">% EJECUTADAS</div>
      <div class="text-2xl font-bold text-emerald-700">{{ $kpi['pct_ejec'] }}%</div>
      <div class="text-xs text-gray-500">{{ $kpi['ejecutadas'] }} de {{ $kpi['total'] }}</div>
    </div>
    <div class="p-4 rounded-xl border border-gray-200 bg-white">
      <div class="text-xs text-gray-500">% NO EJECUTADAS</div>
      <div class="text-2xl font-bold text-amber-700">{{ $kpi['pct_no_ejec'] }}%</div>
      <div class="text-xs text-gray-500">{{ $kpi['no_ejec'] }} de {{ $kpi['total'] }}</div>
    </div>
  </div>

  <!-- Mapa de colores -->
  <div class="flex items-center gap-4 text-sm text-gray-700">
    <div class="flex items-center gap-2">
      <span class="inline-block w-4 h-4 rounded border" style="background:#ecfdf5;border-color:#bbf7d0;"></span>
      <span>Programada <strong>impartida</strong></span>
    </div>
    <div class="flex items-center gap-2">
      <span class="inline-block w-4 h-4 rounded border" style="background:#fffbeb;border-color:#fde68a;"></span>
      <span>Programada <strong>no impartida</strong></span>
    </div>
  </div>

  <div class="overflow-auto rounded-xl border border-gray-200">
    <table class="table min-w-[1100px]">
      <thead>
        <tr>
          <th style="min-width:280px;">Tema</th>
          <th>Categoría</th>
          <th># Capacitaciones</th>
          <th>Horas Capacitaciones</th>
          <th>% del total de sesiones</th>
          <th>% del total de horas</th>
          <th># de personas</th>
          <th>Horas-hombre</th>
        </tr>
      </thead>
      <tbody>
        @forelse($rows as $r)
        @php
        $rowClass = ($r->programada === 'SI' && (int)$r->sesiones > 0)
            ? 'row-prog-done'
            : ( ($r->programada === 'SI' && (int)$r->sesiones === 0)
                ? 'row-prog-pending'
                : ''
            );
        @endphp
          <tr class="{{ $rowClass }}">
            <td>
              <div class="font-semibold">{{ $r->tema }}</div>
              @if($r->programada === 'SI')
                <div class="text-xs text-gray-500">Programada</div>
              @endif
            </td>
            <td class="text-center">{{ $r->categoria ?? '—' }}</td>
            <td class="text-center">{{ (int)$r->sesiones }}</td>
            <td class="text-center">{{ number_format($r->horas_cap, 2) }}</td>
            <td class="text-center">{{ number_format($r->pct_sesiones, 2) }}%</td>
            <td class="text-center">{{ number_format($r->pct_horas, 2) }}%</td>
            <td class="text-center">{{ (int)$r->personas }}</td>
            <td class="text-center font-semibold">{{ number_format($r->horas_hombre, 2) }}</td>
          </tr>
        @empty
          <tr><td colspan="8" class="text-center text-gray-500 py-6">Sin resultados para el año seleccionado.</td></tr>
        @endforelse
      </tbody>
      <tfoot>
        <tr>
          <th class="text-right">Totales:</th>
          <th></th>
          <th class="text-center">{{ $tot['sesiones'] }}</th>
          <th class="text-center">{{ number_format($tot['horas_cap'], 2) }}</th>
          <th class="text-center">100%</th>
          <th class="text-center">100%</th>
          <th class="text-center">{{ $tot['personas'] }}</th>
          <th class="text-center">{{ number_format($tot['horas_hombre'], 2) }}</th>
        </tr>
      </tfoot>
    </table>
  </div>

  <p class="text-sm text-gray-500">
    Notas:
    1) <em># Capacitaciones</em> cuenta fechas distintas impartidas en el año. 
    2) <em>Horas Capacitaciones</em> = (duración en horas) × (# Capacitaciones). 
    3) <em>Horas-hombre</em> = (filas de asistencia) × (duración en horas).
    4) <em># de personas</em> son empleados distintos que recibieron esa capacitación durante el año.
  </p>
</div>
@endsection
