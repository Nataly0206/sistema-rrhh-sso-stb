@extends('layouts.capacitacion')
@section('title', 'Informe mensual de Capacitaciones')

@section('content')
<style>
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

    :root{ --brand:#00B0F0; }

    /* Fila destacada para programadas (VERDE) */
  .row-prog{ background:#ecfdf5; }               /* verde muy suave (emerald-50) */
  .row-prog:hover{ background:#d1fae5; }         /* hover un poco más intenso */
  .row-prog td{ border-bottom-color:#bbf7d0; }   /* línea inferior en verde (emerald-200) */
  .tema-prog{ color:#065f46; }                   /* título en verde oscuro (emerald-900) */
  .pill-prog{
    display:inline-block; padding:.1rem .45rem; border-radius:.5rem;
    background:#bbf7d0; color:#065f46; font-weight:700; font-size:.72rem; margin-left:.4rem;
  }


    /* Fila programada en VERDE: pinta las celdas, no solo el <tr> */
.table tbody tr.row-prog > td{
  background:#ecfdf5 !important;   /* emerald-50 */
  border-bottom-color:#bbf7d0;      /* emerald-200 */
}
.table tbody tr.row-prog:hover > td{
  background:#d1fae5 !important;    /* hover un poco más intenso */
}

/* Color del título y la "pill" */
.tema-prog{ color:#065f46; } /* emerald-900 */
.pill-prog{
  display:inline-block; padding:.1rem .45rem; border-radius:.5rem;
  background:#bbf7d0; color:#065f46; font-weight:700; font-size:.72rem; margin-left:.4rem;
}

/* Mantener colores al imprimir */
@media print{
  .table tbody tr.row-prog > td,
  .pill-prog, .badge, .chip{
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  }
</style>

<div class="p-6 space-y-4">
  <div class="flex items-center justify-between flex-wrap gap-3">
    <div>
      <h1 class="text-2xl font-bold">Informe mensual de Capacitaciones</h1>
      <p class="text-gray-600">Del <span class="font-semibold">{{ $start }}</span> al <span class="font-semibold">{{ $end }}</span></p>
    </div>
  </div>

  <form method="GET" action="{{ route('capacitaciones.informe-mensual') }}"
        class="flex items-end gap-3 flex-wrap bg-white p-3 rounded-xl border border-gray-200">
    <div>
      <label class="block text-sm font-medium text-gray-700">Mes</label>
      <select name="mes" class="mt-1 rounded-xl border-gray-300 focus:border-brand focus:ring-0">
        @foreach($months as $mVal=>$mName)
          <option value="{{ $mVal }}" @selected((int)$mVal === (int)$month)>{{ $mName }}</option>
        @endforeach
      </select>
    </div>
    <div>
      <label class="block text-sm font-medium text-gray-700">Año</label>
      <select name="anio" class="mt-1 rounded-xl border-gray-300 focus:border-brand focus:ring-0">
        @foreach($years as $y)
          <option value="{{ $y }}" @selected((int)$y === (int)$year)>{{ $y }}</option>
        @endforeach
      </select>
    </div>
    <div class="grow min-w-[220px]">
      <label class="block text-sm font-medium text-gray-700">Buscar capacitación</label>
      <input type="text" name="q" value="{{ $q }}" placeholder="Ej: Generalidades de la Empresa"
             class="mt-1 w-full rounded-xl border-gray-300 focus:border-brand focus:ring-0" />
    </div>
    <div class="flex gap-2">
      <button class="btn btn-primary">Filtrar</button>
      <a class="btn" style="background:#64748b;border-color:#64748b;color:#fff"
        href="{{ route('capacitaciones.informe-mensual.export', request()->only('anio','mes','q')) }}">
        Descargar Excel
        </a>
      <a class="btn btn-outline" href="{{ route('capacitaciones.informe-mensual') }}">Limpiar</a>
    </div>
  </form>

  <div class="overflow-auto rounded-xl border border-gray-200">
    <table class="table min-w-[1200px]">
      <thead>
        <tr>
          <th style="width:50px;">N°</th>
          <th style="min-width:260px;">Tema</th>
          <th>Participantes</th>
          <th>Departamento(s)</th>
          <th style="min-width:220px;">Fecha(s) del mes</th>
          <th>Duración</th>
          <th>Categoría</th>
          <th>No. Horas-Hombre</th>
          <th>Número (Veces Impartidas)</th>
          <th>Programada</th>
        </tr>
      </thead>
        <tbody>
    @php $i = 1; @endphp
    @forelse($rows as $r)
        @php $isProg = strtoupper(trim($r->programada ?? '')) === 'SI'; @endphp
        <tr class="{{ $isProg ? 'row-prog' : '' }}"><!-- << NUEVO: color por fila si es programada -->
        <td class="text-center text-gray-700">{{ $i++ }}</td>
        <td>
            <div class="font-semibold {{ $isProg ? 'tema-prog' : '' }}">
            {{ $r->tema }}
            @if($isProg)
                <span class="pill-prog">Programada</span> <!-- << NUEVO: pill junto al tema -->
            @endif
            </div>
        </td>
        <td class="text-center">
            <span class="badge b-info">{{ $r->participantes }}</span>
        </td>
        <td>
        @php $deps = preg_split('/\s*,\s*/', $r->departamentos ?? '', -1, PREG_SPLIT_NO_EMPTY); @endphp
        @if(!empty($deps))
            @foreach($deps as $dep)
            <span class="chip">{{ $dep }}</span>
            @endforeach
        @else
            <span class="text-gray-400">–</span>
        @endif
        </td>

        <td>
        @php $fechs = preg_split('/\s*,\s*/', $r->fechas ?? '', -1, PREG_SPLIT_NO_EMPTY); @endphp
        @if(!empty($fechs))
            @foreach($fechs as $f)
            <span class="chip">{{ $f }}</span>
            @endforeach
        @else
            <span class="text-gray-400">–</span>
        @endif
        </td>
        <td class="text-center">{{ (int)$r->duracion }} min</td>
        <td class="text-center">{{ $r->categoria ?? '—' }}</td>
        <td class="text-center font-semibold">{{ (int)$r->horas_hombre }}</td>
        <td class="text-center">{{ (int)$r->numero }}</td>
        <td class="text-center">
            @if($isProg)
            <span class="badge b-ok">SI</span>
            @else
            <span class="badge b-warn">NO</span>
            @endif
        </td>
        </tr>
    @empty
        <tr><td colspan="10" class="text-center text-gray-500 py-6">Sin resultados para el período seleccionado.</td></tr>
    @endforelse
    </tbody>
      <tfoot>
        <tr>
          <th colspan="2" class="text-right">Totales:</th>
          <th class="text-center">{{ $totales['participantes'] }}</th>
          <th></th>
          <th></th>
          <th></th>
          <th></th>
          <th class="text-center">{{ $totales['horas_hombre'] }}</th>
          <th class="text-center">{{ $totales['numero'] }}</th>
          <th></th>
        </tr>
      </tfoot>
    </table>
  </div>

  <p class="text-sm text-gray-500">
    Nota: se incluyen todas las capacitaciones <strong>programadas (SI)</strong> aunque no se hayan impartido en el mes (aparecerán con 0 participantes),
    y además se agregan las que <strong>sí tuvieron asistentes en el mes</strong> aunque no estuvieran programadas.
  </p>
</div>
@endsection
