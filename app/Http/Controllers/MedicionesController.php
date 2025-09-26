<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\MessageBag;
use App\Exports\ReporteIluminacionExport;
use App\Exports\ReporteRuidoExport;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class MedicionesController extends Controller
{

    private function xlTitle(string $name): string
    {
        $name = preg_replace('/[\\\\\\/\\*\\?\\:\\[\\]]/', ' ', $name);
        $name = trim($name);
        if ($name === '') $name = 'Hoja';
        return mb_substr($name, 0, 31);
    }

    private function ensureSheet(Spreadsheet $book, string $wantedName, Worksheet $tpl): Worksheet
    {
        $title = $this->xlTitle($wantedName);

        // Si ya existe, úsala
        if ($sheet = $book->getSheetByName($title)) {
            return $sheet;
        }

        // Si no existe, generar un título único (Nombre, Nombre (2), ...)
        $base = $title;
        $i = 2;
        while ($book->getSheetByName($title)) {
            $title = $this->xlTitle(mb_substr($base, 0, 28) . " ($i)");
            $i++;
        }

        // Clonar la plantilla y agregarla
        $new = clone $tpl;
        $new->setTitle($title);      // fijar título antes de addSheet()
        $book->addSheet($new);

        return $new;
    }

    public function captureBatch(Request $request)
{
    $localizaciones = DB::table('localizacion')
        ->select('id_localizacion','localizacion')
        ->orderBy('localizacion')
        ->get();

    // Puestos con su localización para filtrar en la vista
    $puestos = DB::table('puesto_trabajo_matriz')
        ->select('id_puesto_trabajo_matriz','puesto_trabajo_matriz')
        ->orderBy('puesto_trabajo_matriz')
        ->get();

    // Estado por localizacion respecto al año actual (si ya hay datos capturados)
    $year = now()->year;

    // Contadores por localizacion para iluminacion y ruido en el año actual
    $luxCounts = DB::table('mediciones_iluminacion')
        ->select('id_localizacion', DB::raw('COUNT(*) as c'))
        ->whereRaw('YEAR(COALESCE(fecha_realizacion_inicio, fecha_realizacion_final)) = ?', [$year])
        ->groupBy('id_localizacion')
        ->pluck('c', 'id_localizacion');

    $ruidoCounts = DB::table('mediciones_ruido')
        ->select('id_localizacion', DB::raw('COUNT(*) as c'))
        ->whereRaw('YEAR(COALESCE(fecha_realizacion_inicio, fecha_realizacion_final)) = ?', [$year])
        ->groupBy('id_localizacion')
        ->pluck('c', 'id_localizacion');

    $locStatus = $localizaciones->map(function ($lo) use ($luxCounts, $ruidoCounts) {
        $lx = (int)($luxCounts[$lo->id_localizacion] ?? 0);
        $rd = (int)($ruidoCounts[$lo->id_localizacion] ?? 0);
        return (object) [
            'id'        => $lo->id_localizacion,
            'nombre'    => $lo->localizacion,
            'cnt_lux'   => $lx,
            'cnt_ruido' => $rd,
            'has'       => ($lx + $rd) > 0,
        ];
    });

    return view('mediciones.captura_batch', [
        'localizaciones' => $localizaciones,
        'puestos'        => $puestos,
        'locStatus'      => $locStatus,
        'year'           => $year,
    ]);
}

public function saveBatch(Request $request)
{
    $mode = $request->input('mode'); // 'create' | 'update'
    if (!in_array($mode, ['create','update'], true)) {
        return back()->withErrors(['mode'=>'Acción inválida'])->withInput();
    }

    // Cabecera común
    $base = $request->validate([
        'id_localizacion'          => ['required','integer','exists:localizacion,id_localizacion'],
        'departamento'             => ['nullable','string','max:500'],
        'nombre_observador'        => ['nullable','string','max:500'],
        'fecha_realizacion_inicio' => ['nullable','date'],
        'fecha_realizacion_final'  => ['nullable','date'],

        'instrumento_ruido'        => ['nullable','string','max:150'],
        'serie_ruido'              => ['nullable','string','max:200'],
        'marca_ruido'              => ['nullable','string','max:100'],
        'nrr'                      => ['nullable','numeric'],

        'instrumento_lux'          => ['nullable','string','max:150'],
        'serie_lux'                => ['nullable','string','max:200'],
        'marca_lux'                => ['nullable','string','max:100'],

        // originales (ocultos) para validar "create"
        'orig_fecha_inicio_ruido'  => ['nullable','date'],
        'orig_fecha_final_ruido'   => ['nullable','date'],
        'orig_fecha_inicio_lux'    => ['nullable','date'],
        'orig_fecha_final_lux'     => ['nullable','date'],

        // ids existentes (para detectar borrados en update)
        'existing_ruido_ids'       => ['nullable','string'],
        'existing_lux_ids'         => ['nullable','string'],
    ]);

    $ruidoP = $request->input('ruido_puntos', []);
    $luxP   = $request->input('iluminacion_puntos', []);

    // Validación de filas
    if (!empty($ruidoP)) {
        $request->validate([
            'ruido_puntos'                                  => ['array'],
            'ruido_puntos.*.id_row'                         => ['nullable','integer'],
            'ruido_puntos.*.id_puesto_trabajo_matriz'       => ['required','integer','exists:puesto_trabajo_matriz,id_puesto_trabajo_matriz'],
            'ruido_puntos.*.punto_medicion'                 => ['nullable','string','max:500'],
            'ruido_puntos.*.nivel_maximo'                   => ['nullable','numeric'],
            'ruido_puntos.*.nivel_minimo'                   => ['nullable','numeric'],
            'ruido_puntos.*.nivel_promedio'                 => ['nullable','numeric'],
            'ruido_puntos.*.nre'                            => ['nullable','numeric'],
            'ruido_puntos.*.limites_aceptables'             => ['nullable','numeric'],
            'ruido_puntos.*.observaciones'                  => ['nullable','string'],
        ]);
    }
    if (!empty($luxP)) {
        $request->validate([
            'iluminacion_puntos'                             => ['array'],
            'iluminacion_puntos.*.id_row'                    => ['nullable','integer'],
            'iluminacion_puntos.*.id_puesto_trabajo_matriz'  => ['required','integer','exists:puesto_trabajo_matriz,id_puesto_trabajo_matriz'],
            'iluminacion_puntos.*.punto_medicion'            => ['nullable','string','max:500'],
            'iluminacion_puntos.*.promedio'                  => ['nullable','numeric'],
            'iluminacion_puntos.*.limites_aceptables'        => ['nullable','numeric'],
            'iluminacion_puntos.*.observaciones'             => ['nullable','string'],
        ]);
    }

    // Reglas "Guardar como nuevo": si hay filas de ese tipo, debe cambiar al menos una fecha
    if ($mode === 'create') {
        $err = [];

        if (!empty($ruidoP)) {
            $oldI = $base['orig_fecha_inicio_ruido'] ?? null;
            $oldF = $base['orig_fecha_final_ruido']  ?? null;
            $newI = $base['fecha_realizacion_inicio'] ?? null;
            $newF = $base['fecha_realizacion_final']  ?? null;
            // Debe existir al menos una fecha y la pareja no debe ser exactamente igual a la original
            if (($newI === $oldI) && ($newF === $oldF)) {
                $err['fecha_realizacion_inicio'] = 'Para RUIdo, cambia la(s) fecha(s) para guardar como nuevo.';
            }
        }
        if (!empty($luxP)) {
            $oldI = $base['orig_fecha_inicio_lux'] ?? null;
            $oldF = $base['orig_fecha_final_lux']  ?? null;
            $newI = $base['fecha_realizacion_inicio'] ?? null;
            $newF = $base['fecha_realizacion_final']  ?? null;
            if (($newI === $oldI) && ($newF === $oldF)) {
                $err['fecha_realizacion_inicio'] = 'Para ILUMINACIÓN, cambia la(s) fecha(s) para guardar como nuevo.';
            }
        }
        if ($err) {
            return back()->withErrors($err)->withInput();
        }
    }

    $null = fn($v)=> ($v === '' || $v === null) ? null : $v;
    $fkLo = (int)$base['id_localizacion'];

    DB::transaction(function() use ($mode, $request, $base, $ruidoP, $luxP, $null, $fkLo) {

        // ---------- RUIdo ----------
        $postedIdsR = [];
        foreach ($ruidoP as $row) {
            $payload = [
                'departamento'              => $null($base['departamento']          ?? null),
                'fecha_realizacion_inicio'  => $null($base['fecha_realizacion_inicio'] ?? null),
                'fecha_realizacion_final'   => $null($base['fecha_realizacion_final']  ?? null),
                'nombre_observador'         => $null($base['nombre_observador']     ?? null),
                'instrumento'               => $null($request->input('instrumento_ruido')),
                'serie'                     => $null($request->input('serie_ruido')),
                'marca'                     => $null($request->input('marca_ruido')),
                'nrr'                       => $null($request->input('nrr')),
                'id_localizacion'           => $fkLo,
                'punto_medicion'            => $null($row['punto_medicion']         ?? null),
                'id_puesto_trabajo_matriz'  => (int)($row['id_puesto_trabajo_matriz'] ?? 0),
                'nivel_maximo'              => $null($row['nivel_maximo']           ?? null),
                'nivel_minimo'              => $null($row['nivel_minimo']           ?? null),
                'nivel_promedio'            => $null($row['nivel_promedio']         ?? null),
                'nre'                       => $null($row['nre']                     ?? null),
                'limites_aceptables'        => $null($row['limites_aceptables']     ?? null),
                'observaciones'             => $null($row['observaciones']          ?? null),
            ];

            $idRow = $row['id_row'] ?? null;
            if ($mode === 'update' && $idRow) {
                DB::table('mediciones_ruido')
                    ->where('id_mediciones_ruido', $idRow)
                    ->update($payload);
                $postedIdsR[] = (int)$idRow;
            } else {
                DB::table('mediciones_ruido')->insert($payload);
            }
        }

        // Eliminar filas quitadas en UPDATE (opcional)
        if ($mode === 'update') {
            $existing = array_filter(array_map('intval', explode(',', (string)($base['existing_ruido_ids'] ?? ''))));
            if ($existing) {
                $toDelete = array_diff($existing, $postedIdsR);
                if ($toDelete) {
                    DB::table('mediciones_ruido')->whereIn('id_mediciones_ruido', $toDelete)->delete();
                }
            }
        }

        // ---------- ILUMINACIÓN ----------
        $postedIdsL = [];
        foreach ($luxP as $row) {
            $payload = [
                'departamento'              => $null($base['departamento']          ?? null),
                'fecha_realizacion_inicio'  => $null($base['fecha_realizacion_inicio'] ?? null),
                'fecha_realizacion_final'   => $null($base['fecha_realizacion_final']  ?? null),
                'nombre_observador'         => $null($base['nombre_observador']     ?? null),
                'instrumento'               => $null($request->input('instrumento_lux')),
                'serie'                     => $null($request->input('serie_lux')),
                'marca'                     => $null($request->input('marca_lux')),
                'id_localizacion'           => $fkLo,
                'punto_medicion'            => $null($row['punto_medicion']         ?? null),
                'id_puesto_trabajo_matriz'  => (int)($row['id_puesto_trabajo_matriz'] ?? 0),
                'promedio'                  => $null($row['promedio']               ?? null),
                'limites_aceptables'        => $null($row['limites_aceptables']     ?? null),
                'observaciones'             => $null($row['observaciones']          ?? null),
            ];

            $idRow = $row['id_row'] ?? null;
            if ($mode === 'update' && $idRow) {
                DB::table('mediciones_iluminacion')
                    ->where('id', $idRow)
                    ->update($payload);
                $postedIdsL[] = (int)$idRow;
            } else {
                DB::table('mediciones_iluminacion')->insert($payload);
            }
        }

        if ($mode === 'update') {
            $existing = array_filter(array_map('intval', explode(',', (string)($base['existing_lux_ids'] ?? ''))));
            if ($existing) {
                $toDelete = array_diff($existing, $postedIdsL);
                if ($toDelete) {
                    DB::table('mediciones_iluminacion')->whereIn('id', $toDelete)->delete();
                }
            }
        }
    });

    return redirect()->route('mediciones.captura')
        ->with('ok', $mode === 'create' ? 'Guardado como nuevo registro.' : 'Cambios guardados en el registro existente.');
}

public function reporteIluminacion(\Illuminate\Http\Request $request)
{
    $yearInput   = $request->input('year');
    $currentYear = (int) now()->year;
    $year        = ($yearInput === null || $yearInput === '') ? $currentYear : (int) $yearInput;

    $availableYears = collect(
        \DB::table('mediciones_iluminacion as m')
            ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
            ->pluck('y')
    )->filter(function ($v) { return !is_null($v); })
     ->map(fn($v) => (int) $v)
     ->filter(function ($v) { return $v > 0; })
     ->unique()
     ->sortDesc()
     ->values();

    if (!$availableYears->contains($year)) {
        $availableYears = $availableYears->concat([$year])->unique()->sortDesc()->values();
    }

    // Catálogo de localizaciones
    $localizaciones = \DB::table('localizacion')
        ->select('id_localizacion','localizacion')
        ->orderBy('localizacion')
        ->get();

    // EM por localización (si hay varias filas del estándar, tomamos el mayor)
    $emByLoc = \DB::table('estandar_iluminacion')
        ->select('id_localizacion', \DB::raw('MAX(em) as em'))
        ->groupBy('id_localizacion')
        ->pluck('em', 'id_localizacion');   // [ id_localizacion => em ]

    // Traer filas del año
    $rows = \DB::table('mediciones_iluminacion as m')
        ->leftJoin('puesto_trabajo_matriz as p','p.id_puesto_trabajo_matriz','=','m.id_puesto_trabajo_matriz')
        ->select(
            'm.id',
            'm.id_localizacion',
            'm.punto_medicion',
            'm.id_puesto_trabajo_matriz',
            'p.puesto_trabajo_matriz as puesto',
            'm.promedio',
            'm.limites_aceptables',
            'm.acciones_correctivas',
            'm.observaciones'
        )
        ->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) = ?', [$year])
        ->orderBy('m.id_localizacion')
        ->orderBy('m.punto_medicion')
        ->get();

    // ⇨ Completar acciones cuando PROM < LÍMITE
    $rows = $rows->map(function ($r) use ($emByLoc) {
        $prom = is_numeric($r->promedio) ? (float)$r->promedio : null;
        $lim  = is_numeric($r->limites_aceptables)
                    ? (float)$r->limites_aceptables
                    : (isset($emByLoc[$r->id_localizacion]) ? (float)$emByLoc[$r->id_localizacion] : null);

        // si ambos son numéricos y el promedio es menor que el límite
        if ($prom !== null && $lim !== null && $prom < $lim) {
            if ($r->acciones_correctivas === null || trim($r->acciones_correctivas) === '') {
                $r->acciones_correctivas = 'Instalar mas lamparas';
            }
        }

        // opcional: llevar el límite final para mostrarlo fácil en la vista
        $r->lim_final = $lim;

        return $r;
    });

    // Agrupar por localización y ordenar dentro del grupo
    $filas = $rows
        ->groupBy('id_localizacion')
        ->map(function ($group) {
            return $group->sort(function ($a, $b) {
                $cmp = strnatcasecmp((string)($a->punto_medicion ?? ''),(string)($b->punto_medicion ?? ''));
                if ($cmp === 0) {
                    $cmp = strcasecmp((string)($a->puesto ?? ''),(string)($b->puesto ?? ''));
                }
                return $cmp;
            })->values();
        });

    $puestos = \DB::table('puesto_trabajo_matriz')
        ->select('id_puesto_trabajo_matriz','puesto_trabajo_matriz')
        ->orderBy('puesto_trabajo_matriz')
        ->get();

    return view('mediciones.reporte_iluminacion', [
        'year'           => $year,
        'localizaciones' => $localizaciones,
        'grupos'         => $filas,
        'puestos'        => $puestos,
        'years'          => $availableYears,
        'emByLoc'        => $emByLoc,
    ]);
}


public function updateAccionIluminacion(\Illuminate\Http\Request $request)
{
    $v = $request->validate([
        'id' => ['required','integer','exists:mediciones_iluminacion,id'],
        'acciones_correctivas' => ['nullable','string','max:5000'],
    ]);

    \DB::table('mediciones_iluminacion')
        ->where('id', $v['id'])
        ->update([
            'acciones_correctivas' => $v['acciones_correctivas'] ?? null,
        ]);

    return response()->json(['ok'=>true]);
}

public function reporteRuido(\Illuminate\Http\Request $request)
{
    $yearInput   = $request->input('year');
    $currentYear = (int) now()->year;
    $year        = ($yearInput === null || $yearInput === '') ? $currentYear : (int) $yearInput;

    $availableYears = collect(
        \DB::table('mediciones_ruido as m')
            ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
            ->pluck('y')
    )->filter(function ($v) { return !is_null($v); })
     ->map(fn($v) => (int) $v)
     ->filter(function ($v) { return $v > 0; })
     ->unique()
     ->sortDesc()
     ->values();

    if (!$availableYears->contains($year)) {
        $availableYears = $availableYears->concat([$year])->unique()->sortDesc()->values();
    }

    // Localizaciones para agrupar
    $localizaciones = \DB::table('localizacion')
        ->select('id_localizacion','localizacion')
        ->orderBy('localizacion')
        ->get();

    // Filas base + área del puesto
    $rawRows = \DB::table('mediciones_ruido as m')
        ->leftJoin('puesto_trabajo_matriz as p','p.id_puesto_trabajo_matriz','=','m.id_puesto_trabajo_matriz')
        ->leftJoin('area as a','a.id_area','=','p.id_area')
        ->select(
            'm.id_mediciones_ruido as id',
            'm.id_localizacion',
            'm.punto_medicion',
            'm.id_puesto_trabajo_matriz',
            'p.puesto_trabajo_matriz as puesto',
            'a.area as area_nombre',
            'm.nivel_maximo',
            'm.nivel_minimo',
            'm.nivel_promedio',
            'm.limites_aceptables',
            'm.acciones_correctivas',
            'm.observaciones'
        )
        ->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) = ?', [$year])
        ->orderBy('m.id_localizacion')
        ->orderBy('m.punto_medicion')
        ->get();

    // Regla: reducción EPP si PROMEDIO > 80 dB
    $REDUCCION_UMBRAL = 80.0;

    $calcRows = $rawRows->map(function ($r) use ($REDUCCION_UMBRAL) {
        $max = is_numeric($r->nivel_maximo) ? (float)$r->nivel_maximo : null;
        $min = is_numeric($r->nivel_minimo) ? (float)$r->nivel_minimo : null;

        $prom = (!is_null($max) && !is_null($min))
            ? ($max + $min) / 2.0
            : (is_numeric($r->nivel_promedio) ? (float)$r->nivel_promedio : null);

        // Límite mostrado (si no viene, 80)
        $lim = is_numeric($r->limites_aceptables) ? (float)$r->limites_aceptables : 80.0;

        // Cálculo de NRR/NRE según umbral 80 dB
        $nrr = null;
        $nre = $prom;
        $obs = $r->acciones_correctivas;

        if (!is_null($prom) && $prom > $REDUCCION_UMBRAL) {
            $nrr = (strcasecmp((string)$r->area_nombre, 'Area Interna') === 0) ? 22.00 : 27.00;
            $nre = $prom - $nrr;

            // Observación automática si viene vacía
            if ($obs === null || trim($obs) === '') {
                $obs = 'Uso obligatorio de protección auditiva';
            }
        }

        // Valores calculados para la vista
        $r->calc_promedio = is_null($prom) ? null : round($prom, 2);
        $r->calc_nrr      = is_null($nrr)  ? null : round($nrr, 2);
        $r->calc_nre      = is_null($nre)  ? null : round($nre, 2);
        $r->lim_final     = $lim;
        $r->acciones_correctivas = $obs; // deja la observación lista por si la muestras o exportas

        return $r;
    });

    $filas = $calcRows
        ->groupBy('id_localizacion')
        ->map(function ($rows) {
            return $rows->sort(function ($a, $b) {
                $cmp = strnatcasecmp((string)($a->punto_medicion ?? ''),(string)($b->punto_medicion ?? ''));
                if ($cmp === 0) {
                    $cmp = strcasecmp((string)($a->puesto ?? ''),(string)($b->puesto ?? ''));
                }
                return $cmp;
            })->values();
        });

    $puestos = \DB::table('puesto_trabajo_matriz')
        ->select('id_puesto_trabajo_matriz','puesto_trabajo_matriz')
        ->orderBy('puesto_trabajo_matriz')
        ->get();

    return view('mediciones.reporte_ruido', [
        'year'           => $year,
        'localizaciones' => $localizaciones,
        'grupos'         => $filas,
        'puestos'        => $puestos,
        'years'          => $availableYears,
    ]);
}

public function updateAccionRuido(\Illuminate\Http\Request $request)
{
    $v = $request->validate([
        'id' => ['required','integer','exists:mediciones_ruido,id_mediciones_ruido'],
        'acciones_correctivas' => ['nullable','string','max:255'],
        '_token' => ['required'],
    ]);

    \DB::table('mediciones_ruido')
        ->where('id_mediciones_ruido', $v['id'])
        ->update([
            'acciones_correctivas' => $v['acciones_correctivas'] ?? null,
        ]);

    return response()->json(['ok'=>true]);
}

// app/Http/Controllers/MedicionesController.php

public function timeline(\Illuminate\Http\Request $request)
{
    $locFilter  = $request->integer('id_localizacion'); // opcional

    $yearFrom = $request->integer('year_from');
    $yearTo   = $request->integer('year_to');
    $yearSingle = $request->integer('year');

    if ($yearSingle && $yearFrom === null && $yearTo === null) {
        $yearFrom = $yearSingle;
        $yearTo   = $yearSingle;
    }

    if ($yearFrom !== null && $yearTo === null) {
        $yearTo = $yearFrom;
    }
    if ($yearTo !== null && $yearFrom === null) {
        $yearFrom = $yearTo;
    }
    if ($yearFrom !== null && $yearTo !== null && $yearFrom > $yearTo) {
        [$yearFrom, $yearTo] = [$yearTo, $yearFrom];
    }

    $yearRange = ($yearFrom !== null && $yearTo !== null) ? [$yearFrom, $yearTo] : null;

    // Catálogos
    $locs = \DB::table('localizacion')
        ->select('id_localizacion','localizacion')
        ->orderBy('localizacion')->get();
    $locNames = $locs->pluck('localizacion','id_localizacion');
    $puestoNames = \DB::table('puesto_trabajo_matriz')
        ->pluck('puesto_trabajo_matriz','id_puesto_trabajo_matriz');

    // Límite de iluminación (EM) por (localización, puesto)
    $limLuxMap = \DB::table('estandar_iluminacion')
        ->select('id_localizacion', \DB::raw('MAX(em) as limite'))
        ->groupBy('id_localizacion')
        ->get()
        ->mapWithKeys(function($r){
            $p = $r->id_puesto_trabajo_matriz ?? 0;
            return [$r->id_localizacion.'|'.$p => $r->limite];
        });

    // Agregados por AÑO + PUNTO + PUESTO (+ loc para agrupar)
    $luxAgg = \DB::table('mediciones_iluminacion as m')
        ->selectRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
        ->addSelect('m.punto_medicion','m.id_puesto_trabajo_matriz','m.id_localizacion')
        ->selectRaw('AVG(m.promedio) as avg_lux, COUNT(*) as cnt_lux')
        ->when($locFilter,  fn($q)=>$q->where('m.id_localizacion',$locFilter))
        ->when($yearRange, fn($q)=>$q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) BETWEEN ? AND ?', $yearRange))
        ->groupBy('y','m.punto_medicion','m.id_puesto_trabajo_matriz','m.id_localizacion')
        ->get();

    $ruidoAgg = \DB::table('mediciones_ruido as m')
        ->selectRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
        ->addSelect('m.punto_medicion','m.id_puesto_trabajo_matriz','m.id_localizacion')
        ->selectRaw('AVG(m.nivel_promedio) as avg_ruido, COUNT(*) as cnt_ruido')
        ->when($locFilter,  fn($q)=>$q->where('m.id_localizacion',$locFilter))
        ->when($yearRange, fn($q)=>$q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) BETWEEN ? AND ?', $yearRange))
        ->groupBy('y','m.punto_medicion','m.id_puesto_trabajo_matriz','m.id_localizacion')
        ->get();

    // Años disponibles
    $yearsAll = collect(
        \DB::table('mediciones_iluminacion as m')
          ->when($locFilter, fn($q)=>$q->where('m.id_localizacion',$locFilter))
          ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')->pluck('y')
    )->merge(
        \DB::table('mediciones_ruido as m')
          ->when($locFilter, fn($q)=>$q->where('m.id_localizacion',$locFilter))
          ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')->pluck('y')
    )->filter()->unique()->sort()->values();
    $years = $yearRange ? collect(range($yearRange[0], $yearRange[1])) : $yearsAll;
    if ($yearRange && $years->isEmpty()) {
        $years = collect(range($yearRange[0], $yearRange[1]));
    }

    // Fusión por fila (loc+punto+puesto) y año
    $rowKey = function($locId, $puestoId, $punto) {
        $p = mb_strtolower(trim($punto ?? ''));
        $pid = $puestoId ?: 0;
        return $locId.'|'.$pid.'|'.$p;
    };
    $cells = [];

    foreach ($luxAgg as $r) {
        $k = $rowKey($r->id_localizacion, $r->id_puesto_trabajo_matriz, $r->punto_medicion);
        $cells[$k]['meta'] = [
            'loc_id'   => $r->id_localizacion,
            'punto'    => $r->punto_medicion,
            'puesto_id'=> $r->id_puesto_trabajo_matriz,
        ];
        $cells[$k]['data'][$r->y]['avg_lux'] = round((float)$r->avg_lux, 2);
        $cells[$k]['data'][$r->y]['cnt_lux'] = (int)$r->cnt_lux;
    }
    foreach ($ruidoAgg as $r) {
        $k = $rowKey($r->id_localizacion, $r->id_puesto_trabajo_matriz, $r->punto_medicion);
        $cells[$k]['meta'] = $cells[$k]['meta'] ?? [
            'loc_id'   => $r->id_localizacion,
            'punto'    => $r->punto_medicion,
            'puesto_id'=> $r->id_puesto_trabajo_matriz,
        ];
        $cells[$k]['data'][$r->y]['avg_ruido'] = round((float)$r->avg_ruido, 2);
        $cells[$k]['data'][$r->y]['cnt_ruido'] = (int)$r->cnt_ruido;
    }

    // Construcción de filas
    $rows = [];
    foreach ($cells as $entry) {
        $meta      = $entry['meta'];
        $locId     = $meta['loc_id'];
        $puestoId  = $meta['puesto_id'] ?: 0;
        $punto     = $meta['punto'] ?: '—';
        $locNom    = $locNames[$locId] ?? '';
        $puestoNom = $puestoNames[$puestoId] ?? '—';

        $nombre = "{$punto} — {$puestoNom}";
        $limKey = $locId.'|'.$puestoId;
        $limLux = $limLuxMap[$limKey] ?? null;

        $cols = [];
        foreach ($years as $y) {
            $d = $entry['data'][$y] ?? [];
            $cols[] = [
                'year'      => $y,
                'avg_lux'   => $d['avg_lux']   ?? null,
                'cnt_lux'   => $d['cnt_lux']   ?? 0,
                'avg_ruido' => $d['avg_ruido'] ?? null,
                'cnt_ruido' => $d['cnt_ruido'] ?? 0,
            ];
        }

        $rows[] = [
            'nombre'   => $nombre,     // Punto — Puesto
            'lim_lux'  => $limLux,
            'columns'  => $cols,
            'loc'      => $locNom,     // Para agrupar visualmente
            'loc_id'   => $locId,
            'punto'    => $punto,
            'puesto'   => $puestoNom,
        ];
    }

    // Orden: Localización > Punto
    usort($rows, function($a,$b){
        $cmp = strcmp($a['loc'] ?? '', $b['loc'] ?? '');
        if ($cmp !== 0) return $cmp;
        return strcmp($a['punto'] ?? '', $b['punto'] ?? '');
    });

    // Agrupar en categorías por localización
    $groups = [];
    foreach ($rows as $r) {
        $id = $r['loc_id'] ?? 0;
        $groups[$id]['loc']   = $r['loc'] ?? '—';
        $groups[$id]['rows'][] = $r;
    }

    // Ordenar por nombre de localización (sin helper)
    $groups = collect($groups)
        ->sortBy(fn($g) => $g['loc'] ?? '', SORT_NATURAL | SORT_FLAG_CASE)
        ->values()
        ->all();
    // Mantener orden por nombre de localización

    return view('mediciones.timeline', [
        'years'          => $years->all(),
        'yearsAll'       => $yearsAll->all(),
        'groups'         => $groups,     // <--- usamos groups en la vista
        'localizaciones' => $locs,
        'locFilter'      => $locFilter,
        'yearFrom'       => $yearFrom,
        'yearTo'         => $yearTo,
    ]);
}

public function timelineExcel(\Illuminate\Http\Request $request)
{
    // --- Hardening salida para evitar XLSX corruptos ---
    if (class_exists('\\Debugbar')) { \Debugbar::disable(); }
    @ini_set('zlib.output_compression', 'Off');
    while (ob_get_level() > 0) { @ob_end_clean(); }

    $locFilter  = $request->integer('id_localizacion'); // opcional
    $yearFrom   = $request->integer('year_from');
    $yearTo     = $request->integer('year_to');
    $yearSingle = $request->integer('year');

    if ($yearSingle && $yearFrom === null && $yearTo === null) { $yearFrom = $yearSingle; $yearTo = $yearSingle; }
    if ($yearFrom !== null && $yearTo === null) $yearTo = $yearFrom;
    if ($yearTo   !== null && $yearFrom === null) $yearFrom = $yearTo;
    if ($yearFrom !== null && $yearTo   !== null && $yearFrom > $yearTo) [ $yearFrom, $yearTo ] = [ $yearTo, $yearFrom ];
    $yearRange = ($yearFrom !== null && $yearTo !== null) ? [$yearFrom, $yearTo] : null;

    // Catálogos
    $locs        = \DB::table('localizacion')->select('id_localizacion','localizacion')->orderBy('localizacion')->get();
    $locNames    = $locs->pluck('localizacion','id_localizacion');
    $puestoNames = \DB::table('puesto_trabajo_matriz')->pluck('puesto_trabajo_matriz','id_puesto_trabajo_matriz');

    // Límite de iluminación (si tu tabla no tiene id_puesto..., queda 0 en la clave)
    $limLuxMap = \DB::table('estandar_iluminacion')
        ->select('id_localizacion', \DB::raw('MAX(em) as limite'))
        ->groupBy('id_localizacion')
        ->get()
        ->mapWithKeys(function($r){
            $p = $r->id_puesto_trabajo_matriz ?? 0;
            return [$r->id_localizacion.'|'.$p => $r->limite];
        });

    // ===== ILUMINACIÓN: valor crudo de 'promedio'; elegir más reciente por fecha y luego por id =====
    $luxRows = \DB::table('mediciones_iluminacion as m')
        ->selectRaw('
            YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y,
            m.punto_medicion,
            m.id_puesto_trabajo_matriz,
            m.id_localizacion,
            m.promedio,
            COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final) as fref,
            m.id
        ')
        ->when($locFilter,  fn($q)=>$q->where('m.id_localizacion',$locFilter))
        ->when($yearRange, fn($q)=>$q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) BETWEEN ? AND ?', $yearRange))
        ->get();

    $luxAgg = $luxRows
        ->groupBy(function($r){
            return implode('|', [
                $r->id_localizacion,
                $r->id_puesto_trabajo_matriz ?: 0,
                mb_strtolower(trim($r->punto_medicion ?? '')),
                $r->y
            ]);
        })
        ->map(function($grp){
            $best = $grp->reduce(function($best, $cur){
                if (!$best) return $cur;
                $dBest = $best->fref ?? '0000-00-00';
                $dCur  = $cur->fref ?? '0000-00-00';
                if ($dCur > $dBest) return $cur;
                if ($dCur === $dBest && $cur->id > $best->id) return $cur;
                return $best;
            });
            return (object)[
                'y'                        => $best->y,
                'punto_medicion'           => $best->punto_medicion,
                'id_puesto_trabajo_matriz' => $best->id_puesto_trabajo_matriz,
                'id_localizacion'          => $best->id_localizacion,
                'lux_val'                  => is_null($best->promedio) ? null : (float)$best->promedio,
                'cnt_lux'                  => $grp->count(),
            ];
        })
        ->values();

    // ===== RUIDO: promedio del punto medio (min+max)/2 por grupo/año =====
    $ruidoAgg = \DB::table('mediciones_ruido as m')
        ->selectRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
        ->addSelect('m.punto_medicion','m.id_puesto_trabajo_matriz','m.id_localizacion')
        ->selectRaw('AVG((COALESCE(m.nivel_minimo, m.nivel_promedio) + COALESCE(m.nivel_maximo, m.nivel_promedio))/2) as avg_ruido,
                     COUNT(*) as cnt_ruido')
        ->when($locFilter,  fn($q)=>$q->where('m.id_localizacion',$locFilter))
        ->when($yearRange, fn($q)=>$q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) BETWEEN ? AND ?', $yearRange))
        ->groupBy('y','m.punto_medicion','m.id_puesto_trabajo_matriz','m.id_localizacion')
        ->get();

    // Años disponibles
    $yearsAll = collect(
        \DB::table('mediciones_iluminacion as m')
          ->when($locFilter, fn($q)=>$q->where('m.id_localizacion',$locFilter))
          ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')->pluck('y')
    )->merge(
        \DB::table('mediciones_ruido as m')
          ->when($locFilter, fn($q)=>$q->where('m.id_localizacion',$locFilter))
          ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')->pluck('y')
    )->filter()->unique()->sort()->values();

    $years = $yearRange ? collect(range($yearRange[0], $yearRange[1])) : $yearsAll;
    if ($yearRange && $years->isEmpty()) $years = collect(range($yearRange[0], $yearRange[1]));

    // ===== Fusión de datos para Excel =====
    $rowKey = function($locId, $puestoId, $punto) {
        $p = mb_strtolower(trim($punto ?? ''));
        $pid = $puestoId ?: 0;
        return $locId.'|'.$pid.'|'.$p;
    };
    $cells = [];

    foreach ($luxAgg as $r) {
        $k = $rowKey($r->id_localizacion, $r->id_puesto_trabajo_matriz, $r->punto_medicion);
        $cells[$k]['meta'] = [
            'loc_id'    => $r->id_localizacion,
            'punto'     => $r->punto_medicion,
            'puesto_id' => $r->id_puesto_trabajo_matriz,
        ];
        $cells[$k]['data'][$r->y]['lux']      = $r->lux_val;      // valor crudo
        $cells[$k]['data'][$r->y]['cnt_lux']  = (int)$r->cnt_lux;
    }
    foreach ($ruidoAgg as $r) {
        $k = $rowKey($r->id_localizacion, $r->id_puesto_trabajo_matriz, $r->punto_medicion);
        $cells[$k]['meta'] = $cells[$k]['meta'] ?? [
            'loc_id'    => $r->id_localizacion,
            'punto'     => $r->punto_medicion,
            'puesto_id' => $r->id_puesto_trabajo_matriz,
        ];
        $cells[$k]['data'][$r->y]['avg_ruido'] = is_null($r->avg_ruido) ? null : (float)$r->avg_ruido;
        $cells[$k]['data'][$r->y]['cnt_ruido'] = (int)$r->cnt_ruido;
    }

    $rows = [];
    foreach ($cells as $entry) {
        $meta      = $entry['meta'];
        $locId     = $meta['loc_id'];
        $puestoId  = $meta['puesto_id'] ?: 0;
        $punto     = $meta['punto'] ?: '—';
        $locNom    = $locNames[$locId] ?? '';
        $puestoNom = $puestoNames[$puestoId] ?? '—';
        $limKey    = $locId.'|'.$puestoId;
        $limLux    = $limLuxMap[$limKey] ?? null;

        $cols = [];
        foreach ($years as $y) {
            $d = $entry['data'][$y] ?? [];
            $cols[] = [
                'year'      => $y,
                'lux'       => $d['lux']        ?? null,
                'cnt_lux'   => $d['cnt_lux']    ?? 0,
                'avg_ruido' => $d['avg_ruido']  ?? null,
                'cnt_ruido' => $d['cnt_ruido']  ?? 0,
            ];
        }

        $rows[] = [
            'loc'     => $locNom,
            'loc_id'  => $locId,
            'punto'   => $punto,
            'puesto'  => $puestoNom,
            'lim_lux' => $limLux,
            'columns' => $cols,
        ];
    }

    usort($rows, function($a,$b){
        $cmp = strcmp($a['loc'] ?? '', $b['loc'] ?? '');
        if ($cmp !== 0) return $cmp;
        return strcmp($a['punto'] ?? '', $b['punto'] ?? '');
    });

    $groups = [];
    foreach ($rows as $r) {
        $id = $r['loc_id'] ?? 0;
        $groups[$id]['loc']    = $r['loc'] ?? '—';
        $groups[$id]['rows'][] = $r;
    }
    $groups = collect($groups)->sortBy(fn($g) => $g['loc'] ?? '', SORT_NATURAL | SORT_FLAG_CASE)->values()->all();

    // ===== Excel =====
    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle(mb_substr('Comparativa', 0, 31));

    $brand = '00B0F0';
    $startYearColIdx = 6; // F
    $lastColIdx = empty($years) ? 5 : ($startYearColIdx + (count($years)*4) - 1);
    $lastCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($lastColIdx);

    // Helper A1
    $addr = function(int $col, int $row) {
        return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
    };

    // Título
    $sheet->setCellValue('A1', 'Comparativa anual por punto y puesto');
    $sheet->mergeCells('A1:C1');
    $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);

    // Cabeceras base
    $sheet->setCellValue('A2', 'Localización');
    $sheet->setCellValue('B2', 'Punto');
    $sheet->setCellValue('C2', 'Puesto');
    $sheet->setCellValue('D2', 'Límite lux');
    $sheet->setCellValue('E2', 'Límite ruido (dBA)');

    // Cabeceras de años
    foreach ($years as $i => $y) {
        $colStart = $startYearColIdx + ($i*4);
        $c1 = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colStart);
        $c2 = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colStart+3);
        $sheet->mergeCells("{$c1}2:{$c2}2");
        $sheet->setCellValue("{$c1}2", $y);
        $sheet->setCellValue($addr($colStart,   3), 'Lux');
        $sheet->setCellValue($addr($colStart+1, 3), 'Lux puntos');
        $sheet->setCellValue($addr($colStart+2, 3), 'Ruido media (dBA)');
        $sheet->setCellValue($addr($colStart+3, 3), 'Ruido puntos');
    }

    // Estilos cabecera
    $sheet->getStyle("A2:{$lastCol}3")->getFont()->setBold(true)->getColor()->setRGB('FFFFFF');
    $sheet->getStyle("A2:{$lastCol}3")->getAlignment()
          ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
          ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER)
          ->setWrapText(true);
    $sheet->getStyle("A2:{$lastCol}3")->getFill()
          ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
          ->getStartColor()->setRGB($brand);
    $sheet->getStyle("A2:{$lastCol}3")->getBorders()->getAllBorders()
          ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN)
          ->getColor()->setRGB('000000');
    $sheet->getRowDimension(2)->setRowHeight(24);
    $sheet->getRowDimension(3)->setRowHeight(22);

    // Anchos y congelar
    $sheet->getColumnDimension('A')->setWidth(28);
    $sheet->getColumnDimension('B')->setWidth(26);
    $sheet->getColumnDimension('C')->setWidth(28);
    $sheet->getColumnDimension('D')->setWidth(14);
    $sheet->getColumnDimension('E')->setWidth(18);
    for ($ci = 6; $ci <= $lastColIdx; $ci++) {
        $sheet->getColumnDimension(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($ci))->setWidth(14);
    }
    $sheet->freezePane('F4');

    // Datos
    $rowNum = 4;
    foreach ($groups as $g) {
        // Fila categoría
        $sheet->mergeCells("A{$rowNum}:{$lastCol}{$rowNum}");
        $sheet->setCellValue("A{$rowNum}", 'Localización: '.($g['loc'] ?? '—'));
        $sheet->getStyle("A{$rowNum}:{$lastCol}{$rowNum}")->getBorders()->getAllBorders()
              ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN)->getColor()->setRGB('000000');
        $sheet->getStyle("A{$rowNum}:{$lastCol}{$rowNum}")->getFill()
              ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('E8F3FF');
        $sheet->getStyle("A{$rowNum}:{$lastCol}{$rowNum}")->getFont()->setBold(true)->getColor()->setRGB('0F172A');
        $rowNum++;

        foreach (($g['rows'] ?? []) as $r) {
            $sheet->setCellValue("A{$rowNum}", $r['loc'] ?? '—');
            $sheet->setCellValue("B{$rowNum}", $r['punto'] ?? '—');
            $sheet->setCellValue("C{$rowNum}", $r['puesto'] ?? '—');

            $sheet->setCellValue("D{$rowNum}", $r['lim_lux'] ?? null);
            $sheet->getStyle("D{$rowNum}")->getNumberFormat()->setFormatCode('#,##0.00');

            $sheet->setCellValue("E{$rowNum}", 85);
            $sheet->getStyle("E{$rowNum}")->getNumberFormat()->setFormatCode('#,##0');

            // index por año
            $byYear = [];
            foreach ($r['columns'] as $c) { $byYear[$c['year']] = $c; }

            foreach ($years as $i => $y) {
                $colStart = $startYearColIdx + ($i*4);
                $luxVal = $byYear[$y]['lux']        ?? null;
                $luxCnt = $byYear[$y]['cnt_lux']    ?? null;
                $ruiAvg = $byYear[$y]['avg_ruido']  ?? null;
                $ruiCnt = $byYear[$y]['cnt_ruido']  ?? null;

                $sheet->setCellValue($addr($colStart,   $rowNum), $luxVal);
                $sheet->setCellValue($addr($colStart+1, $rowNum), $luxCnt);
                $sheet->setCellValue($addr($colStart+2, $rowNum), $ruiAvg);
                $sheet->setCellValue($addr($colStart+3, $rowNum), $ruiCnt);

                $sheet->getStyle($addr($colStart,   $rowNum))->getNumberFormat()->setFormatCode('#,##0.00');
                $sheet->getStyle($addr($colStart+1, $rowNum))->getNumberFormat()->setFormatCode('#,##0');
                $sheet->getStyle($addr($colStart+2, $rowNum))->getNumberFormat()->setFormatCode('#,##0.00');
                $sheet->getStyle($addr($colStart+3, $rowNum))->getNumberFormat()->setFormatCode('#,##0');
            }

            // bordes fila
            $sheet->getStyle("A{$rowNum}:{$lastCol}{$rowNum}")->getBorders()->getAllBorders()
                  ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN)->getColor()->setRGB('000000');
            $rowNum++;
        }
    }

    // Alineación y página
    $sheet->getStyle("A4:C{$rowNum}")->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
    $sheet->getStyle("D4:{$lastCol}{$rowNum}")->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $ps = $sheet->getPageSetup();
    $ps->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
    $ps->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER);
    $ps->setFitToWidth(1)->setFitToHeight(0);
    $margins = $sheet->getPageMargins();
    $margins->setTop(0.3)->setBottom(0.3)->setLeft(0.25)->setRight(0.25);

    // ===== Guardar y descargar de forma ultra-segura =====
    $filename = 'Comparativa_mediciones_'.date('Ymd_His').'.xlsx';
    $tmp = tempnam(sys_get_temp_dir(), 'xlsx_');

    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
    $writer->setPreCalculateFormulas(false);
    $writer->save($tmp);

    // Verificación ZIP: debe iniciar con "PK"
    $fh = fopen($tmp, 'rb');
    $head = $fh ? fread($fh, 2) : '';
    if ($fh) fclose($fh);
    if ($head !== 'PK') {
        return response('XLSX inválido: el archivo no inicia con PK (puede haber salida previa en la respuesta).', 500);
    }

    return response()->streamDownload(function() use ($tmp) {
        $out = fopen('php://output', 'wb');
        $in  = fopen($tmp, 'rb');
        stream_copy_to_stream($in, $out);
        fclose($in);
        fclose($out);
        @unlink($tmp);
    }, $filename, [
        'Content-Type'  => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control' => 'no-store, no-cache, must-revalidate, max-age=0',
        'Pragma'        => 'no-cache',
        'Expires'       => '0',
    ]);
}

public function deleteIluminacion(\Illuminate\Http\Request $request)
{
    $v = $request->validate([
        'id' => ['required','integer','exists:mediciones_iluminacion,id'],
        // _token lo verifica el middleware CSRF
    ]);

    \DB::table('mediciones_iluminacion')
        ->where('id', $v['id'])
        ->delete();

    // Para llamadas AJAX
    if ($request->wantsJson()) {
        return response()->json(['ok' => true]);
    }

    // Por si llamaras desde un form normal
    return back()->with('ok', 'Registro eliminado.');
}

public function deleteRuido(\Illuminate\Http\Request $request)
{
    $v = $request->validate([
        'id' => ['required','integer','exists:mediciones_ruido,id_mediciones_ruido'],
    ]);

    \DB::table('mediciones_ruido')
        ->where('id_mediciones_ruido', $v['id'])
        ->delete();

    if ($request->wantsJson()) {
        return response()->json(['ok' => true]);
    }
    return back()->with('ok', 'Registro de ruido eliminado.');
}

public function updateIluminacionRow(\Illuminate\Http\Request $request)
{
    $v = $request->validate([
        'id'                 => ['required','integer','exists:mediciones_iluminacion,id'],
        'punto_medicion'     => ['nullable','string','max:500'],
        'id_puesto_trabajo_matriz' => ['nullable','integer','exists:puesto_trabajo_matriz,id_puesto_trabajo_matriz'],
        'promedio'           => ['nullable','numeric'],
        'limites_aceptables' => ['nullable','numeric'],
    ]);

    $null = fn($x) => ($x === '' ? null : $x);
    $puesto = $request->input('id_puesto_trabajo_matriz');
    $puesto = ($puesto === null || $puesto === '') ? null : (int) $puesto;

    \DB::table('mediciones_iluminacion')
        ->where('id', $v['id'])
        ->update([
            'punto_medicion'            => $null($request->input('punto_medicion')),
            'id_puesto_trabajo_matriz' => $puesto,
            'promedio'                 => $null($request->input('promedio')),
            'limites_aceptables'       => $null($request->input('limites_aceptables')),
        ]);

    return response()->json(['ok' => true]);
}

public function updateRuidoRow(\Illuminate\Http\Request $request)
{
    $v = $request->validate([
        'id'                 => ['required','integer','exists:mediciones_ruido,id_mediciones_ruido'],
        'punto_medicion'     => ['nullable','string','max:500'],
        'id_puesto_trabajo_matriz' => ['nullable','integer','exists:puesto_trabajo_matriz,id_puesto_trabajo_matriz'],
        'nivel_maximo'       => ['nullable','numeric'],
        'nivel_minimo'       => ['nullable','numeric'],
        'nivel_promedio'     => ['nullable','numeric'],
        'nrr'                => ['nullable','numeric'],
        'nre'                => ['nullable','numeric'],
        'limites_aceptables' => ['nullable','numeric'],
    ]);

    $null = fn($x) => ($x === '' ? null : $x);
    $puesto = $request->input('id_puesto_trabajo_matriz');
    $puesto = ($puesto === null || $puesto === '') ? null : (int) $puesto;

    \DB::table('mediciones_ruido')
        ->where('id_mediciones_ruido', $v['id'])
        ->update([
            'punto_medicion'            => $null($request->input('punto_medicion')),
            'id_puesto_trabajo_matriz' => $puesto,
            'nivel_maximo'             => $null($request->input('nivel_maximo')),
            'nivel_minimo'             => $null($request->input('nivel_minimo')),
            'nivel_promedio'           => $null($request->input('nivel_promedio')),
            'nrr'                      => $null($request->input('nrr')),
            'nre'                      => $null($request->input('nre')),
            'limites_aceptables'       => $null($request->input('limites_aceptables')),
        ]);

    return response()->json(['ok' => true]);
}

public function charts(\Illuminate\Http\Request $request)
{
    $locId = $request->integer('id_localizacion');   // requerido para no saturar

    $yearFrom = $request->integer('year_from');
    $yearTo   = $request->integer('year_to');
    $yearSingle = $request->integer('year');

    if ($yearSingle && $yearFrom === null && $yearTo === null) {
        $yearFrom = $yearSingle;
        $yearTo   = $yearSingle;
    }

    if ($yearFrom !== null && $yearTo === null) {
        $yearTo = $yearFrom;
    }
    if ($yearTo !== null && $yearFrom === null) {
        $yearFrom = $yearTo;
    }
    if ($yearFrom !== null && $yearTo !== null && $yearFrom > $yearTo) {
        [$yearFrom, $yearTo] = [$yearTo, $yearFrom];
    }

    $yearRange = ($yearFrom !== null && $yearTo !== null) ? [$yearFrom, $yearTo] : null;

    // Catálogos
    $locs = \DB::table('localizacion')
        ->select('id_localizacion','localizacion')
        ->orderBy('localizacion')->get();
    $locNames = $locs->pluck('localizacion','id_localizacion');

    $puestoNames = \DB::table('puesto_trabajo_matriz')
        ->pluck('puesto_trabajo_matriz','id_puesto_trabajo_matriz');

    // Años disponibles (para el datalist del filtro)
    $yearsAll = collect(
        \DB::table('mediciones_iluminacion as m')
          ->when($locId, fn($q)=>$q->where('m.id_localizacion',$locId))
          ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
          ->pluck('y')
    )->merge(
        \DB::table('mediciones_ruido as m')
          ->when($locId, fn($q)=>$q->where('m.id_localizacion',$locId))
          ->selectRaw('DISTINCT YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
          ->pluck('y')
    )->filter()->unique()->sort()->values();

    // Si no hay localización seleccionada, no armamos series para evitar 1000 líneas
    $seriesLux = $seriesRuido = [];
    $years = $yearRange ? range($yearRange[0], $yearRange[1]) : $yearsAll->all();

    if ($locId) {
        // Agregados por AÑO + PUNTO + PUESTO dentro de la localización
        $luxAgg = \DB::table('mediciones_iluminacion as m')
            ->selectRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
            ->addSelect('m.punto_medicion','m.id_puesto_trabajo_matriz')
            ->selectRaw('AVG(m.promedio) as avg_lux')
            ->where('m.id_localizacion',$locId)
            ->when($yearRange, fn($q)=>$q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) BETWEEN ? AND ?', $yearRange))
            ->groupBy('y','m.punto_medicion','m.id_puesto_trabajo_matriz')
            ->get();

        $ruidoAgg = \DB::table('mediciones_ruido as m')
            ->selectRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) as y')
            ->addSelect('m.punto_medicion','m.id_puesto_trabajo_matriz')
            ->selectRaw('AVG(m.nivel_promedio) as avg_ruido')
            ->where('m.id_localizacion',$locId)
            ->when($yearRange, fn($q)=>$q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) BETWEEN ? AND ?', $yearRange))
            ->groupBy('y','m.punto_medicion','m.id_puesto_trabajo_matriz')
            ->get();

        // Etiquetas de años visibles (si filtran por año, solo ese; si no, todos)
        $years = $yearRange ? range($yearRange[0], $yearRange[1]) : $yearsAll->all();

        // Mapear por fila (punto+puesto) y año
        $key = fn($punto,$puestoId) => mb_strtolower(trim($punto ?? '')).'|'.($puestoId ?: 0);
        $rows = []; // key => ['label'=>..., 'data'=>[y=>valor]]
        foreach ($luxAgg as $r) {
            $k = $key($r->punto_medicion, $r->id_puesto_trabajo_matriz);
            $label = ($r->punto_medicion ?: '—').' — '.($puestoNames[$r->id_puesto_trabajo_matriz ?? 0] ?? '—');
            $rows[$k]['label'] = $label;
            $rows[$k]['lux'][$r->y] = round((float)$r->avg_lux,2);
        }
        foreach ($ruidoAgg as $r) {
            $k = $key($r->punto_medicion, $r->id_puesto_trabajo_matriz);
            $label = ($r->punto_medicion ?: '—').' — '.($puestoNames[$r->id_puesto_trabajo_matriz ?? 0] ?? '—');
            $rows[$k]['label'] = $label;
            $rows[$k]['ruido'][$r->y] = round((float)$r->avg_ruido,2);
        }

        // Construir series (alineadas a la lista de years)
        foreach ($rows as $r) {
            $dataLux   = []; $hasLux = false;
            $dataRuido = []; $hasRui = false;
            foreach ($years as $y) {
                $vL = $r['lux'][$y]   ?? null; $dataLux[]   = $vL; $hasLux = $hasLux || $vL !== null;
                $vR = $r['ruido'][$y] ?? null; $dataRuido[] = $vR; $hasRui = $hasRui || $vR !== null;
            }
            if ($hasLux)   $seriesLux[]   = ['label'=>$r['label'], 'data'=>$dataLux];
            if ($hasRui)   $seriesRuido[] = ['label'=>$r['label'], 'data'=>$dataRuido];
        }

        // Ordenar series por nombre para que sean consistentes
        usort($seriesLux,   fn($a,$b)=>strcasecmp($a['label'],$b['label']));
        usort($seriesRuido, fn($a,$b)=>strcasecmp($a['label'],$b['label']));
    }

    return view('mediciones.charts', [
        'localizaciones' => $locs,
        'locId'          => $locId,
        'locName'        => $locId ? ($locNames[$locId] ?? '') : '',
        'years'          => $years,
        'yearsAll'       => $yearsAll->all(),
        'yearFrom'       => $yearFrom,
        'yearTo'         => $yearTo,
        'seriesLux'      => $seriesLux,
        'seriesRuido'    => $seriesRuido,
    ]);
}

public function exportIluminacion(Request $request)
{
    $year  = $request->integer('year') ?: (int) now()->year;
    $locId = $request->integer('id_localizacion');
    $fname = 'reporte_iluminacion' . ($locId ? "_loc{$locId}" : '') . ($year ? "_{$year}" : '') . '.xlsx';

    $spreadsheet = (new ReporteIluminacionExport($year, $locId))->build();

    $tempDir = storage_path('app/temp_exports');
    if (!is_dir($tempDir) && !mkdir($tempDir, 0775, true) && !is_dir($tempDir)) {
        abort(500, 'No se pudo preparar el directorio temporal.');
    }

    $tmpFile = tempnam($tempDir, 'xls_');
    if ($tmpFile === false) {
        abort(500, 'No se pudo crear el archivo temporal.');
    }
    @unlink($tmpFile);
    $tmpPath = $tmpFile . '.xlsx';

    $writer = new Xlsx($spreadsheet);
    $writer->save($tmpPath);
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);

    if (function_exists('ini_set')) { @ini_set('zlib.output_compression', 'Off'); }
    while (ob_get_level() > 0) { @ob_end_clean(); }

    // (opcional) sanity check
    if (!is_file($tmpPath) || filesize($tmpPath) < 100) {
        @unlink($tmpPath);
        abort(500, 'El archivo no se generó correctamente.');
    }

    return response()->download($tmpPath, $fname, [
        'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
        'Pragma'                    => 'public',
        'Content-Transfer-Encoding' => 'binary',
    ])->deleteFileAfterSend(true);
}

    public function exportRuido(Request $request)
{
    $year = $request->integer('year') ?: (int) now()->year;
    $locId = $request->integer('id_localizacion');
    $fname = 'reporte_ruido' . ($locId ? "_loc{$locId}" : '') . ($year ? "_{$year}" : '') . '.xlsx';

    $spreadsheet = (new ReporteRuidoExport($year, $locId))->build();

    $tempDir = storage_path('app/temp_exports');
    if (!is_dir($tempDir) && !mkdir($tempDir, 0775, true) && !is_dir($tempDir)) {
        abort(500, 'No se pudo preparar el directorio temporal.');
    }

    $tmpFile = tempnam($tempDir, 'xls_');
    if ($tmpFile === false) {
        abort(500, 'No se pudo crear el archivo temporal.');
    }
    @unlink($tmpFile);
    $tmpPath = $tmpFile . '.xlsx';

    $writer = new Xlsx($spreadsheet);
    $writer->save($tmpPath);
    $spreadsheet->disconnectWorksheets();
    unset($spreadsheet);

    if (function_exists('ini_set')) { @ini_set('zlib.output_compression', 'Off'); }
    while (ob_get_level() > 0) { @ob_end_clean(); }

    if (!is_file($tmpPath) || filesize($tmpPath) < 100) {
        @unlink($tmpPath);
        abort(500, 'El archivo no se generó correctamente.');
    }

    return response()->download($tmpPath, $fname, [
        'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
        'Pragma'                    => 'public',
        'Content-Transfer-Encoding' => 'binary',
    ])->deleteFileAfterSend(true);
}

public function exportIluminacionDesdePlantilla(Request $request)
{
    $year = $request->integer('year') ?: (int) now()->year;

    $tplPath = storage_path('app/public/medicion_iluminacion.xlsx');
    if (!is_file($tplPath)) {
        abort(404, 'No se encontró la plantilla medicion_iluminacion.xlsx en storage/app/public');
    }
    $spreadsheet = IOFactory::load($tplPath);
    /** @var Worksheet $tplSheet */
    $tplSheet = $spreadsheet->getSheetByName('Plantilla') ?? $spreadsheet->getActiveSheet();

    $headerRow   = 12;
    $firstData   = 13;
    $perPage     = 30;
    $colsRange   = 'A:G';

    $colNo     = 'A';
    $colArea   = 'B';
    $colPunto  = 'C';
    $colPuesto = 'D';
    $colProm   = 'E';
    $colLim    = 'F';
    $colObs    = 'G';

    $emByLoc = DB::table('estandar_iluminacion')
        ->select('id_localizacion', DB::raw('MAX(em) as em'))
        ->groupBy('id_localizacion')
        ->pluck('em', 'id_localizacion');

    $base = DB::table('mediciones_iluminacion as m')
        ->leftJoin('localizacion as loc', 'loc.id_localizacion', '=', 'm.id_localizacion')
        ->leftJoin('puesto_trabajo_matriz as p', 'p.id_puesto_trabajo_matriz', '=', 'm.id_puesto_trabajo_matriz')
        ->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) = ?', [$year])
        ->orderBy('m.departamento')
        ->orderBy('m.id_localizacion')
        ->orderBy('m.punto_medicion');

    $departamentos = (clone $base)->pluck('m.departamento')->filter()->unique()->values();

    $copyRow = function(Worksheet $sheet, int $srcRow, int $dstRow, string $rangeCols) {
        [$colStart, $colEnd] = explode(':', $rangeCols);
        for ($col = $colStart; $col <= $colEnd; $col++) {
            $src = $col.$srcRow; $dst = $col.$dstRow;
            $cell = $sheet->getCell($src);
            $sheet->setCellValueExplicit($dst, $cell->getValue(), $cell->getDataType());
            $sheet->duplicateStyle($sheet->getStyle($src), $dst);
            $sheet->getColumnDimension($col)->setWidth($sheet->getColumnDimension($col)->getWidth());
        }
        $sheet->getRowDimension($dstRow)->setRowHeight($sheet->getRowDimension($srcRow)->getRowHeight());

        foreach ($sheet->getMergeCells() as $merge) {
            [$mStart, $mEnd] = explode(':', $merge);
            if (preg_match('/([A-Z]+)(\d+)/', $mStart, $m1) && (int)$m1[2] === $srcRow &&
                preg_match('/([A-Z]+)(\d+)/', $mEnd, $m2)) {
                $sheet->mergeCells($m1[1].$dstRow.':'.$m2[1].$dstRow);
            }
        }
    };

    $dataColumns = ['A','B','C','D','E','F','G'];
    $highestTplRow = $tplSheet->getHighestRow();

    $footerTemplateRows = [];
    $detectedFooterStart = null;
    for ($probe = $firstData; $probe <= $highestTplRow; $probe++) {
        $cellA = $tplSheet->getCell('A'.$probe)->getValue();
        $cellATrim = is_string($cellA) ? trim($cellA) : $cellA;
        if (is_numeric($cellATrim)) {
            continue;
        }
        $rowHasContent = false;
        foreach ($dataColumns as $col) {
            $val = $tplSheet->getCell($col.$probe)->getValue();
            if ($val !== null && $val !== '') {
                $rowHasContent = true;
                break;
            }
        }
        if ($rowHasContent) {
            $detectedFooterStart = $probe;
            break;
        }
    }
    if ($detectedFooterStart === null) {
        $detectedFooterStart = $firstData + 30;
    }
    for ($probe = $detectedFooterStart; $probe <= $highestTplRow; $probe++) {
        $rowHasContent = false;
        foreach ($dataColumns as $col) {
            $val = $tplSheet->getCell($col.$probe)->getValue();
            if ($val !== null && $val !== '') {
                $rowHasContent = true;
                break;
            }
        }
        if (!$rowHasContent) {
            if (!empty($footerTemplateRows)) {
                break;
            }
            continue;
        }
        $footerTemplateRows[] = $probe;
    }
    if (!$footerTemplateRows) {
        $footerTemplateRows = [$firstData + 30];
    }

    $footerRowCount = count($footerTemplateRows);
    $perPage = 30;
    $footerBaseRow = $firstData + $perPage;
    $blockHeight   = 1 + $perPage + $footerRowCount;


    foreach ($departamentos as $dep) {

        // ✅ usar helper como método y nombre correcto de variable
        $sheet = $this->ensureSheet($spreadsheet, $dep, $tplSheet);

        $rows = (clone $base)
            ->where('m.departamento', $dep)   // <-- $dep (no $dept)
            ->select(
                'm.departamento',
                'm.nombre_observador',
                'm.instrumento',
                'm.serie',
                'm.marca',
                'm.id_localizacion',
                'loc.localizacion as area_nombre',
                'm.punto_medicion',
                'p.puesto_trabajo_matriz as puesto_nombre',
                'm.promedio',
                'm.limites_aceptables',
                'm.observaciones',
                'm.fecha_realizacion_inicio',
                'm.fecha_realizacion_final'
            )
            ->get();

        if ($rows->isEmpty()) { continue; }

        // ---- Encabezado (D4.., B8, E8)
        $minIni = $rows->min('fecha_realizacion_inicio');
        $maxFin = $rows->max('fecha_realizacion_final');

        $sheet->setCellValue('D4', $dep ?? '');
        $sheet->setCellValue('D5', (string) ($rows->first()->nombre_observador ?? ''));

        $fechaRealizacion = '';
        if ($minIni && $maxFin) {
            $ini = (new \DateTime($minIni))->format('d/m/Y');
            $fin = (new \DateTime($maxFin))->format('d/m/Y');
            $fechaRealizacion = $minIni === $maxFin ? $ini : $ini.' al '.$fin;
        } elseif ($minIni) {
            $fechaRealizacion = (new \DateTime($minIni))->format('d/m/Y');
        } elseif ($maxFin) {
            $fechaRealizacion = (new \DateTime($maxFin))->format('d/m/Y');
        }
        $sheet->setCellValue('D6', $fechaRealizacion);

        $sheet->setCellValue('D7', (string) ($rows->first()->instrumento ?? ''));
        $sheet->setCellValue('B8', (string) ($rows->first()->serie ?? ''));
        $sheet->setCellValue('E8', (string) ($rows->first()->marca ?? ''));

        $footerBaseRows = [];

        foreach ($footerTemplateRows as $idx => $tplRow) {

            $destRow = $footerBaseRow + $idx;

            if ($tplRow !== $destRow) {

                $copyRow($sheet, $tplRow, $destRow, $colsRange);

                foreach ($dataColumns as $col) {

                    $sheet->setCellValue($col.$tplRow, '');

                }

            }

            $footerBaseRows[] = $destRow;

        }


        $totalRows   = $rows->count();

        $pagesNeeded = max(1, (int) ceil($totalRows / $perPage));


        for ($pageIdx = 0; $pageIdx < $pagesNeeded; $pageIdx++) {

            $offset = $pageIdx * $blockHeight;

            for ($r = 0; $r < $perPage; $r++) {

                $rowNumber = $firstData + $offset + $r;

                foreach ($dataColumns as $col) {

                    $sheet->setCellValue($col.$rowNumber, '');

                }

            }

        }


        // ---- Datos paginados

        $i = 0;

        foreach ($rows as $rec) {

            $page      = intdiv($i, $perPage);

            $indexInPg = $i % $perPage;

            $offset    = $page * $blockHeight;

            $headerDst = $headerRow + $offset;

            $dstRow    = $firstData + $offset + $indexInPg;


            if ($indexInPg === 0 && $page > 0) {

                $copyRow($sheet, $headerRow, $headerDst, $colsRange);

                foreach ($footerBaseRows as $srcFooterRow) {

                    $copyRow($sheet, $srcFooterRow, $srcFooterRow + $offset, $colsRange);

                }

            }

            $copyRow($sheet, $firstData, $dstRow, $colsRange);


            $lim = $rec->limites_aceptables;
            if ($lim === null) {
                $lim = $emByLoc[$rec->id_localizacion] ?? null;
            }

            $sheet->setCellValue($colNo     . $dstRow, $indexInPg + 1);
            $sheet->setCellValue($colArea   . $dstRow, (string) ($rec->area_nombre ?? ''));
            $sheet->setCellValue($colPunto  . $dstRow, (string) ($rec->punto_medicion ?? ''));
            $sheet->setCellValue($colPuesto . $dstRow, (string) ($rec->puesto_nombre ?? ''));
            $sheet->setCellValue($colProm   . $dstRow, is_null($rec->promedio) ? '' : (float)$rec->promedio);
            $sheet->setCellValue($colLim    . $dstRow, is_null($lim) ? '' : (float)$lim);
            $sheet->setCellValue($colObs    . $dstRow, (string) ($rec->observaciones ?? ''));
            $i++;
        }
    }

    if ($tpl = $spreadsheet->getSheetByName('Plantilla')) {
        $idx = $spreadsheet->getIndex($tpl);
        $spreadsheet->removeSheetByIndex($idx);
    }

    while (ob_get_level()) { ob_end_clean(); }
    $filename = "reporte_iluminacion_departamentos_{$year}.xlsx";

    return response()->streamDownload(function () use ($spreadsheet) {
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
    }, $filename, [
        'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
        'Pragma'                    => 'no-cache',
        'Expires'                   => '0',
        'Content-Transfer-Encoding' => 'binary',
    ]);
}

public function exportRuidoDesdePlantilla(Request $request)
{
    $year = $request->integer('year') ?: (int) now()->year;

    $tplPath = storage_path('app/public/medicion_ruido.xlsx');
    if (!is_file($tplPath)) {
        abort(404, 'No se encontr� la plantilla medicion_ruido.xlsx en storage/app/public');
    }
    $spreadsheet = IOFactory::load($tplPath);
    /** @var Worksheet $tplSheet */
    $tplSheet = $spreadsheet->getSheetByName('Plantilla') ?? $spreadsheet->getActiveSheet();

    $headerRows     = [11, 12, 13];
    $headerRowStart = $headerRows[0];
    $headerRowEnd   = $headerRows[count($headerRows) - 1];
    $headerHeight   = $headerRowEnd - $headerRowStart + 1;
    $firstData      = 14;
    $perPage        = 30;
    $colsRange      = 'A:J';

    $structureSheet = clone $tplSheet;

    $structureBook = new Spreadsheet();
    $structureBook->removeSheetByIndex(0);
    $structureBook->addExternalSheet($structureSheet);


    $dataColumns = ['A','B','C','D','E','F','G','H','I','J'];


    $mergeCache = [];


    $copyRow = function(Worksheet $srcSheet, Worksheet $dstSheet, int $srcRow, int $dstRow, string $rangeCols) use (&$mergeCache) {

        [$colStart, $colEnd] = explode(':', $rangeCols);

        for ($col = $colStart; $col <= $colEnd; $col++) {

            $src = $col.$srcRow;

            $dst = $col.$dstRow;


            $cell = $srcSheet->getCell($src);

            $dstSheet->setCellValueExplicit($dst, $cell->getValue(), $cell->getDataType());

            $dstSheet->duplicateStyle($srcSheet->getStyle($src), $dst);

            $dstSheet->getColumnDimension($col)->setWidth(

                $srcSheet->getColumnDimension($col)->getWidth()

            );

        }


        $dstSheet->getRowDimension($dstRow)->setRowHeight(

            $srcSheet->getRowDimension($srcRow)->getRowHeight()

        );


        $rowOffset = $dstRow - $srcRow;

        if ($rowOffset === 0) {

            return;

        }


        $sheetKey = spl_object_hash($dstSheet);

        if (!isset($mergeCache[$sheetKey])) {

            $mergeCache[$sheetKey] = [];

        }


        foreach ($srcSheet->getMergeCells() as $merge) {

            [$mStart, $mEnd] = explode(':', $merge);

            if (!preg_match('/([A-Z]+)(\\d+)/', $mStart, $m1) || !preg_match('/([A-Z]+)(\\d+)/', $mEnd, $m2)) {

                continue;

            }


            $startCol = $m1[1];

            $startRow = (int) $m1[2];

            $endCol   = $m2[1];

            $endRow   = (int) $m2[2];


            if ($srcRow < $startRow || $srcRow > $endRow) {

                continue;

            }


            $dstStartRow = $startRow + $rowOffset;

            $dstEndRow   = $endRow + $rowOffset;

            $dstRange    = $startCol.$dstStartRow.':'.$endCol.$dstEndRow;


            if (isset($mergeCache[$sheetKey][$dstRange])) {

                continue;

            }


            $dstSheet->mergeCells($dstRange);

            $mergeCache[$sheetKey][$dstRange] = true;

        }

    };


    $highestTplRow = $tplSheet->getHighestRow();
    $footerTemplateRows = [];
    $detectedFooterStart = null;
    for ($probe = $firstData; $probe <= $highestTplRow; $probe++) {
        $cellA = $tplSheet->getCell('A'.$probe)->getValue();
        $cellATrim = is_string($cellA) ? trim($cellA) : $cellA;
        $isFormula = is_string($cellATrim) && strlen($cellATrim) > 0 && $cellATrim[0] === '=';
        if (is_numeric($cellATrim) || $isFormula) {
            continue;
        }
        $rowHasContent = false;
        foreach ($dataColumns as $col) {
            $val = $tplSheet->getCell($col.$probe)->getValue();
            if ($val !== null && $val !== '') {
                $rowHasContent = true;
                break;
            }
        }
        if ($rowHasContent) {
            $detectedFooterStart = $probe;
            break;
        }
    }
    if ($detectedFooterStart === null) {
        $detectedFooterStart = $firstData + $perPage;
    }
    for ($probe = $detectedFooterStart; $probe <= $highestTplRow; $probe++) {
        $rowHasContent = false;
        foreach ($dataColumns as $col) {
            $val = $tplSheet->getCell($col.$probe)->getValue();
            if ($val !== null && $val !== '') {
                $rowHasContent = true;
                break;
            }
        }
        if (!$rowHasContent) {
            if (!empty($footerTemplateRows)) {
                break;
            }
            continue;
        }
        $footerTemplateRows[] = $probe;
    }
    if (!$footerTemplateRows) {
        $footerTemplateRows = [$firstData + $perPage];
    }

    $footerRowCount = count($footerTemplateRows);
    $footerBaseRow  = $firstData + $perPage;
    $blockHeight    = $headerHeight + $perPage + $footerRowCount;

    $base = DB::table('mediciones_ruido as m')
        ->leftJoin('localizacion as loc', 'loc.id_localizacion', '=', 'm.id_localizacion')
        ->leftJoin('puesto_trabajo_matriz as p', 'p.id_puesto_trabajo_matriz', '=', 'm.id_puesto_trabajo_matriz')
        ->leftJoin('area as a', 'a.id_area', '=', 'p.id_area')
        ->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) = ?', [$year])
        ->orderBy('m.departamento')
        ->orderBy('m.id_localizacion')
        ->orderBy('m.punto_medicion');

    $departamentos = (clone $base)->pluck('m.departamento')->filter()->unique()->values();

    $firstSheet = true;

    foreach ($departamentos as $dept) {
        $rows = (clone $base)
            ->where('m.departamento', $dept)
            ->select(
                'm.departamento',
                'm.nombre_observador',
                'm.instrumento',
                'm.serie',
                'm.marca',
                'm.nrr',
                'm.id_localizacion',
                'loc.localizacion as localizacion_nombre',
                'p.puesto_trabajo_matriz as puesto_nombre',
                'm.punto_medicion',
                'm.nivel_maximo',
                'm.nivel_minimo',
                'm.nivel_promedio',
                'm.limites_aceptables',
                'm.fecha_realizacion_inicio',
                'm.fecha_realizacion_final',
                'm.observaciones',
                'a.area as area_nombre'
            )
            ->get();

        if ($rows->isEmpty()) {
            continue;
        }

        $sheet = $firstSheet ? $tplSheet : $this->ensureSheet($spreadsheet, (string) $dept, $tplSheet);
        if ($firstSheet) {
            $sheet->setTitle($this->xlTitle((string) $dept));
            $firstSheet = false;
        }

        $minIni = $rows->min('fecha_realizacion_inicio');
        $maxFin = $rows->max('fecha_realizacion_final');

        $sheet->setCellValue('D4', (string) ($dept ?? ''));
        $sheet->setCellValue('D5', (string) ($rows->first()->nombre_observador ?? ''));

        $fechaRealizacion = '';
        if ($minIni && $maxFin) {
            $ini = (new \DateTime($minIni))->format('d/m/Y');
            $fin = (new \DateTime($maxFin))->format('d/m/Y');
            $fechaRealizacion = $minIni === $maxFin ? $ini : $ini.' al '.$fin;
        } elseif ($minIni) {
            $fechaRealizacion = (new \DateTime($minIni))->format('d/m/Y');
        } elseif ($maxFin) {
            $fechaRealizacion = (new \DateTime($maxFin))->format('d/m/Y');
        }
        $sheet->setCellValue('D6', $fechaRealizacion);

        $sheet->setCellValue('D7', (string) ($rows->first()->instrumento ?? ''));
        $sheet->setCellValue('B8', (string) ($rows->first()->serie ?? ''));
        $sheet->setCellValue('E8', (string) ($rows->first()->marca ?? ''));

        $headerNrr = null;
        foreach ($rows as $metaRow) {
            if ($metaRow->nrr !== null && $metaRow->nrr !== '') {
                $headerNrr = $metaRow->nrr;
                break;
            }
        }
        $sheet->setCellValue('B9', $headerNrr !== null ? $headerNrr : '');

        foreach ($footerTemplateRows as $idx => $tplRow) {

            $destRow = $footerBaseRow + $idx;

            $copyRow($structureSheet, $sheet, $tplRow, $destRow, $colsRange);

        }


        $totalRows   = $rows->count();
        $pagesNeeded = max(1, (int) ceil($totalRows / $perPage));

        for ($pageIdx = 0; $pageIdx < $pagesNeeded; $pageIdx++) {
            $offset = $pageIdx * $blockHeight;
            for ($r = 0; $r < $perPage; $r++) {
                $rowNumber = $firstData + $offset + $r;
                foreach ($dataColumns as $col) {
                    $sheet->setCellValue($col.$rowNumber, '');
                }
            }
        }

        $i = 0;
        foreach ($rows as $rec) {
            $page      = intdiv($i, $perPage);
            $indexInPg = $i % $perPage;
            $offset    = $page * $blockHeight;
            $dstRow    = $firstData + $offset + $indexInPg;

            if ($indexInPg === 0 && $page > 0) {

                foreach ($headerRows as $srcHeaderRow) {

                    $copyRow($structureSheet, $sheet, $srcHeaderRow, $srcHeaderRow + $offset, $colsRange);

                }


                for ($tplDataIdx = 0; $tplDataIdx < $perPage; $tplDataIdx++) {

                    $srcDataRow  = $firstData + $tplDataIdx;

                    $destDataRow = $srcDataRow + $offset;


                    $copyRow($structureSheet, $sheet, $srcDataRow, $destDataRow, $colsRange);


                    foreach ($dataColumns as $col) {

                        $sheet->setCellValue($col.$destDataRow, '');

                    }

                }


                foreach ($footerTemplateRows as $tplRow) {

                    $copyRow($structureSheet, $sheet, $tplRow, $tplRow + $offset, $colsRange);

                }

            }


            $templateRow = $firstData + $indexInPg;

            $copyRow($structureSheet, $sheet, $templateRow, $dstRow, $colsRange);

            $sheet->setCellValue('A' . $dstRow, $indexInPg + 1);
            $sheet->setCellValue('B' . $dstRow, (string) ($rec->localizacion_nombre ?? ''));
            $sheet->setCellValue('C' . $dstRow, (string) ($rec->punto_medicion ?? ''));
            $sheet->setCellValue('D' . $dstRow, (string) ($rec->puesto_nombre ?? ''));
            $sheet->setCellValue('E' . $dstRow, $max === null ? '' : $max);
            $sheet->setCellValue('F' . $dstRow, $min === null ? '' : $min);
            $sheet->setCellValue('G' . $dstRow, $prm === null ? '' : round($prm, 2));
            $sheet->setCellValue('H' . $dstRow, $nre === null ? '' : round($nre, 2));
            $sheet->setCellValue('I' . $dstRow, $lim === null ? '' : $lim);
            $sheet->setCellValue('J' . $dstRow, (string) ($rec->observaciones ?? ''));

            $i++;
        }
    }

    while (ob_get_level()) { ob_end_clean(); }
    $filename = "reporte_ruido_departamentos_{$year}.xlsx";

    return response()->streamDownload(function () use ($spreadsheet) {
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
    }, $filename, [
        'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
        'Pragma'                    => 'no-cache',
        'Expires'                   => '0',
        'Content-Transfer-Encoding' => 'binary',
    ]);
}

public function prefill(Request $request)
{
    $id = (int) $request->query('id_localizacion');
    if ($id <= 0) {
        return response()->json(['ok' => false, 'error' => 'id_localizacion requerido'], 422);
    }

    // ---------- RUIDO: lote más reciente por (fecha_inicio, fecha_fin) ----------
    $lastR = DB::table('mediciones_ruido')
        ->where('id_localizacion', $id)
        ->orderByRaw('COALESCE(fecha_realizacion_inicio, fecha_realizacion_final) DESC')
        ->orderByDesc('id_mediciones_ruido')
        ->first([
            'departamento',
            'nombre_observador',
            'instrumento',
            'serie',
            'marca',
            'nrr',
            'fecha_realizacion_inicio',
            'fecha_realizacion_final',
        ]);

    $ruidoRows = collect();
    if ($lastR) {
        $ruidoRows = DB::table('mediciones_ruido as m')
            ->leftJoin('puesto_trabajo_matriz as p', 'p.id_puesto_trabajo_matriz', '=', 'm.id_puesto_trabajo_matriz')
            ->where('m.id_localizacion', $id)
            ->where(function ($qq) use ($lastR) {
                if (!is_null($lastR->fecha_realizacion_inicio)) {
                    $qq->where('m.fecha_realizacion_inicio', $lastR->fecha_realizacion_inicio);
                } else {
                    $qq->whereNull('m.fecha_realizacion_inicio');
                }
                if (!is_null($lastR->fecha_realizacion_final)) {
                    $qq->where('m.fecha_realizacion_final', $lastR->fecha_realizacion_final);
                } else {
                    $qq->whereNull('m.fecha_realizacion_final');
                }
            })
            ->select(
                'm.id_mediciones_ruido as id_row',
                'm.id_puesto_trabajo_matriz',
                'm.punto_medicion',
                'm.nivel_maximo',
                'm.nivel_minimo',
                'm.nivel_promedio',
                'm.nre',
                'm.limites_aceptables',
                'm.observaciones'
            )
            ->orderBy('m.punto_medicion')
            ->orderBy('m.id_puesto_trabajo_matriz')
            ->get();
    }

    // ---------- ILUMINACIÓN: lote más reciente por (fecha_inicio, fecha_fin) ----------
    $lastL = DB::table('mediciones_iluminacion')
        ->where('id_localizacion', $id)
        ->orderByRaw('COALESCE(fecha_realizacion_inicio, fecha_realizacion_final) DESC')
        ->orderByDesc('id')
        ->first([
            'departamento',
            'nombre_observador',
            'instrumento',
            'serie',
            'marca',
            'fecha_realizacion_inicio',
            'fecha_realizacion_final',
        ]);

    $luxRows = collect();
    if ($lastL) {
        $luxRows = DB::table('mediciones_iluminacion as m')
            ->leftJoin('puesto_trabajo_matriz as p', 'p.id_puesto_trabajo_matriz', '=', 'm.id_puesto_trabajo_matriz')
            ->where('m.id_localizacion', $id)
            ->where(function ($qq) use ($lastL) {
                if (!is_null($lastL->fecha_realizacion_inicio)) {
                    $qq->where('m.fecha_realizacion_inicio', $lastL->fecha_realizacion_inicio);
                } else {
                    $qq->whereNull('m.fecha_realizacion_inicio');
                }
                if (!is_null($lastL->fecha_realizacion_final)) {
                    $qq->where('m.fecha_realizacion_final', $lastL->fecha_realizacion_final);
                } else {
                    $qq->whereNull('m.fecha_realizacion_final');
                }
            })
            ->select(
                'm.id as id_row',
                'm.id_puesto_trabajo_matriz',
                'm.punto_medicion',
                'm.promedio',
                'm.limites_aceptables',
                'm.observaciones'
            )
            ->orderBy('m.punto_medicion')
            ->orderBy('m.id_puesto_trabajo_matriz')
            ->get();
    }

    return response()->json([
        'ok' => true,

        'ruido_meta' => $lastR ? [
            'departamento'       => $lastR->departamento,
            'nombre_observador'  => $lastR->nombre_observador,
            'instrumento'        => $lastR->instrumento,
            'serie'              => $lastR->serie,
            'marca'              => $lastR->marca,
            'nrr'                => $lastR->nrr,
            'fecha_inicio'       => $lastR->fecha_realizacion_inicio,
            'fecha_final'        => $lastR->fecha_realizacion_final,
        ] : null,
        'ruido_rows' => $ruidoRows,

        'lux_meta' => $lastL ? [
            'departamento'       => $lastL->departamento,
            'nombre_observador'  => $lastL->nombre_observador,
            'instrumento'        => $lastL->instrumento,
            'serie'              => $lastL->serie,
            'marca'              => $lastL->marca,
            'fecha_inicio'       => $lastL->fecha_realizacion_inicio,
            'fecha_final'        => $lastL->fecha_realizacion_final,
        ] : null,
        'lux_rows' => $luxRows,
    ]);
}
}

