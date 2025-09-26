<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Illuminate\Support\Facades\Schema;

class VerificacionRiesgosController extends Controller
{
    public function index(Request $request)
    {
        // ======= Parámetros de UI =======
        $mode   = $request->input('mode', 'verificacion'); // verificacion | buscar
        $by     = $request->input('by', 'puesto');         // puesto|riesgo|epp|capacitacion|senalizacion|otras
        $id     = (int) $request->input('id', 0);          // id del riesgo/epp/etc cuando by != puesto
        $soloSi = (bool) $request->boolean('solo_si', $mode==='buscar' && $by==='puesto');

        // ======= Listas para selects =======
        $puestosRaw = DB::table('puesto_trabajo_matriz as ptm')
            ->select('ptm.id_puesto_trabajo_matriz as id', 'ptm.puesto_trabajo_matriz as nombre')
            ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado', '!=', 0); })
            ->orderBy('ptm.puesto_trabajo_matriz')
            ->get();

        $puestos = $puestosRaw->map(fn($p) => ['id'=>$p->id, 'nombre'=>$p->nombre])->all();

        $riesgosLista = DB::table('riesgo as r')
            ->join('tipo_riesgo as tr','tr.id_tipo_riesgo','=','r.id_tipo_riesgo')
            ->select('r.id_riesgo as id', DB::raw("CONCAT(tr.tipo_riesgo,' | ', r.nombre_riesgo) as label"))
            ->orderBy('tr.tipo_riesgo')->orderBy('r.nombre_riesgo')->get()
            ->map(fn($x)=>['id'=>$x->id,'label'=>$x->label])->all();

        $eppLista = DB::table('epp')
            ->select('id_epp as id', DB::raw("COALESCE(NULLIF(TRIM(equipo),''), CONCAT_WS(' ', NULLIF(TRIM(marca),''), NULLIF(TRIM(codigo),''))) as label"))
            ->orderBy('label')->get()->map(fn($x)=>['id'=>$x->id,'label'=>$x->label])->all();

        $capacitacionLista = DB::table('capacitacion')
            ->select('id_capacitacion as id', DB::raw("COALESCE(NULLIF(TRIM(capacitacion),''), CONCAT('Capacitacion #', id_capacitacion)) as label"))
            ->orderBy('label')->get()->map(fn($x)=>['id'=>$x->id,'label'=>$x->label])->all();

        $senalizacionLista = DB::table('senalizacion')
            ->select('id_senalizacion as id', DB::raw("COALESCE(NULLIF(TRIM(senalizacion),''), CONCAT('Señalización #', id_senalizacion)) as label"))
            ->orderBy('label')->get()->map(fn($x)=>['id'=>$x->id,'label'=>$x->label])->all();

        $otrasLista = DB::table('otras_medidas')
            ->select('id_otras_medidas as id', DB::raw("COALESCE(NULLIF(TRIM(otras_medidas),''), CONCAT('Medida #', id_otras_medidas)) as label"))
            ->orderBy('label')->get()->map(fn($x)=>['id'=>$x->id,'label'=>$x->label])->all();

        // ======= Resolver puesto seleccionado (para verificación y buscar por puesto) =======
        $puestoId = (int) $request->input('puesto', 0);
        if (($mode === 'verificacion' || ($mode==='buscar' && $by==='puesto')) && $puestoId === 0 && !empty($puestos)) {
            $puestoId = (int) $puestos[0]['id'];
        }

        // ======= Variables de salida =======
        $grupos = [];            // para tarjetas de verificación
        $resultados = [];        // para tabla de resultados
        $criterioNombre = null;  // etiqueta del criterio buscado

        // ======= Helpers =======
        $esSiSql = "LOWER(TRIM(rv.valor)) IN ('sí','si','1')";

        // ======= VERIFICACIÓN (y BUSCAR por PUESTO) =======
        if ($mode === 'verificacion' || ($mode==='buscar' && $by==='puesto')) {

            if (empty($puestos)) {
                // No hay puestos: render vacío
                return view('riesgos.verificacion', compact(
                    'puestos','puestoId','grupos','mode','by','id','soloSi',
                    'riesgosLista','eppLista','capacitacionLista','senalizacionLista','otrasLista'
                ))->with([
                    'puestoSeleccionado' => null,
                ]);
            }

            // Traer TODOS los riesgos y valor del puesto (default "No")
            $riesgos = DB::table('riesgo as r')
                ->join('tipo_riesgo as tr', 'tr.id_tipo_riesgo', '=', 'r.id_tipo_riesgo')
                ->leftJoin('riesgo_valor as rv', function ($join) use ($puestoId) {
                    $join->on('rv.id_riesgo', '=', 'r.id_riesgo')
                         ->where('rv.id_puesto_trabajo_matriz', '=', $puestoId);
                })
                ->select([
                    'tr.tipo_riesgo as tipo',
                    'r.id_riesgo',
                    'r.nombre_riesgo as riesgo',
                    DB::raw("COALESCE(NULLIF(TRIM(rv.valor), ''), 'No') as valor"),
                ])
                ->orderBy('tr.tipo_riesgo')
                ->orderBy('r.nombre_riesgo')
                ->get();

            // Estructura por tipo
            foreach ($riesgos as $row) {
                $tipo = $row->tipo ?: 'SIN TIPO';
                $grupos[$tipo]['riesgos'][] = [
                    'id_riesgo' => (int) $row->id_riesgo,
                    'nombre'    => $row->riesgo,
                    'valor'     => $row->valor,
                ];
            }

            // IDs de riesgos "Sí" (para medidas)
            $idsSi = collect($riesgos)->filter(function ($r) {
                    $v = mb_strtolower(trim((string)$r->valor));
                    return in_array($v, ['sí','si','1'], true);
                })
                ->pluck('id_riesgo')->unique()->values();

            // Inicializar medidas vacías
            foreach ($grupos as $tipo => &$g) {
                $g['medidas'] = ['epp'=>[], 'capacitacion'=>[], 'senalizacion'=>[], 'otras'=>[]];
            } unset($g);

            if ($idsSi->isNotEmpty()) {
                // Cargar medidas por riesgo (catálogo)
                $epp = DB::table('medidas_riesgo_puesto as mrp')
                    ->join('epp as e', 'e.id_epp', '=', 'mrp.id_epp')
                    ->whereIn('mrp.id_riesgo', $idsSi)
                    ->select('mrp.id_riesgo', DB::raw('COALESCE(NULLIF(TRIM(e.equipo),""), CONCAT_WS(" ", NULLIF(TRIM(e.marca),""), NULLIF(TRIM(e.codigo),""))) as nombre'))
                    ->get();

                $cap = DB::table('medidas_riesgo_puesto as mrp')
                    ->join('capacitacion as c', 'c.id_capacitacion', '=', 'mrp.id_capacitacion')
                    ->whereIn('mrp.id_riesgo', $idsSi)
                    ->select('mrp.id_riesgo', DB::raw('COALESCE(NULLIF(TRIM(c.capacitacion),""), CONCAT("Capacitacion #", c.id_capacitacion)) as nombre'))
                    ->get();

                $sen = DB::table('medidas_riesgo_puesto as mrp')
                    ->join('senalizacion as s', 's.id_senalizacion', '=', 'mrp.id_senalizacion')
                    ->whereIn('mrp.id_riesgo', $idsSi)
                    ->select('mrp.id_riesgo', DB::raw('COALESCE(NULLIF(TRIM(s.senalizacion),""), CONCAT("Señalización #", s.id_senalizacion)) as nombre'))
                    ->get();

                $otr = DB::table('medidas_riesgo_puesto as mrp')
                    ->join('otras_medidas as o', 'o.id_otras_medidas', '=', 'mrp.id_otras_medidas')
                    ->whereIn('mrp.id_riesgo', $idsSi)
                    ->select('mrp.id_riesgo', DB::raw('COALESCE(NULLIF(TRIM(o.otras_medidas),""), CONCAT("Medida #", o.id_otras_medidas)) as nombre'))
                    ->get();

                $map = function ($rows) {
                    $m = [];
                    foreach ($rows as $r) { $m[$r->id_riesgo][] = $r->nombre; }
                    return $m;
                };
                $mEpp = $map($epp); $mCap = $map($cap); $mSen = $map($sen); $mOtr = $map($otr);

                // Agregar medidas únicas por categoría (solo de riesgos "Sí")
                foreach ($grupos as $tipo => &$grupo) {
                    $idsTipoSi = collect($grupo['riesgos'])
                        ->filter(function ($r) { $v = mb_strtolower(trim((string)$r['valor'])); return in_array($v, ['sí','si','1'], true); })
                        ->pluck('id_riesgo');

                    $grupo['medidas'] = [
                        'epp'          => collect($idsTipoSi)->flatMap(fn($id)=>$mEpp[$id] ?? [])->unique()->values()->all(),
                        'capacitacion' => collect($idsTipoSi)->flatMap(fn($id)=>$mCap[$id] ?? [])->unique()->values()->all(),
                        'senalizacion' => collect($idsTipoSi)->flatMap(fn($id)=>$mSen[$id] ?? [])->unique()->values()->all(),
                        'otras'        => collect($idsTipoSi)->flatMap(fn($id)=>$mOtr[$id] ?? [])->unique()->values()->all(),
                    ];
                } unset($grupo);
            }

            // Para la vista
            return view('riesgos.verificacion', [
                'puestos'            => $puestos,
                'puestoSeleccionado' => $puestoId,
                'grupos'             => $grupos,
                'mode'               => $mode,
                'by'                 => $by,
                'idSeleccion'        => $id,
                'soloSi'             => $soloSi,
                'riesgosLista'       => $riesgosLista,
                'eppLista'           => $eppLista,
                'capacitacionLista'  => $capacitacionLista,
                'senalizacionLista'  => $senalizacionLista,
                'otrasLista'         => $otrasLista,
            ]);
        }

        // ======= BUSCAR por RIESGO: lista de puestos que lo tienen en "Sí" =======
        if ($mode === 'buscar' && $by === 'riesgo' && $id > 0) {
            $rInfo = DB::table('riesgo as r')->join('tipo_riesgo as tr','tr.id_tipo_riesgo','=','r.id_tipo_riesgo')
                ->where('r.id_riesgo',$id)->select(DB::raw("CONCAT(tr.tipo_riesgo,' | ', r.nombre_riesgo) as label"))->first();
            $criterioNombre = $rInfo->label ?? null;

            $resultados = DB::table('puesto_trabajo_matriz as ptm')
                ->join('riesgo_valor as rv', function ($join) use ($id) {
                    $join->on('rv.id_puesto_trabajo_matriz','=','ptm.id_puesto_trabajo_matriz')
                         ->where('rv.id_riesgo','=',$id);
                })
                ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado','!=',0); })
                ->whereRaw($esSiSql)
                ->leftJoin('riesgo as r','r.id_riesgo','=','rv.id_riesgo')
                ->leftJoin('tipo_riesgo as tr','tr.id_tipo_riesgo','=','r.id_tipo_riesgo')
                ->groupBy('ptm.id_puesto_trabajo_matriz','ptm.puesto_trabajo_matriz')
                ->orderBy('ptm.puesto_trabajo_matriz')
                ->select([
                    'ptm.id_puesto_trabajo_matriz as id',
                    'ptm.puesto_trabajo_matriz as puesto',
                    DB::raw("GROUP_CONCAT(DISTINCT CONCAT(tr.tipo_riesgo,' | ', r.nombre_riesgo) ORDER BY tr.tipo_riesgo, r.nombre_riesgo SEPARATOR ', ') as riesgos"),
                ])
                ->get()
                ->map(fn($x)=>['id'=>$x->id,'puesto'=>$x->puesto,'riesgos'=>$x->riesgos])->all();
        }

        // ======= BUSCAR por MEDIDAS: puestos que requieren la medida (porque algún riesgo 'Sí' la dispara) =======
        $mapMedida = [
            'epp'          => ['col' => 'id_epp',          'tabla' => 'epp',            'pk'=>'id_epp',           'nombre'=>"COALESCE(NULLIF(TRIM(equipo),''), CONCAT_WS(' ', NULLIF(TRIM(marca),''), NULLIF(TRIM(codigo),'')))"],
            'capacitacion' => ['col' => 'id_capacitacion', 'tabla' => 'capacitacion',   'pk'=>'id_capacitacion',  'nombre'=>"COALESCE(NULLIF(TRIM(capacitacion),''), CONCAT('Capacitacion #', id_capacitacion))"],
            'senalizacion' => ['col' => 'id_senalizacion', 'tabla' => 'senalizacion',   'pk'=>'id_senalizacion',  'nombre'=>"COALESCE(NULLIF(TRIM(senalizacion),''), CONCAT('Señalización #', id_senalizacion))"],
            'otras'        => ['col' => 'id_otras_medidas','tabla' => 'otras_medidas',  'pk'=>'id_otras_medidas', 'nombre'=>"COALESCE(NULLIF(TRIM(otras_medidas),''), CONCAT('Medida #', id_otras_medidas))"],
        ];

        if ($mode==='buscar' && isset($mapMedida[$by]) && $id > 0) {
            $info = $mapMedida[$by];

            // nombre del criterio
            $criterioRow = DB::table($info['tabla'])
                ->select(DB::raw($info['nombre'].' as label'))
                ->where($info['pk'], $id)->first();
            $criterioNombre = $criterioRow->label ?? null;

            // puestos que requieren esa medida
            $resultados = DB::table('puesto_trabajo_matriz as ptm')
                ->join('riesgo_valor as rv', 'rv.id_puesto_trabajo_matriz','=','ptm.id_puesto_trabajo_matriz')
                ->join('medidas_riesgo_puesto as mrp', 'mrp.id_riesgo','=','rv.id_riesgo')
                ->join('riesgo as r','r.id_riesgo','=','rv.id_riesgo')
                ->join('tipo_riesgo as tr','tr.id_tipo_riesgo','=','r.id_tipo_riesgo')
                ->where("mrp.{$info['col']}", $id)
                ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado','!=',0); })
                ->whereRaw($esSiSql)
                ->groupBy('ptm.id_puesto_trabajo_matriz','ptm.puesto_trabajo_matriz')
                ->orderBy('ptm.puesto_trabajo_matriz')
                ->select([
                    'ptm.id_puesto_trabajo_matriz as id',
                    'ptm.puesto_trabajo_matriz as puesto',
                    DB::raw("GROUP_CONCAT(DISTINCT CONCAT(tr.tipo_riesgo,' | ', r.nombre_riesgo) ORDER BY tr.tipo_riesgo, r.nombre_riesgo SEPARATOR ', ') as riesgos"),
                ])
                ->get()
                ->map(fn($x)=>['id'=>$x->id,'puesto'=>$x->puesto,'riesgos'=>$x->riesgos])->all();
        }

        // Render de BUSCAR (tabla)
        return view('riesgos.verificacion', [
            'puestos'            => $puestos,
            'puestoSeleccionado' => null, // no aplica en tabla
            'grupos'             => [],   // no se usan en tabla
            'mode'               => $mode,
            'by'                 => $by,
            'idSeleccion'        => $id,
            'soloSi'             => $soloSi,
            'resultados'         => $resultados,
            'criterioNombre'     => $criterioNombre,
            'riesgosLista'       => $riesgosLista,
            'eppLista'           => $eppLista,
            'capacitacionLista'  => $capacitacionLista,
            'senalizacionLista'  => $senalizacionLista,
            'otrasLista'         => $otrasLista,
        ]);
    }

public function exportPlanAccion(Request $request)
{
    // ========= 0) Parámetros y plantilla =========
    $anio = intval($request->input('anio', date('Y')));

    $tplPath = storage_path('app/public/formato_plan_de_accion.xlsx');
    if (!is_file($tplPath)) {
        return back()->with('error', 'No se encontró la plantilla formato_plan_de_accion.xlsx');
    }

    try { $spreadsheet = IOFactory::load($tplPath); }
    catch (\Throwable $e) {
        return back()->with('error', 'No se pudo abrir la plantilla: '.$e->getMessage());
    }
    $sheet = $spreadsheet->getActiveSheet();

    // ========= 1) Normalizador y catálogo de riesgos =========
    $norm = function (string $s): string {
        $s = trim(mb_strtoupper($s, 'UTF-8'));
        $s = @iconv('UTF-8', 'ASCII//TRANSLIT//IGNORE', $s) ?: $s;
        $s = preg_replace('/[^A-Z0-9 ]+/', ' ', $s);
        $s = preg_replace('/\s+/', ' ', $s);
        return str_replace(' ', '', $s);
    };

    $riesgosDb = DB::table('riesgo')->select('id_riesgo','nombre_riesgo')->get();
    $mapRiesgoId = [];
    foreach ($riesgosDb as $r) {
        $k = $norm((string)($r->nombre_riesgo ?? ''));
        if ($k !== '') $mapRiesgoId[$k] = (int)$r->id_riesgo;
    }
    $riesgoKeys = array_keys($mapRiesgoId);

    $findClosestId = function (string $needle) use ($mapRiesgoId, $riesgoKeys) {
        if (isset($mapRiesgoId[$needle])) return $mapRiesgoId[$needle];
        foreach ($riesgoKeys as $k) {
            if (str_contains($k, $needle) || str_contains($needle, $k)) return $mapRiesgoId[$k];
        }
        $bestK=null; $bestD=PHP_INT_MAX; $bestLen=0;
        foreach ($riesgoKeys as $k) {
            $d = levenshtein($needle, $k);
            if ($d < $bestD) { $bestD = $d; $bestK = $k; $bestLen = max(strlen($needle), strlen($k)); }
        }
        if ($bestK !== null) {
            $thr = max(3, (int)ceil($bestLen * 0.25));
            if ($bestD <= $thr) return $mapRiesgoId[$bestK];
        }
        return null;
    };

    // ========= 2) Puestos y “sí” por riesgo =========
    $puestos = DB::table('puesto_trabajo_matriz as ptm')
        ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado','!=',0); })
        ->select('ptm.id_puesto_trabajo_matriz as id','ptm.puesto_trabajo_matriz as nombre')
        ->orderBy('ptm.puesto_trabajo_matriz')
        ->get();

    $esSiSql = "LOWER(TRIM(rv.valor)) IN ('si','s','sí','s\xC3\xAD','1')";
    $puestosPorRiesgo = [];
    $rvRows = DB::table('riesgo_valor as rv')
        ->join('puesto_trabajo_matriz as ptm','ptm.id_puesto_trabajo_matriz','=','rv.id_puesto_trabajo_matriz')
        ->whereRaw($esSiSql)
        ->select('rv.id_riesgo','rv.id_puesto_trabajo_matriz','ptm.puesto_trabajo_matriz as nombre')
        ->get();
    foreach ($rvRows as $row) {
        $rid = (int)$row->id_riesgo;
        if (!isset($puestosPorRiesgo[$rid])) $puestosPorRiesgo[$rid] = [];
        $puestosPorRiesgo[$rid][$row->id_puesto_trabajo_matriz] = (string)$row->nombre;
    }

    // ========= 2.1) Puestos con QUÍMICOS (excluye “NINGUNO” y “No manipula…”) =========
    // Tomamos el nombre del químico desde la tabla quimico (varios posibles campos) y filtramos en PHP usando el normalizador.
    $excluir = [
        $norm('NINGUNO'),
        $norm('No manipula productos químicos solo recibe los del área asignada'),
    ];
    $qpRows = DB::table('quimico_puesto as qp')
        ->join('puesto_trabajo_matriz as ptm','ptm.id_puesto_trabajo_matriz','=','qp.id_puesto_trabajo_matriz')
        ->leftJoin('quimico as q','q.id_quimico','=','qp.id_quimico')
        ->whereNotNull('qp.id_quimico')
        ->orderBy('ptm.puesto_trabajo_matriz')
        ->select([
            'qp.id_puesto_trabajo_matriz',
            'ptm.puesto_trabajo_matriz as nombre',
            // Intento robusto: distintos posibles nombres de columna en quimico
            DB::raw("COALESCE(NULLIF(TRIM(q.nombre_comercial), ''), 'NINGUNO') as quimico")
        ])
        ->get();

    // Quedarnos solo con puestos que tengan AL MENOS un químico no excluido
    $puestosQuimicoMap = [];
    foreach ($qpRows as $r) {
        $qNorm = $norm((string)$r->quimico);
        if (!in_array($qNorm, $excluir, true)) {
            $puestosQuimicoMap[$r->id_puesto_trabajo_matriz] = (string)$r->nombre;
        }
    }
    // Array plano y ordenado por nombre
    $puestosQuimico = array_values($puestosQuimicoMap);
    sort($puestosQuimico, SORT_NATURAL|SORT_FLAG_CASE);

    // ========= 3) Medidas por riesgo (incluye IDs para buscar fechas) =========
    $medidasPorRiesgo = [];
    $mrpRows = DB::table('medidas_riesgo_puesto as mrp')
        ->leftJoin('epp as epp', 'epp.id_epp', '=', 'mrp.id_epp')
        ->leftJoin('capacitacion as c', 'c.id_capacitacion', '=', 'mrp.id_capacitacion')
        ->leftJoin('senalizacion as s', 's.id_senalizacion', '=', 'mrp.id_senalizacion')
        ->leftJoin('otras_medidas as o', 'o.id_otras_medidas', '=', 'mrp.id_otras_medidas')
        ->select(
            'mrp.id_riesgo',
            'mrp.id_epp', 'epp.equipo as epp',
            'mrp.id_capacitacion', 'c.capacitacion as cap',
            'mrp.id_senalizacion', 's.senalizacion as senal',
            'mrp.id_otras_medidas', 'o.otras_medidas as otras'
        )
        ->get();

    foreach ($mrpRows as $r) {
        $rid = (int)$r->id_riesgo;
        if (!isset($medidasPorRiesgo[$rid])) {
            $medidasPorRiesgo[$rid] = [
                'epp'=>[], 'cap'=>[], 'senal'=>[], 'otras'=>[],
                'epp_ids'=>[], 'cap_ids'=>[]
            ];
        }
        $val = function($x){ $x = trim((string)$x); return $x !== '' ? $x : null; };
        if (($v=$val($r->epp))   !== null) { $medidasPorRiesgo[$rid]['epp'][$v] = true; $medidasPorRiesgo[$rid]['epp_ids'][$v] = (int)$r->id_epp; }
        if (($v=$val($r->cap))   !== null) { $medidasPorRiesgo[$rid]['cap'][$v] = true; $medidasPorRiesgo[$rid]['cap_ids'][$v] = (int)$r->id_capacitacion; }
        if (($v=$val($r->senal)) !== null) { $medidasPorRiesgo[$rid]['senal'][$v] = true; }
        if (($v=$val($r->otras)) !== null) { $medidasPorRiesgo[$rid]['otras'][$v] = true; }
    }

    // ========= 4) Helpers para fechas =========

    // Para EPP: tablas candidatas (ajusta si usas otras)
    $eppDateTables = [
        ['table' => 'asignacion_epp',   'id' => 'id_epp'],
        ['table' => 'epp_entrega',      'id' => 'id_epp'],
        ['table' => 'empleado_epp',     'id' => 'id_epp'],
        ['table' => 'asignaciones_epp', 'id' => 'id_epp'],
        ['table' => 'epp_empleado',     'id' => 'id_epp'],
    ];

    // Detecta TODAS las columnas DATE/DATETIME/TIMESTAMP
    $dbName = DB::getDatabaseName();
    $detectDateCols = function (string $table) use ($dbName) {
        if (!Schema::hasTable($table)) return [];
        $rows = DB::select("
            SELECT COLUMN_NAME as col
            FROM information_schema.columns
            WHERE table_schema = ? AND table_name = ?
              AND DATA_TYPE IN ('date','datetime','timestamp')
        ", [$dbName, $table]);
        return $rows ? array_map(fn($r) => $r->col, $rows) : [];
    };

    // Trae TODAS las fechas (sin límite) por ID desde todas las columnas fecha detectadas
    $fetchDatesMap = function(array $ids, array $tableSpecs) use ($anio, $detectDateCols) {
        $out = [];
        $ids = array_values(array_filter(array_map('intval', $ids), fn($v)=>$v>0));
        if (empty($ids)) return $out;

        foreach ($tableSpecs as $spec) {
            $table = $spec['table'];
            $idField = $spec['id'];
            if (!Schema::hasTable($table) || !Schema::hasColumn($table, $idField)) continue;

            $dateCols = $detectDateCols($table);
            if (empty($dateCols)) continue;

            foreach ($dateCols as $dateCol) {
                $rows = DB::table($table)
                    ->whereIn($idField, $ids)
                    ->whereYear($dateCol, $anio) // <-- elimina esta línea si quieres TODAS las anualidades
                    ->select([$idField.' as id', DB::raw("DATE(`{$dateCol}`) as d")])
                    ->get();

                foreach ($rows as $r) {
                    if (!$r->d) continue;
                    $id = (int)$r->id;
                    $out[$id][] = (string)$r->d;
                }
            }
        }

        foreach ($out as $id => $arr) {
            $arr = array_unique(array_filter($arr));
            sort($arr);
            $out[$id] = $arr;
        }
        return $out;
    };

    // Parseador SQL para VARCHAR -> DATE en asistencia_capacitacion.fecha_recibida
    $varcharDate = function(string $qualifiedCol) {
        return "COALESCE(
            STR_TO_DATE($qualifiedCol, '%Y-%m-%d'),
            STR_TO_DATE($qualifiedCol, '%d/%m/%Y'),
            STR_TO_DATE($qualifiedCol, '%d-%m-%Y'),
            STR_TO_DATE($qualifiedCol, '%m/%d/%Y'),
            STR_TO_DATE($qualifiedCol, '%m-%d-%Y'),
            STR_TO_DATE($qualifiedCol, '%Y/%m/%d'),
            STR_TO_DATE($qualifiedCol, '%Y.%m.%d')
        )";
    };

    // Fechas de CAPACITACIÓN (asistencia_capacitacion.fecha_recibida VARCHAR)
    // Relación: a.id_capacitacion_instructor -> capacitacion_instructor.id_capacitacion
    $fetchCapAttendanceDates = function(array $capIds) use ($anio, $varcharDate) {
        $map = [];
        $capIds = array_values(array_filter(array_map('intval', $capIds), fn($v)=>$v>0));
        if (empty($capIds) || !Schema::hasTable('asistencia_capacitacion')) return $map;

        $dateExpr = $varcharDate('a.fecha_recibida');

        if (Schema::hasTable('capacitacion_instructor')
            && Schema::hasColumn('capacitacion_instructor','id_capacitacion')
            && Schema::hasColumn('asistencia_capacitacion','id_capacitacion_instructor')) {

            $rows = DB::table('asistencia_capacitacion as a')
                ->join('capacitacion_instructor as ci','ci.id_capacitacion_instructor','=','a.id_capacitacion_instructor')
                ->whereIn('ci.id_capacitacion', $capIds)
                ->whereRaw("YEAR($dateExpr) = ?", [$anio]) // <-- elimina esta línea si quieres TODAS las anualidades
                ->select(DB::raw('ci.id_capacitacion as id'), DB::raw("DATE($dateExpr) as d"))
                ->get();

        } elseif (Schema::hasColumn('asistencia_capacitacion','id_capacitacion')) {
            $rows = DB::table('asistencia_capacitacion as a')
                ->whereIn('a.id_capacitacion', $capIds)
                ->whereRaw("YEAR($dateExpr) = ?", [$anio]) // <-- elimina esta línea si quieres TODAS las anualidades
                ->select(DB::raw('a.id_capacitacion as id'), DB::raw("DATE($dateExpr) as d"))
                ->get();
        } else {
            $rows = collect();
        }

        foreach ($rows as $r) {
            if (!$r->d) continue;
            $map[(int)$r->id][] = (string)$r->d;
        }
        foreach ($map as &$arr) {
            $arr = array_values(array_unique(array_filter($arr)));
            sort($arr);
        }
        return $map;
    };

    // Formateo: TODAS las fechas (sin límite)
    $fmtDates = function(array $dates) {
        if (empty($dates)) return '';
        $dates = array_map(fn($d)=>\Carbon\Carbon::parse($d)->format('d/m/Y'), $dates);
        return implode(', ', $dates);
    };

    // ========= 5) Estilos de columnas (incluye G) =========
    $rowLimit = 1000;
    try {
        $sheet->getColumnDimension('C')->setWidth(80);
        $sheet->getStyle('C7:C'.$rowLimit)->getAlignment()->setWrapText(true);
        $sheet->getStyle('C7:C'.$rowLimit)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);

        $sheet->getColumnDimension('E')->setWidth(70);
        $sheet->getStyle('E7:E'.$rowLimit)->getAlignment()->setWrapText(true);
        $sheet->getStyle('E7:E'.$rowLimit)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);

        $sheet->getColumnDimension('G')->setWidth(60);
        $sheet->getStyle('G7:G'.$rowLimit)->getAlignment()->setWrapText(true);
        $sheet->getStyle('G7:G'.$rowLimit)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
    } catch (\Throwable $e) {}

    // ========= 6) Relleno por riesgo (B=>riesgo, C=>puestos, E=>medidas, G=>fechas EPP/CAP) =========
    $colRiesgo  = 'B';
    $colPuestos = 'C';
    $row        = 7;
    $maxEmpty   = 30;
    $emptyStreak= 0;

    while ($row <= $rowLimit) {
        $name = (string)($sheet->getCell($colRiesgo.$row)->getValue());
        $nameTrim = trim($name);
        if ($nameTrim === '') {
            $emptyStreak++;
            if ($emptyStreak >= $maxEmpty) break;
            $row++;
            continue;
        }
        $emptyStreak = 0;

        $key = $norm($nameTrim);

        // SIEMPRE resolvemos el id del riesgo (para que haya medidas en E/G)
        $rid = $findClosestId($key);

        // Lista de puestos: si es un rótulo de QUÍMICOS, usamos $puestosQuimico; si no, los "Sí" del riesgo
        $esQuimicos = str_contains($key, 'QUIMIC'); // coincide con "EXPOSICION A QUIMICOS", "QUIMICOS", etc.
        if ($esQuimicos) {
            $puestosLista = $puestosQuimico;
        } elseif ($rid !== null && isset($puestosPorRiesgo[$rid])) {
            $puestosLista = array_values($puestosPorRiesgo[$rid]);
            sort($puestosLista, SORT_NATURAL|SORT_FLAG_CASE);
        } else {
            $puestosLista = [];
        }

        $sheet->setCellValue($colPuestos.$row, count($puestosLista) > 0
            ? (count($puestosLista).' puestos:'."\n".implode("\n", $puestosLista))
            : '');

        // --- Medidas en E/G (solo si tenemos rid resuelto) ---
        if ($rid !== null && isset($medidasPorRiesgo[$rid])) {
            $mm = $medidasPorRiesgo[$rid];
            $joinKeys = function($arr){ $keys = array_keys($arr); sort($keys, SORT_NATURAL|SORT_FLAG_CASE); return implode(', ', $keys); };

            $txtEpp   = $joinKeys($mm['epp']);
            $txtCap   = $joinKeys($mm['cap']);
            $txtSenal = $joinKeys($mm['senal']);
            $txtOtras = $joinKeys($mm['otras']);

            $sheet->setCellValue('E'.$row,     $txtEpp);
            $sheet->setCellValue('E'.($row+1), $txtCap);
            $sheet->setCellValue('E'.($row+2), $txtSenal);
            $sheet->setCellValue('E'.($row+3), $txtOtras);

            // Fechas EPP
            $eppIds = array_values(array_filter(array_map('intval', array_values($mm['epp_ids'])), fn($v)=>$v>0));
            $eppDatesMap = $fetchDatesMap($eppIds, $eppDateTables);
            $gLinesEpp = [];
            $eppNames = array_keys($mm['epp']); sort($eppNames, SORT_NATURAL|SORT_FLAG_CASE);
            foreach ($eppNames as $nm) {
                $id = $mm['epp_ids'][$nm] ?? 0;
                $dates = $id ? ($eppDatesMap[$id] ?? []) : [];
                if (!empty($dates)) $gLinesEpp[] = $nm.': '.$fmtDates($dates);
            }
            $sheet->setCellValue('G'.$row, implode("\n", $gLinesEpp));

            // Fechas CAP (asistencia_capacitacion.fecha_recibida VARCHAR)
            $capIds = array_values(array_filter(array_map('intval', array_values($mm['cap_ids'])), fn($v)=>$v>0));
            $capDatesMap = $fetchCapAttendanceDates($capIds);
            $gLinesCap = [];
            $capNames = array_keys($mm['cap']); sort($capNames, SORT_NATURAL|SORT_FLAG_CASE);
            foreach ($capNames as $nm) {
                $id = $mm['cap_ids'][$nm] ?? 0;
                $dates = $id ? ($capDatesMap[$id] ?? []) : [];
                if (!empty($dates)) $gLinesCap[] = $nm.': '.$fmtDates($dates);
            }
            $sheet->setCellValue('G'.($row+1), implode("\n", $gLinesCap));
        } else {
            // limpiar si no hay medidas
            $sheet->setCellValue('E'.$row,     '');
            $sheet->setCellValue('E'.($row+1), '');
            $sheet->setCellValue('E'.($row+2), '');
            $sheet->setCellValue('E'.($row+3), '');
            $sheet->setCellValue('G'.$row,     '');
            $sheet->setCellValue('G'.($row+1), '');
        }

        // Ajustes wrap/vertical por fila base
        foreach (['C','E','G'] as $col) {
            try {
                $sheet->getStyle($col.$row)->getAlignment()->setWrapText(true);
                $sheet->getStyle($col.$row)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
            } catch (\Throwable $e) {}
        }
        foreach (['E','G'] as $col) {
            try {
                $sheet->getStyle($col.($row+1))->getAlignment()->setWrapText(true);
                $sheet->getStyle($col.($row+1))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
            } catch (\Throwable $e) {}
        }

        $row++;
    }

    // ========= 7) Descargar =========
    $filename = 'plan_accion_control_riesgos_'.$anio.'_'.date('Ymd_His').'.xlsx';
    try {
        $tmp = storage_path('app/'.uniqid('plan_accion_', true).'.xlsx');
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($tmp);
        while (ob_get_level() > 0) { @ob_end_clean(); }
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
    } catch (\Throwable $e) {
        return back()->with('error', 'No se pudo generar el archivo de Excel: '.$e->getMessage());
    }

    return response()->download($tmp, $filename, [
        'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
        'Pragma'                    => 'no-cache',
        'Content-Transfer-Encoding' => 'binary',
    ])->deleteFileAfterSend(true);
}



    public function export(Request $request)
    {
        $mode   = $request->input('mode', 'verificacion');
        $by     = $request->input('by', 'puesto');
        $id     = (int) $request->input('id', 0);
        $soloSi = (bool) $request->boolean('solo_si', $mode==='buscar' && $by==='puesto');

        // Common lists
        $puestos = DB::table('puesto_trabajo_matriz as ptm')
            ->select('ptm.id_puesto_trabajo_matriz as id', 'ptm.puesto_trabajo_matriz as nombre')
            ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado', '!=', 0); })
            ->orderBy('ptm.puesto_trabajo_matriz')
            ->get();

        $ss = new Spreadsheet();
        $sh = $ss->getActiveSheet();
        $sh->setTitle('Resultados');

        $row = 1;
        $setHeader = function(array $cols) use ($sh, &$row) {
            $colIndex = 1;
            foreach ($cols as $h) {
                $colLetter = Coordinate::stringFromColumnIndex($colIndex++);
                $sh->setCellValue($colLetter.$row, $h);
            }
            $sh->getStyle("A{$row}:".Coordinate::stringFromColumnIndex((int)count($cols))."{$row}")
               ->getFont()->setBold(true);
            $sh->getStyle("A{$row}:".Coordinate::stringFromColumnIndex((int)count($cols))."{$row}")
               ->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $row++;
        };

        $esSiSql = "LOWER(TRIM(rv.valor)) IN ('s��','si','1')";

        if ($mode === 'verificacion' || ($mode==='buscar' && $by==='puesto')) {
            $puestoId = (int) $request->input('puesto', 0);
            if ($puestoId === 0 && $puestos->isNotEmpty()) { $puestoId = (int) $puestos->first()->id; }
            $puestoNombre = optional($puestos->firstWhere('id', $puestoId))->nombre;

            $sh->setCellValue("A{$row}", 'Verificación por Puesto');
            $row += 2;
            $sh->setCellValue("A".($row-1), 'Puesto:');
            $sh->setCellValue("B".($row-1), $puestoNombre);

            // Riesgos + valor
            $riesgos = DB::table('riesgo as r')
                ->join('tipo_riesgo as tr', 'tr.id_tipo_riesgo', '=', 'r.id_tipo_riesgo')
                ->leftJoin('riesgo_valor as rv', function ($join) use ($puestoId) {
                    $join->on('rv.id_riesgo', '=', 'r.id_riesgo')
                         ->where('rv.id_puesto_trabajo_matriz', '=', $puestoId);
                })
                ->select([
                    'tr.tipo_riesgo as tipo',
                    'r.nombre_riesgo as riesgo',
                    DB::raw("COALESCE(NULLIF(TRIM(rv.valor), ''), 'No') as valor"),
                ])
                ->orderBy('tr.tipo_riesgo')
                ->orderBy('r.nombre_riesgo')
                ->get();

            if ($soloSi) {
                $riesgos = $riesgos->filter(function ($r) {
                    $v = mb_strtolower(trim((string)$r->valor));
                    return in_array($v, ['s��','si','1'], true);
                })->values();
            }

            $setHeader(['Tipo', 'Riesgo', 'Valor']);
            foreach ($riesgos as $r) {
                $sh->setCellValue("A{$row}", $r->tipo);
                $sh->setCellValue("B{$row}", $r->riesgo);
                $sh->setCellValue("C{$row}", $r->valor);
                $row++;
            }

            // Anchos
            $sh->getColumnDimension('A')->setWidth(26);
            $sh->getColumnDimension('B')->setWidth(52);
            $sh->getColumnDimension('C')->setWidth(10);

        } else {
            // Buscar por Riesgo o Medidas: exportar tabla de resultados de puestos y riesgos relacionados
            $rows = collect();

            if ($by === 'riesgo' && $id > 0) {
                $rows = DB::table('puesto_trabajo_matriz as ptm')
                    ->join('riesgo_valor as rv', function ($join) use ($id) {
                        $join->on('rv.id_puesto_trabajo_matriz','=','ptm.id_puesto_trabajo_matriz')
                             ->where('rv.id_riesgo','=',$id);
                    })
                    ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado','!=',0); })
                    ->whereRaw($esSiSql)
                    ->leftJoin('riesgo as r','r.id_riesgo','=','rv.id_riesgo')
                    ->leftJoin('tipo_riesgo as tr','tr.id_tipo_riesgo','=','r.id_tipo_riesgo')
                    ->groupBy('ptm.id_puesto_trabajo_matriz','ptm.puesto_trabajo_matriz')
                    ->orderBy('ptm.puesto_trabajo_matriz')
                    ->select([
                        'ptm.puesto_trabajo_matriz as puesto',
                        DB::raw("GROUP_CONCAT(DISTINCT CONCAT(tr.tipo_riesgo,' | ', r.nombre_riesgo) ORDER BY tr.tipo_riesgo, r.nombre_riesgo SEPARATOR ', ') as riesgos"),
                    ])->get();
            } else {
                $mapMedida = [
                    'epp'          => ['col' => 'id_epp'],
                    'capacitacion' => ['col' => 'id_capacitacion'],
                    'senalizacion' => ['col' => 'id_senalizacion'],
                    'otras'        => ['col' => 'id_otras_medidas'],
                ];
                if (isset($mapMedida[$by]) && $id > 0) {
                    $col = $mapMedida[$by]['col'];
                    $rows = DB::table('puesto_trabajo_matriz as ptm')
                        ->join('riesgo_valor as rv', 'rv.id_puesto_trabajo_matriz','=','ptm.id_puesto_trabajo_matriz')
                        ->join('medidas_riesgo_puesto as mrp', 'mrp.id_riesgo','=','rv.id_riesgo')
                        ->join('riesgo as r','r.id_riesgo','=','rv.id_riesgo')
                        ->join('tipo_riesgo as tr','tr.id_tipo_riesgo','=','r.id_tipo_riesgo')
                        ->where("mrp.{$col}", $id)
                        ->where(function ($q) { $q->whereNull('ptm.estado')->orWhere('ptm.estado','!=',0); })
                        ->whereRaw($esSiSql)
                        ->groupBy('ptm.id_puesto_trabajo_matriz','ptm.puesto_trabajo_matriz')
                        ->orderBy('ptm.puesto_trabajo_matriz')
                        ->select([
                            'ptm.puesto_trabajo_matriz as puesto',
                            DB::raw("GROUP_CONCAT(DISTINCT CONCAT(tr.tipo_riesgo,' | ', r.nombre_riesgo) ORDER BY tr.tipo_riesgo, r.nombre_riesgo SEPARATOR ', ') as riesgos"),
                        ])->get();
                }
            }

            $setHeader(['Puesto', 'Riesgos relacionados']);
            foreach ($rows as $r) {
                $sh->setCellValue("A{$row}", $r->puesto);
                $sh->setCellValue("B{$row}", $r->riesgos);
                $row++;
            }
            $sh->getColumnDimension('A')->setWidth(48);
            $sh->getColumnDimension('B')->setWidth(80);
        }

        // Stream download
        $file = 'Verificacion_'.date('Ymd_His').'.xlsx';
        $writer = new Xlsx($ss);
        if (ob_get_length()) { ob_end_clean(); }
        return response()->streamDownload(function () use ($writer) {
            $writer->save('php://output');
        }, $file, [
            'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ]);
    }
}
