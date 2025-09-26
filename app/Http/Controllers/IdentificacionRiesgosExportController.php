<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\IOFactory;

class IdentificacionRiesgosExportController extends Controller
{
    public function export(Request $request)
    {
        $request->validate([
            'ptm_id' => 'required|integer',
        ]);
        $ptmId = (int) $request->input('ptm_id');

        // =========== 0) PUESTO + último registro de identificación (opcional) ===========
        $irLast = DB::table('identificacion_riesgos')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->orderByDesc('id_identificacion_riesgos')
            ->limit(1);

        $pt = DB::table('puesto_trabajo_matriz as pt')
            ->leftJoin('departamento as d', 'd.id_departamento', '=', 'pt.id_departamento')
            ->leftJoin('localizacion as lo', 'lo.id_localizacion', '=', 'pt.id_localizacion')
            ->leftJoinSub($irLast, 'ir', function ($j) {
                $j->on('ir.id_puesto_trabajo_matriz', '=', 'pt.id_puesto_trabajo_matriz');
            })
            ->where('pt.id_puesto_trabajo_matriz', $ptmId)
            ->select(
                'pt.id_puesto_trabajo_matriz',
                'pt.id_localizacion',
                'pt.puesto_trabajo_matriz as puesto',
                'pt.num_empleados',
                'pt.descripcion_general',
                'pt.actividades_diarias',
                'd.departamento',
                'ir.*' // por si ya guardas algo aquí
            )
            ->first();

        if (!$pt) {
            return back()->with('error', 'No se encontró el puesto en "puesto_trabajo_matriz".');
        }

        $locId = $pt->id_localizacion;

        // =========== 1) ESTÁNDARES POR LOCALIZACIÓN ===========
        $stdLight = $locId ? DB::table('estandar_iluminacion')
            ->where('id_localizacion', $locId)
            ->orderByDesc('id_estandar_iluminacion')
            ->first(['em', 'ugr', 'ra', 'observaciones']) : null;

        $stdNoise = $locId ? DB::table('estandar_ruido')
            ->where('id_localizacion', $locId)
            ->orderByDesc('id_estandar_ruido')
            ->first(['nivel_ruido', 'tiempo_max_exposicion']) : null;

        $stdTemp = $locId ? DB::table('estandar_temperatura')
            ->where('id_localizacion', $locId)
            ->orderByDesc('id_estandar_temperatura')
            ->first(['rango_temperatura']) : null;

        // =========== 2) ÚLTIMAS MEDICIONES POR PUESTO/LOCALIZACIÓN ===========
        $lastLight = DB::table('mediciones_iluminacion')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->when($locId, fn($q) => $q->where('id_localizacion', $locId))
            ->orderByRaw('COALESCE(fecha_realizacion_inicio, fecha_realizacion_final) DESC')
            ->orderByDesc('id')
            ->first([
                'departamento', 'fecha_realizacion_inicio', 'fecha_realizacion_final',
                'nombre_observador', 'instrumento', 'serie', 'marca',
                'promedio', 'limites_aceptables', 'observaciones'
            ]);

        $lastNoise = DB::table('mediciones_ruido')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->when($locId, fn($q) => $q->where('id_localizacion', $locId))
            ->orderByRaw('COALESCE(fecha_realizacion_inicio, fecha_realizacion_final) DESC')
            ->orderByDesc('id_mediciones_ruido')
            ->first([
                'departamento', 'fecha_realizacion_inicio', 'fecha_realizacion_final',
                'nombre_observador', 'instrumento', 'serie', 'marca', 'nrr',
                'nivel_maximo', 'nivel_minimo', 'nivel_promedio',
                'limites_aceptables', 'observaciones'
            ]);

        // Observador/fecha para firmas: prioriza mediciones recientes si no hay en IR
        $obsLight = $lastLight->nombre_observador ?? null;
        $obsNoise = $lastNoise->nombre_observador ?? null;
        $obsMeas  = $obsLight ?: $obsNoise;

        $dateLight = $lastLight ? ($lastLight->fecha_realizacion_final ?: $lastLight->fecha_realizacion_inicio) : null;
        $dateNoise = $lastNoise ? ($lastNoise->fecha_realizacion_final ?: $lastNoise->fecha_realizacion_inicio) : null;
        $fechaMedicion = max((string)$dateLight, (string)$dateNoise); // YYYY-MM-DD compara bien como string

        // =========== 3) QUÍMICOS DEL PUESTO ===========
        $quimicos = DB::table('quimico_puesto as qp')
            ->join('quimico as q', 'q.id_quimico', '=', 'qp.id_quimico')
            ->where('qp.id_puesto_trabajo_matriz', $ptmId)
            ->where(function ($q) {
                // Solo activos si usas "estado" (1=activo) y evita "NINGUNO"
                $q->whereNull('q.estado')->orWhere('q.estado', 1);
            })
            ->select(
                'q.nombre_comercial',
                'q.uso',
                'qp.frecuencia',
                'qp.duracion_exposicion',
                'qp.capacitacion',
                'qp.epp'
            )
            ->get();

        // =========== 4) CABECERA PARA EXCEL ===========
        $header = [
            'departamento' => $pt->departamento ?? '',
            'puesto'       => $pt->puesto ?? '',
            'empleados'    => $pt->num_empleados ?? '',
            'descripcion'  => $pt->descripcion_general ?? '',
            'actividades'  => $pt->actividades_diarias ?? '',
        ];

        // =========== 5) ABRIR PLANTILLA ===========
        // =========== 5) ABRIR PLANTILLA ===========
        $tplPath = storage_path('app/public/formato_identificacion_riesgos.xlsx');
        if (!is_file($tplPath)) {
            return back()->with('error', 'No se encontró la plantilla formato_identificacion_riesgos.xlsx');
        }
        try {
            $spreadsheet = IOFactory::load($tplPath);
        } catch (\Throwable $e) {
            return back()->with('error', 'No se pudo abrir la plantilla de Excel: '.$e->getMessage());
        }

        // 👇👇 **FALTA EN TU CÓDIGO**: crea la hoja activa ANTES de usar $sheet
        $sheet = $spreadsheet->getSheetByName('Hoja1') ?? $spreadsheet->getActiveSheet();

        // =========== 6) ENCABEZADOS ===========
        $sheet->setCellValue('E8',  (string) $header['departamento']);
        $sheet->setCellValue('E9',  (string) $header['puesto']);
        $sheet->setCellValue('E10', (string) $header['empleados']);
        $sheet->setCellValue('E11', (string) $header['descripcion']);
        $sheet->setCellValue('A13', (string) $header['actividades']);

        // =========== 7) ESFUERZO VISUAL ===========
        // D22: estándar (EM); G22: lux medido (promedio); J22: periodo de medición
        $sheet->setCellValue('D22', (string)($stdLight->em ?? ''));
        $sheet->setCellValue('G22', (string)($lastLight->promedio ?? ''));

        if ($lastLight && ($lastLight->fecha_realizacion_inicio || $lastLight->fecha_realizacion_final)) {
            $fi = $lastLight->fecha_realizacion_inicio ?: null;
            $ff = $lastLight->fecha_realizacion_final   ?: null;
            $periodo = $fi && $ff ? ($fi.' a '.$ff) : ($fi ?: $ff);
            $sheet->setCellValue('J22', $periodo);
        }

        // =========== 8) EXPOSICIÓN A RUIDO ===========
        // D25: nivel de dB expuesto (promedio con rango si hay); F25: duración (estándar); H25: EPP (si lo manejas)
        if ($lastNoise) {
            $dbTxt = (string) ($lastNoise->nivel_promedio ?? '');
            if (!is_null($lastNoise->nivel_minimo) && !is_null($lastNoise->nivel_maximo)) {
                $dbTxt = trim($dbTxt.' (min '.$lastNoise->nivel_minimo.' – max '.$lastNoise->nivel_maximo.')');
            }
            $sheet->setCellValue('D25', $dbTxt);
            // Puedes, opcionalmente, anotar el NRR en observaciones si lo deseas en otra celda
        }
        $sheet->setCellValue('F25', (string) ($stdNoise->tiempo_max_exposicion ?? ''));
        // Si tienes un campo EPP para ruido:
        // $sheet->setCellValue('H25', (string) ($pt->epp_ruido ?? ''));

        // =========== 9) STRESS TÉRMICO ===========
        // A28: descripción; D28: grados (si tuvieras medición); F28: tiempo exposición (si aplica)
        $sheet->setCellValue('A28', $stdTemp ? ('Rango estándar: '.$stdTemp->rango_temperatura) : (string) ($pt->descripcion_temperatura ?? ''));
        // Mantengo D28/F28 si en el futuro agregas mediciones de temperatura
        $sheet->setCellValue('D28', (string) ($pt->nivel_mediciones_temperatura ?? ''));
        $sheet->setCellValue('F28', (string) ($pt->tiempo_exposicion_temperatura ?? ''));

        // =========== 10) QUÍMICOS (fila 31) ===========
        if ($quimicos->count() > 0) {
            $qDesc = $quimicos->map(function ($r) {
                return trim($r->nombre_comercial.($r->uso ? ' ('.$r->uso.')' : ''));
            })->filter()->unique()->values()->all();

            $dur   = $quimicos->pluck('duracion_exposicion')->filter()->unique()->values()->all();
            $freq  = $quimicos->pluck('frecuencia')->filter()->unique()->values()->all();
            $epp   = $quimicos->pluck('epp')->filter()->unique()->values()->all();
            $cap   = $quimicos->pluck('capacitacion')->filter()->unique()->values()->all();

            $sheet->setCellValue('A31', implode('; ', $qDesc));  // Descripción del químico
            // E31: "Tipo de exposición" -> si luego conectamos a quimico_tipo_exposicion, lo llenamos aquí
            $sheet->setCellValue('G31', implode('; ', $dur));    // Duración
            $sheet->setCellValue('H31', implode('; ', $freq));   // Frecuencia
            $sheet->setCellValue('I31', implode('; ', $epp));    // EPP utilizado
            $sheet->setCellValue('K31', implode('; ', $cap));    // Capacitación
        }

        // =========== 11) BLOQUES DE CONDICIONES / MAQUINARIA / EMERGENCIAS (igual que antes) ===========
        // CONDICIONES DE INSTALACIONES (34–39)
        $instRows = [
            34 => ['val' => $pt->paredes_muros_losas_trabes ?? null, 'obs' => $pt->paredes_muros_losas_trabes_obs ?? ''],
            35 => ['val' => $pt->pisos ?? null,                      'obs' => $pt->pisos_obs ?? ''],
            36 => ['val' => $pt->techos ?? null,                     'obs' => $pt->techos_obs ?? ''],
            37 => ['val' => $pt->puertas_ventanas ?? null,           'obs' => $pt->puertas_ventanas_obs ?? ''],
            38 => ['val' => $pt->escaleras_rampas ?? null,           'obs' => $pt->escaleras_rampas_obs ?? ''],
            39 => ['val' => $pt->anaqueles_estanterias ?? null,      'obs' => $pt->anaqueles_estanterias_obs ?? ''],
        ];
        $instColByEnum = ['A' => 'E', 'NA' => 'F', 'N/A' => 'G'];
        foreach ($instRows as $row => $it) {
            $enum = strtoupper((string) ($it['val'] ?? ''));
            $sheet->setCellValue("E{$row}", '');
            $sheet->setCellValue("F{$row}", '');
            $sheet->setCellValue("G{$row}", '');
            if (isset($instColByEnum[$enum])) {
                $sheet->setCellValue($instColByEnum[$enum].$row, 'X');
            }
            $sheet->setCellValue("H{$row}", (string) ($it['obs'] ?? ''));
        }

        // MAQUINARIA / EQUIPO / HERRAMIENTAS (46–57)
        $maqRows = [
            46 => ['val' => $pt->maquinaria_equipos ?? null,       'obs' => $pt->maquinaria_equipos_obs ?? ''],
            47 => ['val' => $pt->mantenimiento_preventivo ?? null,  'obs' => $pt->mantenimiento_preventivo_obs ?? ''],
            48 => ['val' => $pt->mantenimiento_correctivo ?? null,  'obs' => $pt->mantenimiento_correctivo_obs ?? ''],
            49 => ['val' => $pt->resguardos_guardas ?? null,        'obs' => $pt->resguardos_guardas_obs ?? ''],
            50 => ['val' => $pt->conexiones_electricas ?? null,     'obs' => $pt->conexiones_electricas_obs ?? ''],
            51 => ['val' => $pt->inspecciones_maquinaria ?? null,   'obs' => $pt->inspecciones_maquinaria_obs ?? ''],
            52 => ['val' => $pt->paros_emergencia ?? null,          'obs' => $pt->paros_emergencia_obs ?? ''],
            53 => ['val' => $pt->entrenamiento_maquinaria ?? null,  'obs' => $pt->entrenamiento_maquinaria_obs ?? ''],
            54 => ['val' => $pt->epp_correspondiente ?? null,       'obs' => $pt->epp_correspondiente_obs ?? ''],
            55 => ['val' => $pt->estado_herramientas ?? null,       'obs' => $pt->estado_herramientas_obs ?? ''],
            56 => ['val' => $pt->inspecciones_herramientas ?? null, 'obs' => $pt->inspecciones_herramientas_obs ?? ''],
            57 => ['val' => $pt->almacenamiento_herramientas ?? null,'obs' => $pt->almacenamiento_herramientas_obs ?? ''],
        ];
        $maqColByEnum = ['A' => 'G', 'NA' => 'H', 'N/A' => 'I'];
        foreach ($maqRows as $row => $it) {
            $enum = strtoupper((string) ($it['val'] ?? ''));
            $sheet->setCellValue("G{$row}", '');
            $sheet->setCellValue("H{$row}", '');
            $sheet->setCellValue("I{$row}", '');
            if (isset($maqColByEnum[$enum])) {
                $sheet->setCellValue($maqColByEnum[$enum].$row, 'X');
            }
            $sheet->setCellValue("J{$row}", (string) ($it['obs'] ?? ''));
        }

        // EMERGENCIA (60–69)
        $emerRows = [
            60 => ['val' => $pt->rutas_evacuacion ?? null,      'obs' => $pt->rutas_evacuacion_obs ?? ''],
            61 => ['val' => $pt->extintores_mangueras ?? null,  'obs' => $pt->extintores_mangueras_obs ?? ''],
            62 => ['val' => $pt->camillas ?? null,              'obs' => $pt->camillas_obs ?? ''],
            63 => ['val' => $pt->botiquin ?? null,              'obs' => $pt->botiquin_obs ?? ''],
            64 => ['val' => $pt->simulacros ?? null,            'obs' => $pt->simulacros_obs ?? ''],
            65 => ['val' => $pt->plan_evacuacion ?? null,       'obs' => $pt->plan_evacuacion_obs ?? ''],
            66 => ['val' => $pt->actuacion_emergencia ?? null,  'obs' => $pt->actuacion_emergencia_obs ?? ''],
            67 => ['val' => $pt->alarmas_emergencia ?? null,    'obs' => $pt->alarmas_emergencia_obs ?? ''],
            68 => ['val' => $pt->alarmas_humo ?? null,          'obs' => $pt->alarmas_humo_obs ?? ''],
            69 => ['val' => $pt->lamparas_emergencia ?? null,   'obs' => $pt->lamparas_emergencia_obs ?? ''],
        ];
        $emerColByEnum = ['A' => 'G', 'NA' => 'H', 'N/A' => 'I'];
        foreach ($emerRows as $row => $it) {
            $enum = strtoupper((string) ($it['val'] ?? ''));
            $sheet->setCellValue("G{$row}", '');
            $sheet->setCellValue("H{$row}", '');
            $sheet->setCellValue("I{$row}", '');
            if (isset($emerColByEnum[$enum])) {
                $sheet->setCellValue($emerColByEnum[$enum].$row, 'X');
            }
            $sheet->setCellValue("J{$row}", (string) ($it['obs'] ?? ''));
        }

        // ERGONÓMICO (76–83)
        $markErgo = function ($value, $row) use ($sheet) {
            $v = strtoupper((string) $value);
            $sheet->setCellValue("F{$row}", $v === 'SI' ? 'X' : '');
            $sheet->setCellValue("G{$row}", $v === 'NO' ? 'X' : '');
            $sheet->setCellValue("H{$row}", ($v === 'NA' || $v === 'N/A') ? 'X' : '');
        };
        $ergonomico = [
            76 => ['movimientos_repetitivos', 'movimientos_repetitivos_obs'],
            77 => ['posturas_forzadas',       'posturas_forzadas_obs'],
            78 => ['suficiente_espacio',      'suficiente_espacio_obs'],
            79 => ['elevacion_brazos',        'elevacion_brazos_obs'],
            80 => ['giros_muneca',            'giros_muneca_obs'],
            81 => ['inclinacion_espalda',     'inclinacion_espalda_obs'],
            82 => ['herramienta_constante',   'herramienta_constante_obs'],
            83 => ['herramienta_vibracion',   'herramienta_vibracion_obs'],
        ];
        foreach ($ergonomico as $row => [$campo, $campoObs]) {
            $markErgo($pt->{$campo} ?? null, $row);
            $sheet->setCellValue("I{$row}", (string) ($pt->{$campoObs} ?? ''));
        }

        // POSTURAS (85–86)
        $sheet->setCellValue('B85', !empty($pt->agachado)      ? 'X' : '');
        $sheet->setCellValue('D85', !empty($pt->rodillas)      ? 'X' : '');
        $sheet->setCellValue('F85', !empty($pt->volteado)      ? 'X' : '');
        $sheet->setCellValue('H85', !empty($pt->parado)        ? 'X' : '');
        $sheet->setCellValue('J85', !empty($pt->sentado)       ? 'X' : '');
        $sheet->setCellValue('L85', !empty($pt->arrastrandose) ? 'X' : '');
        $sheet->setCellValue('B86', !empty($pt->subiendo)      ? 'X' : '');
        $sheet->setCellValue('D86', !empty($pt->balanceandose) ? 'X' : '');
        $sheet->setCellValue('F86', !empty($pt->corriendo)     ? 'X' : '');
        $sheet->setCellValue('H86', !empty($pt->empujando)     ? 'X' : '');
        $sheet->setCellValue('J86', !empty($pt->halando)       ? 'X' : '');
        $sheet->setCellValue('L86', !empty($pt->girando)       ? 'X' : '');

        // ELECTRICIDAD (100–105)
        $sheet->setCellValue('D100', (string) ($pt->senalizacion_delimitacion ?? ''));
        $sheet->setCellValue('H100', (string) ($pt->capacitacion_certificacion ?? ''));
        $sheet->setCellValue('L100', (string) ($pt->alta_tension ?? ''));
        $sheet->setCellValue('D101', (string) ($pt->hoja_trabajo ?? ''));
        $sheet->setCellValue('H101', (string) ($pt->epp_correspondiente_obs ?? ''));
        $sheet->setCellValue('L101', (string) ($pt->zonas_estatica ?? ''));
        $sheet->setCellValue('D102', (string) ($pt->bloqueo_tarjetas ?? ''));
        $sheet->setCellValue('H102', (string) ($pt->aviso_trabajo_electrico ?? ''));
        $sheet->setCellValue('L102', (string) ($pt->ausencia_tension ?? ''));
        $sheet->setCellValue('B104', !empty($pt->cables_ordenados) ? 'X' : '');
        $sheet->setCellValue('D104', !empty($pt->tomacorrientes) ? 'X' : '');
        $sheet->setCellValue('F104', !empty($pt->cajas_interruptores) ? 'X' : '');
        $sheet->setCellValue('H104', !empty($pt->extensiones) ? 'X' : '');
        $sheet->setCellValue('J104', !empty($pt->cables_aislamiento) ? 'X' : '');
        $sheet->setCellValue('L104', !empty($pt->senalizacion_riesgo_electrico) ? 'X' : '');
        $sheet->setCellValue('B105', (string) ($pt->observaciones_electrico ?? ''));

        // CAÍDA MISMO NIVEL (108–109)
        $sheet->setCellValue('B108', !empty($pt->pisos_adecuado) ? 'X' : '');
        $sheet->setCellValue('D108', !empty($pt->vias_libres) ? 'X' : '');
        $sheet->setCellValue('F108', !empty($pt->rampas_identificados) ? 'X' : '');
        $sheet->setCellValue('H108', !empty($pt->gradas_barandas) ? 'X' : '');
        $sheet->setCellValue('J108', !empty($pt->sistemas_antideslizante) ? 'X' : '');
        $sheet->setCellValue('L108', !empty($pt->prevencion_piso_resbaloso) ? 'X' : '');
        $sheet->setCellValue('B109', (string) ($pt->observaciones_caida_nivel ?? ''));

        // =========== 12) TABLA DE IDENTIFICACIÓN DE RIESGO (118–165) ===========
        $riesgos = DB::table('riesgo_valor as rv')
            ->join('riesgo as r', 'r.id_riesgo', '=', 'rv.id_riesgo')
            ->where('rv.id_puesto_trabajo_matriz', $ptmId)
            ->select('r.nombre_riesgo as nombre', 'rv.valor', 'rv.observaciones')
            ->get();

        $norm = function (?string $s) {
            $s = (string) $s;
            $s = trim($s);
            if ($s === '') return '';
            $s = mb_strtoupper($s, 'UTF-8');
            $s = strtr($s, ['Á'=>'A','É'=>'E','Í'=>'I','Ó'=>'O','Ú'=>'U','Ü'=>'U','Ñ'=>'N']);
            $s = preg_replace('/[^A-Z0-9\/\s]/u', '', $s);
            $s = preg_replace('/\s+/', ' ', $s);
            return $s;
        };

        $agg = [];
        foreach ($riesgos as $r) {
            $k = $norm($r->nombre);
            if (!isset($agg[$k])) $agg[$k] = ['valor' => null, 'obs' => []];
            $v = mb_strtoupper(trim((string) $r->valor), 'UTF-8'); // SI/NO/N/A/NA
            if ($v === 'SI' || $agg[$k]['valor'] === null) {
                $agg[$k]['valor'] = 'SI';
            } elseif ($v === 'NO' && $agg[$k]['valor'] !== 'SI') {
                $agg[$k]['valor'] = 'NO';
            } elseif (($v === 'N/A' || $v === 'NA') && $agg[$k]['valor'] === null) {
                $agg[$k]['valor'] = 'N/A';
            }
            if (!empty($r->observaciones)) $agg[$k]['obs'][] = (string) $r->observaciones;
        }

        $find = function (string $label) use ($norm, $agg) {
            $L = $norm($label);
            if ($L === '') return null;
            if (isset($agg[$L])) return $agg[$L];
            foreach ($agg as $k => $dat) {
                if (str_contains($k, $L) || str_contains($L, $k)) return $dat;
            }
            return null;
        };

        for ($row = 118; $row <= 165; $row++) {
            $peligro = (string) $sheet->getCell("B{$row}")->getValue(); // Columna PELIGRO
            if (trim($peligro) === '') continue;
            $m = $find($peligro);
            if (!$m) continue;
            $val = $m['valor'] ?? '';
            $sheet->setCellValue("H{$row}", $val === 'SI' ? 'X' : '');
            $sheet->setCellValue("I{$row}", $val === 'NO' ? 'X' : '');
            $sheet->setCellValue("J{$row}", !empty($m['obs']) ? implode('; ', array_unique($m['obs'])) : '');
        }

        // =========== 13) FIRMAS Y FECHAS (166–168) ===========
        $realizadaPor   = $pt->evaluacion_realizada_por ?? $pt->realizada_por ?? $pt->observador ?? $obsMeas ?? '';
        $fechaRealizada = $pt->fecha_realizada ?? $pt->fecha_evaluacion ?? $pt->fecha ?? $fechaMedicion ?? '';
        $revisadaPor    = $pt->evaluacion_revisada_por ?? $pt->revisada_por ?? '';
        $fechaRevisada  = $pt->fecha_revisada ?? '';
        $fechaProxima   = $pt->fecha_proxima_evaluacion ?? $pt->fecha_proxima ?? '';

        $sheet->setCellValue('B166', (string) $realizadaPor);
        $sheet->setCellValue('K166', (string) $fechaRealizada);
        $sheet->setCellValue('B167', (string) $revisadaPor);
        $sheet->setCellValue('K167', (string) $fechaRevisada);
        $sheet->setCellValue('C168', (string) $fechaProxima);

        // =========== 14) DESCARGAR ===========
        $filename = 'identificacion_riesgos_'.$ptmId.'_'.date('Ymd_His').'.xlsx';
        try {
            $tmp = storage_path('app/'.uniqid('identificacion_', true).'.xlsx');
            $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            $writer->save($tmp);
            if (!is_file($tmp) || filesize($tmp) < 500) { $writer->save($tmp); }
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
}
