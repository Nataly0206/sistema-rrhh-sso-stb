<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Arr;
use PhpOffice\PhpSpreadsheet\IOFactory;

class IdentificacionRiesgosExportController extends Controller
{
    public function export(Request $request)
    {
        // ===== 0) Parámetro =====
        $ptmId = (int) $request->input('ptm_id');
        if ($ptmId <= 0) {
            Log::warning('IdentificacionRiesgosExport: ptm_id requerido', ['payload' => $request->all()]);
            return response('ptm_id requerido', 422);
        }

        Log::info('IdentificacionRiesgosExport inicio', ['ptm_id' => $ptmId]);

        // ===== 1) Datos base =====
        $pt = DB::table('puesto_trabajo_matriz')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->first([
                'puesto_trabajo_matriz', 'id_departamento', 'num_empleados',
                'descripcion_general', 'actividades_diarias', 'id_area'
            ]);
        if (!$pt) return response('Puesto no encontrado', 404);

        $depto = $pt->id_departamento
            ? DB::table('departamento')->where('id_departamento', $pt->id_departamento)->value('departamento')
            : null;

        // Último registro de identificación
        $ir = DB::table('identificacion_riesgos')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->orderByDesc('id_identificacion_riesgos')
            ->first();

        // ===== 2) Agregados Químicos (preferir vista; fallback si no existe) =====
        $agg = [];
        try {
            $aggRow = DB::table('v_analisis_riesgos')
                ->where('puesto_id', $ptmId)
                ->selectRaw('
                    MAX(quimicos) as quimicos,
                    MAX(quimicos_detalle) as quimicos_detalle,
                    MAX(tipos_exposicion) as tipos_exposicion,
                    MAX(epp_quimicos) as epp_quimicos,
                    MAX(cap_quimicos) as cap_quimicos,
                    MAX(ilu_estandar_em) as ilu_std_em
                ')
                ->first();
            $agg = $aggRow ? (array) $aggRow : [];
        } catch (\Throwable $e) {
            // Fallback: construir químicos directo
            Log::notice('v_analisis_riesgos no disponible, usando fallback químicos: '.$e->getMessage());

            // quimicos_detalle: "Nombre [freq/dur]"
            $detalle = DB::table('quimico_puesto as qp')
                ->leftJoin('quimico as q', 'q.id_quimico', '=', 'qp.id_quimico')
                ->where('qp.id_puesto_trabajo_matriz', $ptmId)
                ->selectRaw("
                    GROUP_CONCAT(DISTINCT CONCAT(q.nombre_comercial, ' [', COALESCE(qp.frecuencia,''), '/', COALESCE(qp.duracion_exposicion,''), ']')
                                 ORDER BY q.nombre_comercial SEPARATOR ' | ') as quimicos_detalle,
                    GROUP_CONCAT(DISTINCT q.nombre_comercial ORDER BY q.nombre_comercial SEPARATOR ' | ') as quimicos,
                    GROUP_CONCAT(DISTINCT qp.epp ORDER BY qp.epp SEPARATOR ' | ') as epp_quimicos,
                    GROUP_CONCAT(DISTINCT qp.capacitacion ORDER BY qp.capacitacion SEPARATOR ' | ') as cap_quimicos
                ")
                ->first();

            $agg['quimicos_detalle'] = $detalle->quimicos_detalle ?? '';
            $agg['quimicos']         = $detalle->quimicos ?? '';
            $agg['epp_quimicos']     = $detalle->epp_quimicos ?? '';
            $agg['cap_quimicos']     = $detalle->cap_quimicos ?? '';

            // tipos_exposicion (opcional, si existe tabla)
            try {
                $te = DB::table('quimico_puesto as qp')
                    ->leftJoin('quimico_tipo_exposicion as qx', 'qx.id_quimico', '=', 'qp.id_quimico')
                    ->leftJoin('tipo_exposicion as te', 'te.id_tipo_exposicion', '=', 'qx.id_tipo_exposicion')
                    ->where('qp.id_puesto_trabajo_matriz', $ptmId)
                    ->selectRaw("GROUP_CONCAT(DISTINCT te.tipo_exposicion ORDER BY te.tipo_exposicion SEPARATOR ' | ') as tipos_exposicion")
                    ->first();
                $agg['tipos_exposicion'] = $te->tipos_exposicion ?? '';
            } catch (\Throwable $e2) {
                $agg['tipos_exposicion'] = '';
            }

            // estándar EM iluminación (si tuvieras estandar_iluminacion)
            try {
                // buscar última localización desde mediciones_iluminacion para estimar std
                $il = DB::table('mediciones_iluminacion')
                    ->where('id_puesto_trabajo_matriz', $ptmId)
                    ->orderByRaw('COALESCE(fecha_realizacion_inicio, fecha_realizacion_final) DESC')
                    ->orderByDesc('id')
                    ->first(['id_localizacion']);
                if ($il && $il->id_localizacion) {
                    $agg['ilu_std_em'] = DB::table('estandar_iluminacion')
                        ->where('id_localizacion', $il->id_localizacion)
                        ->value('em');
                }
            } catch (\Throwable $e3) {}
        }

        // ===== 3) Plantilla =====
        $candidates = [
            storage_path('app/public/formato_identificacion_riesgos.xlsx'),
            public_path('formato_identificacion_riesgos.xlsx'),
            resource_path('templates/formato_identificacion_riesgos.xlsx'),
        ];
        $tpl = null;
        foreach ($candidates as $c) if (is_file($c)) { $tpl = $c; break; }
        if (!$tpl) return response('No se encontró la plantilla de Excel', 500);

        try { $spreadsheet = IOFactory::load($tpl); }
        catch (\Throwable $e) {
            Log::error('Error cargando plantilla Excel', ['msg' => $e->getMessage()]);
            return response('No se pudo abrir la plantilla de Excel', 500);
        }
        $ws = $spreadsheet->getSheetByName('Hoja1') ?: $spreadsheet->getActiveSheet();

        // ===== 4) Helpers =====
        $sv = static fn($x) => $x === null ? '' : (string)$x;

        // Tri-estado A/NA/N/A → E/F/G con "X"
        $markTri = function (int $row, ?string $val) use ($ws): void {
            $ws->setCellValue("G{$row}", ''); $ws->setCellValue("H{$row}", ''); $ws->setCellValue("I{$row}", '');
            if ($val === 'A')      $ws->setCellValue("G{$row}", 'X');
            elseif ($val === 'NA') $ws->setCellValue("H{$row}", 'X');
            elseif ($val === 'N/A')$ws->setCellValue("I{$row}", 'X');
        };

        // Helper: 0/1 -> "No"/"Si" (vacío si viene null o cadena vacía)
        $yn = static function ($v) {
            if ($v === null || $v === '') return '';
            return ((int)$v === 1) ? 'Si' : 'No';
        };

        $markTri2 = function (int $row, ?string $val) use ($ws): void {
            $ws->setCellValue("E{$row}", ''); $ws->setCellValue("F{$row}", ''); $ws->setCellValue("G{$row}", '');
            if ($val === 'A')      $ws->setCellValue("E{$row}", 'X');
            elseif ($val === 'NA') $ws->setCellValue("F{$row}", 'X');
            elseif ($val === 'N/A')$ws->setCellValue("G{$row}", 'X');
        };

        // SI/NO/NA → E/G/H con "X"
        $markSINO = function (int $row, ?string $val) use ($ws): void {
            $ws->setCellValue("F{$row}", ''); $ws->setCellValue("G{$row}", ''); $ws->setCellValue("H{$row}", '');
            $v = strtoupper($sv = (string)$val);
            if ($v === 'SI') $ws->setCellValue("F{$row}", 'X');
            elseif ($v === 'NO') $ws->setCellValue("G{$row}", 'X');
            elseif ($v === 'NA' || $v === 'N/A') $ws->setCellValue("H{$row}", 'X');
        };

        // Marca "X" si bool/1
        $markX = function (string $addr, $flag) use ($ws): void {
            $ws->setCellValue($addr, ((int)($flag ?? 0) === 1) ? 'X' : '');
        };

        // ===== 5) Relleno =====

        // Datos Generales
        $ws->setCellValue('E8',  $sv($depto));
        $ws->setCellValue('E9',  $sv($pt->puesto_trabajo_matriz));
        $ws->setCellValue('E10', $sv($pt->num_empleados));
        $ws->setCellValue('E11', $sv($pt->descripcion_general));
        $ws->setCellValue('A13', $sv($pt->actividades_diarias));

        if ($ir) {
            // Esfuerzo físico (filas 16–19)
            // Cargar
            $ws->setCellValue('B16', $sv($ir->descripcion_carga_cargar));
            $ws->setCellValue('D16', $sv($ir->equipo_apoyo_cargar));
            $ws->setCellValue('F16', $sv($ir->duracion_carga_cargar));
            $ws->setCellValue('G16', $sv($ir->distancia_carga_cargar));
            $ws->setCellValue('H16', $sv($ir->frecuencia_carga_cargar));
            $ws->setCellValue('I16', $sv($ir->epp_cargar));
            $ws->setCellValue('K16', $sv($ir->peso_cargar));
            $ws->setCellValue('L16', $sv($ir->capacitacion_cargar));
            // Halar
            $ws->setCellValue('B17', $sv($ir->descripcion_carga_halar));
            $ws->setCellValue('D17', $sv($ir->equipo_apoyo_halar));
            $ws->setCellValue('F17', $sv($ir->duracion_carga_halar));
            $ws->setCellValue('G17', $sv($ir->distancia_carga_halar));
            $ws->setCellValue('H17', $sv($ir->frecuencia_carga_halar));
            $ws->setCellValue('I17', $sv($ir->epp_halar));
            $ws->setCellValue('K17', $sv($ir->peso_halar));
            $ws->setCellValue('L17', $sv($ir->capacitacion_halar));
            // Empujar
            $ws->setCellValue('B18', $sv($ir->descripcion_carga_empujar));
            $ws->setCellValue('D18', $sv($ir->equipo_apoyo_empujar));
            $ws->setCellValue('F18', $sv($ir->duracion_carga_empujar));
            $ws->setCellValue('G18', $sv($ir->distancia_carga_empujar));
            $ws->setCellValue('H18', $sv($ir->frecuencia_carga_empujar));
            $ws->setCellValue('I18', $sv($ir->epp_empujar));
            $ws->setCellValue('K18', $sv($ir->peso_empujar));
            $ws->setCellValue('L18', $sv($ir->capacitacion_empujar));
            // Sujetar
            $ws->setCellValue('B19', $sv($ir->descripcion_carga_sujetar));
            $ws->setCellValue('D19', $sv($ir->equipo_apoyo_sujetar));
            $ws->setCellValue('F19', $sv($ir->duracion_carga_sujetar));
            $ws->setCellValue('G19', $sv($ir->distancia_carga_sujetar));
            $ws->setCellValue('H19', $sv($ir->frecuencia_carga_sujetar));
            $ws->setCellValue('I19', $sv($ir->epp_sujetar));
            $ws->setCellValue('K19', $sv($ir->peso_sujetar));
            $ws->setCellValue('L19', $sv($ir->capacitacion_sujetar));

            $ws->setCellValue('D71', $yn($ir->sustancias_inflamables));
            $ws->setCellValue('D72', $yn($ir->senalización_de_riesgos));
            $ws->setCellValue('D73', $yn($ir->trasiego_liquidos));

            $ws->setCellValue('H71', $yn($ir->ventilacion_natural));
            $ws->setCellValue('H72', $yn($ir->fuentes_calor));
            $ws->setCellValue('H73', $yn($ir->cilindros_presion));

            $ws->setCellValue('L71', $yn($ir->limpiezas_regulares));
            $ws->setCellValue('L72', $yn($ir->maquinaria_friccion));
            $ws->setCellValue('L73', $yn($ir->derrames_sustancias));


            // Esfuerzo Visual (fila 22): D=Estándar (EM), G=Nivel medido, J=Tiempo

            // Ruido (fila 25: A=desc, D=nivel, F=tiempo)

            // Temperatura (fila 28: A=desc, D=grado/nivel, F=tiempo)
        }

        // Capacitaciones química

        // Condiciones de instalaciones (34–39)
        if ($ir) {
            $markTri2(34, $ir->paredes_muros_losas_trabes);  $ws->setCellValue('H34', $sv($ir->paredes_muros_losas_trabes_obs));
            $markTri2(35, $ir->pisos);                       $ws->setCellValue('H35', $sv($ir->pisos_obs));
            $markTri2(36, $ir->techos);                      $ws->setCellValue('H36', $sv($ir->techos_obs));
            $markTri2(37, $ir->puertas_ventanas);            $ws->setCellValue('H37', $sv($ir->puertas_ventanas_obs));
            $markTri2(38, $ir->escaleras_rampas);            $ws->setCellValue('H38', $sv($ir->escaleras_rampas_obs));
            $markTri2(39, $ir->anaqueles_estanterias);       $ws->setCellValue('H39', $sv($ir->anaqueles_estanterias_obs));
        }

        // Maquinaria, Equipo y Herramientas (46–57)
        if ($ir) {
            $markTri(46, $ir->maquinaria_equipos);          $ws->setCellValue('J46', $sv($ir->maquinaria_equipos_obs));
            $markTri(47, $ir->mantenimiento_preventivo);    $ws->setCellValue('J47', $sv($ir->mantenimiento_preventivo_obs));
            $markTri(48, $ir->mantenimiento_correctivo);    $ws->setCellValue('J48', $sv($ir->mantenimiento_correctivo_obs));
            $markTri(49, $ir->resguardos_guardas);          $ws->setCellValue('J49', $sv($ir->resguardos_guardas_obs));
            $markTri(50, $ir->conexiones_electricas);       $ws->setCellValue('J50', $sv($ir->conexiones_electricas_obs));
            $markTri(51, $ir->inspecciones_maquinaria);     $ws->setCellValue('J51', $sv($ir->inspecciones_maquinaria_obs));
            $markTri(52, $ir->paros_emergencia);            $ws->setCellValue('J52', $sv($ir->paros_emergencia_obs));
            $markTri(53, $ir->entrenamiento_maquinaria);    $ws->setCellValue('J53', $sv($ir->entrenamiento_maquinaria_obs));
            $markTri(54, $ir->epp_correspondiente);         $ws->setCellValue('J54', $sv($ir->epp_correspondiente_obs));
            $markTri(55, $ir->estado_herramientas);         $ws->setCellValue('J55', $sv($ir->estado_herramientas_obs));
            $markTri(56, $ir->inspecciones_herramientas);   $ws->setCellValue('J56', $sv($ir->inspecciones_herramientas_obs));
            $markTri(57, $ir->almacenamiento_herramientas); $ws->setCellValue('J57', $sv($ir->almacenamiento_herramientas_obs));
        }

        // Equipos y servicios de emergencia (60–69)
        if ($ir) {
            $markTri(60, $ir->rutas_evacuacion);            $ws->setCellValue('J60', $sv($ir->rutas_evacuacion_obs));
            $markTri(61, $ir->extintores_mangueras);        $ws->setCellValue('J61', $sv($ir->extintores_mangueras_obs));
            $markTri(62, $ir->camillas);                    $ws->setCellValue('J62', $sv($ir->camillas_obs));
            $markTri(63, $ir->botiquin);                    $ws->setCellValue('J63', $sv($ir->botiquin_obs));
            $markTri(64, $ir->simulacros);                  $ws->setCellValue('J64', $sv($ir->simulacros_obs));
            $markTri(65, $ir->plan_evacuacion);             $ws->setCellValue('J65', $sv($ir->plan_evacuacion_obs));
            $markTri(66, $ir->actuacion_emergencia);        $ws->setCellValue('J66', $sv($ir->actuacion_emergencia_obs));
            $markTri(67, $ir->alarmas_emergencia);          $ws->setCellValue('J67', $sv($ir->alarmas_emergencia_obs));
            $markTri(68, $ir->alarmas_humo);                $ws->setCellValue('J68', $sv($ir->alarmas_humo_obs));
            $markTri(69, $ir->lamparas_emergencia);         $ws->setCellValue('J69', $sv($ir->lamparas_emergencia_obs));
        }

        // Ergonómico (76–83) + Observaciones en I-col
        if ($ir) {
            $markSINO(76, $ir->movimientos_repetitivos);    $ws->setCellValue('I76', $sv($ir->movimientos_repetitivos_obs));
            $markSINO(77, $ir->posturas_forzadas);          $ws->setCellValue('I77', $sv($ir->posturas_forzadas_obs));
            $markSINO(78, $ir->suficiente_espacio);         $ws->setCellValue('I78', $sv($ir->suficiente_espacio_obs));
            $markSINO(79, $ir->elevacion_brazos);           $ws->setCellValue('I79', $sv($ir->elevacion_brazos_obs));
            $markSINO(80, $ir->giros_muneca);               $ws->setCellValue('I80', $sv($ir->giros_muneca_obs));
            $markSINO(81, $ir->inclinacion_espalda);        $ws->setCellValue('I81', $sv($ir->inclinacion_espalda_obs));
            $markSINO(82, $ir->herramienta_constante);      $ws->setCellValue('I82', $sv($ir->herramienta_constante_obs));
            $markSINO(83, $ir->herramienta_vibracion);      $ws->setCellValue('I83', $sv($ir->herramienta_vibracion_obs));
        }

        // Posturas (85–86)
        if ($ir) {
            $markX('B85', $ir->agachado);
            $markX('D85', $ir->rodillas);
            $markX('D86', $ir->balanceandose);
            $markX('F85', $ir->volteado);
            $markX('H85', $ir->parado);
            $markX('J85', $ir->sentado);
            $markX('B86', $ir->subiendo);
            $markX('F86', $ir->corriendo);
            $markX('H86', $ir->empujando);
            $markX('J86', $ir->halando);
            $markX('L86', $ir->arrastrandose);
            $markX('L86', $ir->girando);
        }

        // Trabajo en alturas (95–98) — usar el bloque derecho G:.. (campos de texto)
        if ($ir) {
            $ws->setCellValue('D95', $sv($ir->altura));                  // "Altura"
            $ws->setCellValue('D96', $sv($ir->altura_inspeccion_epp));   // "EPP utilizado"
            $ws->setCellValue('J97', $sv($ir->altura_capacitacion));     // "Capacitación recibida"
            $ws->setCellValue('J98', $sv($ir->hoja_trabajo));     
            $ws->setCellValue('J95', $sv($ir->medios_anclaje));     // "Capacitación recibida"
            $ws->setCellValue('J96', $sv($ir->altura_epp));   
            $ws->setCellValue('D97', $sv($ir->altura_senalizacion));     // "Capacitación recibida"
            $ws->setCellValue('D98', $sv($ir->aviso_altura));      
            // "Firma hoja de trabajo seguro" (si usas otro, cambia aquí)
            // Si quieres también anotar "Señalización" del lado izquierdo, podríamos usar otra celda libre o concatenar en G97.
        }

        // Riesgo eléctrico (100–102 + 105) — escribir "SI/NO/NA" en cada casilla
        if ($ir) {
            $ws->setCellValue('D100', strtoupper($sv($ir->senalizacion_delimitacion)));
            $ws->setCellValue('H100', strtoupper($sv($ir->capacitacion_certificacion)));

            $ws->setCellValue('H101', strtoupper($sv($ir->epp_utilizado_electri)));
            $ws->setCellValue('H102', strtoupper($sv($ir->aviso_trabajo_electrico)));

            $ws->setCellValue('L101', strtoupper($sv($ir->zonas_estatica)));
            $ws->setCellValue('L102', strtoupper($sv($ir->ausencia_tension)));
            $ws->setCellValue('L100', strtoupper($sv($ir->alta_tension)));

            $ws->setCellValue('D101', strtoupper($sv($ir->hoja_trabajo)));
            $ws->setCellValue('D101', strtoupper($sv($ir->epp_utilizado_electri)));
            $ws->setCellValue('D101', strtoupper($sv($ir->zonas_estatica)));

            $ws->setCellValue('D102', strtoupper($sv($ir->bloqueo_tarjetas)));
            $ws->setCellValue('D102', strtoupper($sv($ir->aviso_trabajo_electrico)));
            $ws->setCellValue('D102', strtoupper($sv($ir->ausencia_tension)));

            $ws->setCellValue('B104', strtoupper($sv($ir->cables_ordenados)));
            $ws->setCellValue('D104', strtoupper($sv($ir->tomacorrientes)));
            $ws->setCellValue('F104', strtoupper($sv($ir->cajas_interruptores)));
            $ws->setCellValue('H104', strtoupper($sv($ir->extensiones)));
            $ws->setCellValue('J104', strtoupper($sv($ir->cables_aislamiento)));
            $ws->setCellValue('L104', strtoupper($sv($ir->senalizacion_riesgo_electrico)));

            $ws->setCellValue('B105', $sv($ir->observaciones_electrico));

            $ws->setCellValue('B165', strtoupper($sv($ir->evaluacion_realizada)));
            $ws->setCellValue('B166', strtoupper($sv($ir->evaluacion_revisada)));
            $ws->setCellValue('K165', strtoupper($sv($ir->fecha_evaluacion_realizada)));
            $ws->setCellValue('K166', strtoupper($sv($ir->fecha_evaluacion_revisada)));
            $ws->setCellValue('C167', strtoupper($sv($ir->fecha_proxima_evaluaci)));
        }

        // Caída a mismo nivel (108–109)
        if ($ir) {
            $ws->setCellValue('B108', strtoupper($sv($ir->pisos_adecuado)));
            $ws->setCellValue('D108', strtoupper($sv($ir->vias_libres)));
            $ws->setCellValue('F108', strtoupper($sv($ir->rampas_identificados)));
            $ws->setCellValue('H108', strtoupper($sv($ir->gradas_barandas)));
            $ws->setCellValue('J108', strtoupper($sv($ir->sistemas_antideslizante)));
            $ws->setCellValue('L108', strtoupper($sv($ir->prevencion_piso_resbaloso)));
            $ws->setCellValue('B109', $sv($ir->observaciones_caida_nivel));

            $ws->setCellValue('C111', $sv($ir->otros_biologico));
            $ws->setCellValue('C112', $sv($ir->otros_psicosocial));
            $ws->setCellValue('C113', $sv($ir->otros_naturales));
        }

        // ===== 5.5) Marcar riesgos por nombre (columna B) con valor SI/NO en H/I y Observaciones en J =====
        try {
            // 1) Traer riesgo_valor unido a riesgo
            $riesgosRows = DB::table('riesgo_valor as rv')
                ->join('riesgo as r', 'r.id_riesgo', '=', 'rv.id_riesgo')
                ->where('rv.id_puesto_trabajo_matriz', $ptmId)
                ->select('r.nombre_riesgo', 'rv.valor', 'rv.observaciones')
                ->get();

            // 2) Normalizador (minúsculas + trim)
            $norm = static function ($s) {
                $s = (string) $s;
                $s = trim($s);
                if (function_exists('mb_strtolower')) {
                    $s = mb_strtolower($s, 'UTF-8');
                } else {
                    $s = strtolower($s);
                }
                return $s;
            };

            // 3) Mapa nombre_riesgo → {valor, obs}
            $riesgosMap = [];
            foreach ($riesgosRows as $row) {
                $key = $norm($row->nombre_riesgo);
                $riesgosMap[$key] = [
                    'valor' => strtoupper(trim((string) $row->valor)),      // "SI" / "NO" / etc.
                    'obs'   => (string) ($row->observaciones ?? ''),
                ];
            }

            // 4) Recorrer la hoja: si B{r} coincide con algún nombre_riesgo, marcar H/I y escribir J
            if (!empty($riesgosMap)) {
                $highestRow = $ws->getHighestDataRow(); // últimas filas con datos
                for ($r = 1; $r <= $highestRow; $r++) {
                    $label = $ws->getCell("B{$r}")->getValue();
                    if ($label === null || $label === '') {
                        continue;
                    }
                    $key = $norm($label);
                    if (!isset($riesgosMap[$key])) {
                        continue;
                    }

                    // Limpiar H / I y marcar según valor
                    $ws->setCellValue("H{$r}", '');
                    $ws->setCellValue("I{$r}", '');
                    if ($riesgosMap[$key]['valor'] === 'SI') {
                        $ws->setCellValue("H{$r}", 'X');  // SI → H
                    } else {
                        $ws->setCellValue("I{$r}", 'X');  // NO/otros → I
                    }

                    // Observaciones en J
                    $ws->setCellValue("J{$r}", $riesgosMap[$key]['obs']);
                }
            }
        } catch (\Throwable $e) {
            Log::error('Error mapeando riesgo_valor a Excel', ['ptm_id' => $ptmId, 'msg' => $e->getMessage()]);
            // No interrumpimos la descarga: seguimos sin marcar la sección si hay error.
        }

        // ===== VISUAL / RUIDO / ESTRÉS TÉRMICO (tablas dedicadas + mediciones + estándares) =====

        // --- VISUAL (A22, D22, G22, J22) ---
        $vis = DB::table('ident_esfuerzo_visual')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->first();

        $mi = DB::table('mediciones_iluminacion')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->orderByDesc('promedio') // ← punto MÁS ALTO
            ->first(['id_localizacion', 'promedio']);

        $emStd = null;
        if ($mi && $mi->id_localizacion) {
            $emStd = DB::table('estandar_iluminacion')
                ->where('id_localizacion', $mi->id_localizacion)
                ->value('em');
        }

        // Relleno Visual
        $ws->setCellValue('A22', $sv(optional($vis)->tipo_esfuerzo_visual));
        $ws->setCellValue('D22', $sv($emStd));
        $ws->setCellValue('G22', $mi ? (is_null($mi->promedio) ? '' : (string)$mi->promedio) : '');
        $ws->setCellValue('J22', $sv(optional($vis)->tiempo_exposicion));


        // --- RUIDO (A25, D25, E25, F25, H25) ---
        $irRuido = DB::table('ident_exposicion_ruido')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->first();

        // Elegir punto MÁS BAJO según (nivel_maximo + nivel_minimo)/2
        $mr = DB::table('mediciones_ruido')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->orderByRaw('(COALESCE(nivel_maximo,0)+COALESCE(nivel_minimo,0))/2 ASC')
            ->orderByDesc('id_mediciones_ruido') // tie-breaker estable
            ->first(['id_localizacion','nivel_maximo','nivel_minimo']);

        $nivelRuidoStd = null;
        if ($mr && $mr->id_localizacion) {
            $nivelRuidoStd = DB::table('estandar_ruido')
                ->where('id_localizacion', $mr->id_localizacion)
                ->value('nivel_ruido');
        }
        $ruidoMedio = null;
        if ($mr) {
            $mx = is_null($mr->nivel_maximo) ? 0 : (float)$mr->nivel_maximo;
            $mn = is_null($mr->nivel_minimo) ? 0 : (float)$mr->nivel_minimo;
            $ruidoMedio = ($mx + $mn) / 2.0;
        }

        // Relleno Ruido
        $ws->setCellValue('A25', $sv(optional($irRuido)->descripcion_ruido));
        $ws->setCellValue('D25', $nivelRuidoStd !== null ? (string)$nivelRuidoStd : '');
        $ws->setCellValue('E25', $ruidoMedio !== null ? (string)round($ruidoMedio, 2) : '');
        $ws->setCellValue('F25', $sv(optional($irRuido)->duracion_exposicion));
        $ws->setCellValue('H25', $sv(optional($irRuido)->epp));


        // --- ESTRÉS TÉRMICO (A28, D28, F28, H28) ---
        $iest = DB::table('ident_estres_termico')
            ->where('id_puesto_trabajo_matriz', $ptmId)
            ->first();

        // Prioridad de id_localizacion para estándar de temperatura:
        // 1) ident_estres_termico.id_localizacion (si existe)
        // 2) mediciones_ruido.id_localizacion
        // 3) mediciones_iluminacion.id_localizacion
        $tempLoc = null;
        if ($iest && isset($iest->id_localizacion) && $iest->id_localizacion) {
            $tempLoc = $iest->id_localizacion;
        } elseif ($mr && $mr->id_localizacion) {
            $tempLoc = $mr->id_localizacion;
        } elseif ($mi && $mi->id_localizacion) {
            $tempLoc = $mi->id_localizacion;
        }

        $tempStd = null;
        if ($tempLoc) {
            $tempStd = DB::table('estandar_temperatura')
                ->where('id_localizacion', $tempLoc)
                ->value('rango_temperatura');
        }

        // Relleno Estrés Térmico
        $ws->setCellValue('A28', $sv(optional($iest)->descripcion_stress_termico));
        $ws->setCellValue('D28', $tempStd !== null ? (string)$tempStd : '');
        $ws->setCellValue('F28', $sv(optional($iest)->duracion_exposicion));
        $ws->setCellValue('H28', $sv(optional($iest)->epp));

        // ===== 5.x) Químicos por puesto (A31, E31, G31, H31, I31, K31) =====
        // A31: nombres de químicos concatenados (sin repetir)
        // E31: tipos de exposición (DISTINCT, sin repetir)
        // G31: duraciones (de quimico_puesto, concatenadas)
        // H31: frecuencias (de quimico_puesto, concatenadas)
        // I31: EPP (de quimico_puesto, concatenado)
        // K31: Capacitaciones (de quimico_puesto, concatenadas)

        try {
            // 1) Nombres + campos de quimico_puesto
            $qpAgg = DB::table('quimico_puesto as qp')
                ->leftJoin('quimico as q', 'q.id_quimico', '=', 'qp.id_quimico')
                ->where('qp.id_puesto_trabajo_matriz', $ptmId)
                ->selectRaw("
                    GROUP_CONCAT(DISTINCT q.nombre_comercial ORDER BY q.nombre_comercial SEPARATOR ' | ')    AS nombres,
                    GROUP_CONCAT(DISTINCT qp.duracion_exposicion ORDER BY qp.duracion_exposicion SEPARATOR ' | ') AS duraciones,
                    GROUP_CONCAT(DISTINCT qp.frecuencia ORDER BY qp.frecuencia SEPARATOR ' | ')              AS frecuencias,
                    GROUP_CONCAT(DISTINCT qp.epp ORDER BY qp.epp SEPARATOR ' | ')                            AS epp_conc,
                    GROUP_CONCAT(DISTINCT qp.capacitacion ORDER BY qp.capacitacion SEPARATOR ' | ')          AS cap_conc
                ")
                ->first();

            // 2) Tipos de exposición (sin repetir)
            $tiposAgg = DB::table('quimico_puesto as qp')
                ->leftJoin('quimico_tipo_exposicion as qx', 'qx.id_quimico', '=', 'qp.id_quimico')
                ->leftJoin('tipo_exposicion as te', 'te.id_tipo_exposicion', '=', 'qx.id_tipo_exposicion')
                ->where('qp.id_puesto_trabajo_matriz', $ptmId)
                ->selectRaw("GROUP_CONCAT(DISTINCT te.tipo_exposicion ORDER BY te.tipo_exposicion SEPARATOR ' | ') AS tipos")
                ->first();

            // 3) Escribir en hoja
            $ws->setCellValue('A31', $sv(optional($qpAgg)->nombres));       // nombres químicos
            $ws->setCellValue('E31', $sv(optional($tiposAgg)->tipos));       // tipos de exposición DISTINCT
            $ws->setCellValue('G31', $sv(optional($qpAgg)->duraciones));     // duración(es)
            $ws->setCellValue('H31', $sv(optional($qpAgg)->frecuencias));    // frecuencia(s)
            $ws->setCellValue('I31', $sv(optional($qpAgg)->epp_conc));       // EPP
            $ws->setCellValue('K31', $sv(optional($qpAgg)->cap_conc));       // capacitación(es)
        } catch (\Throwable $e) {
            Log::error('Error agregando químicos a A31/E31/G31/H31/I31/K31', [
                'ptm_id' => $ptmId,
                'msg' => $e->getMessage()
            ]);
            // no interrumpimos la descarga
        }

        // ===== 5.y) QUÍMICO: marcar SI/NO por atributos (filas 131–134, H=SI, I=NO) =====
        try {
            $aggChem = DB::table('quimico_puesto as qp')
                ->join('quimico as q', 'q.id_quimico', '=', 'qp.id_quimico')
                ->where('qp.id_puesto_trabajo_matriz', $ptmId)
                ->selectRaw('
                    MAX(COALESCE(q.particulas_polvo,0))       AS particulas,
                    MAX(COALESCE(q.sustancias_corrosivas,0))  AS corrosivas,
                    MAX(COALESCE(q.sustancias_toxicas,0))     AS toxicas,
                    MAX(COALESCE(q.sustancias_irritantes,0))  AS irritantes
                ')
                ->first();

            // Helper: marca SI/NO en H/I de la fila dada
            $markYesNo = function (int $row, bool $yes) use ($ws): void {
                $ws->setCellValue("H{$row}", $yes ? 'X' : '');
                $ws->setCellValue("I{$row}", $yes ? '' : 'X');
            };

            // Si no hay registros, trata como todos NO
            $markYesNo(131, $aggChem && (int)$aggChem->particulas   > 0); // Partículas de polvo, humos, gases y vapores
            $markYesNo(132, $aggChem && (int)$aggChem->corrosivas   > 0); // Sustancias corrosivas
            $markYesNo(133, $aggChem && (int)$aggChem->toxicas      > 0); // Sustancias tóxicas
            $markYesNo(134, $aggChem && (int)$aggChem->irritantes   > 0); // Sustancias irritantes o alergizantes
        } catch (\Throwable $e) {
            \Log::error('Error marcando QUIMICO SI/NO', ['ptm_id' => $ptmId, 'msg' => $e->getMessage()]);
            // no interrumpimos la descarga
        }

        // ===== 6) Descargar =====
        $filename = 'Identificacion_Riesgos_'.$ptmId.'.xlsx';
        return response()->streamDownload(function () use ($spreadsheet) {
            try {
                while (ob_get_level() > 0) { @ob_end_clean(); }
                $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
                $writer->save('php://output');
            } finally {
                $spreadsheet->disconnectWorksheets();
            }
        }, $filename, [
            'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Transfer-Encoding' => 'binary',
            'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
            'Pragma'                    => 'no-cache',
            'X-Accel-Buffering'         => 'no',
        ]);
    }
}
