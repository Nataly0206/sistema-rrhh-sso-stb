<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

class InformeCapacitacionesAnualController extends Controller
{
    public function index(Request $request)
    {
        $year = (int)($request->input('anio') ?: date('Y'));
        $q    = trim((string)$request->input('q', ''));

        $start = Carbon::createFromDate($year, 1, 1)->format('Y-m-d');
        $end   = Carbon::createFromDate($year, 12, 31)->format('Y-m-d');

        // Normalizador de fecha_recibida (varchar) -> DATE
        $fechaExpr = "
        DATE(COALESCE(
            STR_TO_DATE(ac.fecha_recibida, '%Y-%m-%d'),
            STR_TO_DATE(ac.fecha_recibida, '%d/%m/%Y'),
            STR_TO_DATE(ac.fecha_recibida, '%d-%m-%Y'),
            STR_TO_DATE(ac.fecha_recibida, '%e/%m/%Y'),
            STR_TO_DATE(ac.fecha_recibida, '%e-%m-%Y'),
            STR_TO_DATE(ac.fecha_recibida, '%d %M %Y'),
            STR_TO_DATE(ac.fecha_recibida, '%e %M %Y'),
            STR_TO_DATE(ac.fecha_recibida, '%e de %M de %Y'),
            CASE
              WHEN ac.fecha_recibida REGEXP '^[0-9]{1,5}$'
              THEN DATE_ADD('1899-12-30', INTERVAL CAST(ac.fecha_recibida AS UNSIGNED) DAY)
            END
        ))";

        // Agregado anual desde asistencias
        $acAgg = DB::table('asistencia_capacitacion as ac')
            ->selectRaw('ac.id_capacitacion_instructor as ci_id')
            ->selectRaw("COUNT(DISTINCT $fechaExpr) as sesiones")          // # Capacitaciones (veces en el año)
            ->selectRaw('COUNT(DISTINCT ac.id_empleado) as personas')      // # personas (distintas en el año)
            ->selectRaw('COUNT(*) as asistencias')                         // filas de asistencia (para horas-hombre)
            ->whereBetween(DB::raw($fechaExpr), [$start, $end])
            ->groupBy('ac.id_capacitacion_instructor');

        $programadaExpr = "UPPER(TRIM(COALESCE(ci.programada, '')))";

        $rows = DB::table('capacitacion_instructor as ci')
            ->join('capacitacion as c', 'c.id_capacitacion', '=', 'ci.id_capacitacion')
            ->leftJoinSub($acAgg, 'a', function ($j) {
                $j->on('a.ci_id', '=', 'ci.id_capacitacion_instructor');
            })
            ->when($q !== '', fn($qq) => $qq->where('c.capacitacion', 'like', "%{$q}%"))
            ->where(function ($w) use ($programadaExpr) {
                // Incluir todas las programadas y las no programadas que se impartieron
                $w->whereRaw("$programadaExpr = 'SI'")->orWhereNotNull('a.ci_id');
            })
            ->selectRaw('ci.id_capacitacion_instructor as ci_id')
            ->selectRaw('c.capacitacion as tema')
            ->selectRaw('ci.num_categoria as categoria')
            ->selectRaw('ci.duracion as dur_min') // ¡minutos!
            ->selectRaw("$programadaExpr as programada_norm")
            ->selectRaw('COALESCE(a.sesiones, 0) as sesiones')
            ->selectRaw('COALESCE(a.personas, 0) as personas')
            ->selectRaw('COALESCE(a.asistencias, 0) as asistencias')
            ->orderBy('tema')
            ->get()
            ->map(function ($r) {
                $dur_h = max(0, (float)$r->dur_min) / 60.0;
                $r->dur_h = $dur_h;
                // Horas Capacitaciones = horas por sesión * # sesiones
                $r->horas_cap = round($dur_h * (int)$r->sesiones, 2);
                // Horas-hombre = (asistencias totales) * horas por sesión
                $r->horas_hombre = round($dur_h * (int)$r->asistencias, 2);
                $r->programada = (strtoupper(trim($r->programada_norm)) === 'SI') ? 'SI' : 'NO';
                return $r;
            });

        // Totales
        $tot = [
            'sesiones'      => (int) $rows->sum('sesiones'),
            'horas_cap'     => (float) $rows->sum('horas_cap'),
            'personas'      => (int) $rows->sum('personas'),
            'horas_hombre'  => (float) $rows->sum('horas_hombre'),
        ];

        // Porcentajes por fila
        $rows = $rows->map(function ($r) use ($tot) {
            $r->pct_sesiones = $tot['sesiones'] > 0 ? round(100 * $r->sesiones / $tot['sesiones'], 2) : 0.0;
            $r->pct_horas    = $tot['horas_cap'] > 0 ? round(100 * $r->horas_cap / $tot['horas_cap'], 2) : 0.0;
            return $r;
        });

        // KPIs: % ejecutadas / no ejecutadas SOLO sobre programadas
        $progTotal       = $rows->where('programada', 'SI')->count();
        $progEjecutadas  = $rows->where('programada', 'SI')->where('sesiones', '>', 0)->count();
        $progNoEjec      = $progTotal - $progEjecutadas;
        $kpi = [
            'total'       => $progTotal,
            'ejecutadas'  => $progEjecutadas,
            'no_ejec'     => $progNoEjec,
            'pct_ejec'    => $progTotal > 0 ? round(100 * $progEjecutadas / $progTotal, 1) : 0.0,
            'pct_no_ejec' => $progTotal > 0 ? round(100 * $progNoEjec / $progTotal, 1) : 0.0,
        ];

        $years = range((int)date('Y'), (int)date('Y') - 10);

        return view('capacitaciones.informe_anual', [
            'rows'  => $rows,
            'tot'   => $tot,
            'kpi'   => $kpi,
            'years' => $years,
            'year'  => $year,
            'q'     => $q,
        ]);
    }

public function export(Request $request)
{
    $year = (int)($request->input('anio') ?: date('Y'));
    $q    = trim((string)$request->input('q', ''));

    $start = \Carbon\Carbon::createFromDate($year, 1, 1)->format('Y-m-d');
    $end   = \Carbon\Carbon::createFromDate($year, 12, 31)->format('Y-m-d');

    $fechaExpr = "
    DATE(COALESCE(
        STR_TO_DATE(ac.fecha_recibida, '%Y-%m-%d'),
        STR_TO_DATE(ac.fecha_recibida, '%d/%m/%Y'),
        STR_TO_DATE(ac.fecha_recibida, '%d-%m-%Y'),
        STR_TO_DATE(ac.fecha_recibida, '%e/%m/%Y'),
        STR_TO_DATE(ac.fecha_recibida, '%e-%m-%Y'),
        STR_TO_DATE(ac.fecha_recibida, '%d %M %Y'),
        STR_TO_DATE(ac.fecha_recibida, '%e %M %Y'),
        STR_TO_DATE(ac.fecha_recibida, '%e de %M de %Y'),
        CASE WHEN ac.fecha_recibida REGEXP '^[0-9]{1,5}$'
        THEN DATE_ADD('1899-12-30', INTERVAL CAST(ac.fecha_recibida AS UNSIGNED) DAY) END
    ))";

    $acAgg = DB::table('asistencia_capacitacion as ac')
        ->selectRaw('ac.id_capacitacion_instructor as ci_id')
        ->selectRaw("COUNT(DISTINCT $fechaExpr) as sesiones")
        ->selectRaw('COUNT(DISTINCT ac.id_empleado) as personas')
        ->selectRaw('COUNT(*) as asistencias')
        ->whereBetween(DB::raw($fechaExpr), [$start, $end])
        ->groupBy('ac.id_capacitacion_instructor');

    $programadaExpr = "UPPER(TRIM(COALESCE(ci.programada, '')))";

    $rows = DB::table('capacitacion_instructor as ci')
        ->join('capacitacion as c', 'c.id_capacitacion', '=', 'ci.id_capacitacion')
        ->leftJoinSub($acAgg, 'a', fn($j) => $j->on('a.ci_id', '=', 'ci.id_capacitacion_instructor'))
        ->when($q !== '', fn($qq) => $qq->where('c.capacitacion', 'like', "%{$q}%"))
        ->where(function ($w) use ($programadaExpr) {
            $w->whereRaw("$programadaExpr = 'SI'")->orWhereNotNull('a.ci_id');
        })
        ->selectRaw('ci.id_capacitacion_instructor as ci_id, c.capacitacion as tema, ci.num_categoria as categoria, ci.duracion as dur_min')
        ->selectRaw("$programadaExpr as programada_norm")
        ->selectRaw('COALESCE(a.sesiones, 0) as sesiones, COALESCE(a.personas, 0) as personas, COALESCE(a.asistencias, 0) as asistencias')
        ->orderBy('tema')
        ->get()
        ->map(function ($r) {
            $dur_h = max(0, (float)$r->dur_min) / 60.0;
            $r->dur_h = $dur_h;
            $r->horas_cap = round($dur_h * (int)$r->sesiones, 2);
            $r->horas_hombre = round($dur_h * (int)$r->asistencias, 2);
            $r->programada = (strtoupper(trim($r->programada_norm)) === 'SI') ? 'SI' : 'NO';
            $r->done = ($r->programada === 'SI' && (int)$r->sesiones > 0);
            $r->pending = ($r->programada === 'SI' && (int)$r->sesiones === 0);
            return $r;
        });

    $tot = [
        'sesiones'     => (int) $rows->sum('sesiones'),
        'horas_cap'    => (float) $rows->sum('horas_cap'),
        'personas'     => (int) $rows->sum('personas'),
        'horas_hombre' => (float) $rows->sum('horas_hombre'),
    ];

    $rows = $rows->map(function ($r) use ($tot) {
        $r->pct_sesiones = $tot['sesiones'] > 0 ? round(100 * $r->sesiones / $tot['sesiones'], 2) : 0.0;
        $r->pct_horas    = $tot['horas_cap'] > 0 ? round(100 * $r->horas_cap / $tot['horas_cap'], 2) : 0.0;
        return $r;
    });

    $progTotal      = $rows->where('programada', 'SI')->count();
    $progEjecutadas = $rows->where('done', true)->count();
    $progNoEjec     = $rows->where('pending', true)->count();
    $pctEjec        = $progTotal > 0 ? round(100 * $progEjecutadas / $progTotal, 1) : 0.0;
    $pctNoEjec      = $progTotal > 0 ? round(100 * $progNoEjec / $progTotal, 1) : 0.0;

    // ====== Excel ======
    $xlsx = new Spreadsheet();
    $sheet = $xlsx->getActiveSheet();
    $sheet->setTitle('Anual');

    // Título / KPIs
    $sheet->setCellValue('A1', 'Informe anual de Capacitaciones');
    $sheet->setCellValue('A2', "Año: $year");
    $sheet->mergeCells('A1:H1');
    $sheet->mergeCells('A2:H2');
    $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);

    $sheet->setCellValue('F1', '% EJECUTADAS');
    $sheet->setCellValue('G1', $pctEjec.'%');
    $sheet->setCellValue('F2', '% NO EJECUTADAS');
    $sheet->setCellValue('G2', $pctNoEjec.'%');
    $sheet->getStyle('F1')->getFont()->setBold(true)->getColor()->setARGB('FF065F46'); // verde oscuro
    $sheet->getStyle('F2')->getFont()->setBold(true)->getColor()->setARGB('FF92400E'); // ámbar oscuro
    $sheet->getStyle('G1:G2')->getFont()->setBold(true);

    // Mapa de colores
    $sheet->setCellValue('A3', 'Mapa de colores: ');
    $sheet->setCellValue('B3', 'Programada impartida');
    $sheet->getStyle('B3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFECFDF5'); // verde
    $sheet->setCellValue('C3', 'Programada no impartida');
    $sheet->getStyle('C3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFBEB'); // amarillo

    // Encabezados
    $header = ['Tema','Categoría','# Capacitaciones','Horas Capacitaciones','% del total de sesiones','% del total de horas','# de personas','Horas-hombre'];
    $sheet->fromArray($header, null, 'A5');
    $sheet->getStyle('A5:H5')->getFont()->setBold(true)->getColor()->setARGB('FFFFFFFF');
    $sheet->getStyle('A5:H5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF00B0F0');
    $sheet->getStyle('A5:H5')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_MEDIUM);

    // Datos
    $rowN = 6;
    foreach ($rows as $r) {
        $sheet->fromArray([
            $r->tema,
            $r->categoria ?? '—',
            (int)$r->sesiones,
            number_format($r->horas_cap, 2, '.', ''),
            number_format($r->pct_sesiones, 2, '.', ''),
            number_format($r->pct_horas, 2, '.', ''),
            (int)$r->personas,
            number_format($r->horas_hombre, 2, '.', ''),
        ], null, "A{$rowN}");

        // Colorear fila según estado
        if ($r->done) { // programada impartida
            $sheet->getStyle("A{$rowN}:H{$rowN}")
                ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFECFDF5');
        } elseif ($r->pending) { // programada no impartida
            $sheet->getStyle("A{$rowN}:H{$rowN}")
                ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFBEB');
        }
        $rowN++;
    }

    // Totales
    $sheet->setCellValue("A{$rowN}", 'Totales:');
    $sheet->getStyle("A{$rowN}")->getFont()->setBold(true);
    $sheet->mergeCells("A{$rowN}:B{$rowN}");
    $sheet->setCellValue("C{$rowN}", $tot['sesiones']);
    $sheet->setCellValue("D{$rowN}", number_format($tot['horas_cap'], 2, '.', ''));
    $sheet->setCellValue("E{$rowN}", '100.00');
    $sheet->setCellValue("F{$rowN}", '100.00');
    $sheet->setCellValue("G{$rowN}", $tot['personas']);
    $sheet->setCellValue("H{$rowN}", number_format($tot['horas_hombre'], 2, '.', ''));

    // Formatos
    $sheet->getStyle("E6:E{$rowN}")->getNumberFormat()->setFormatCode('0.00"%"');
    $sheet->getStyle("F6:F{$rowN}")->getNumberFormat()->setFormatCode('0.00"%"');

    foreach (range('A','H') as $col) { $sheet->getColumnDimension($col)->setAutoSize(true); }
    $sheet->freezePane('A6');

    $file = "informe_anual_capacitaciones_{$year}.xlsx";
    $writer = new Xlsx($xlsx);
    return response()->streamDownload(function() use ($writer) {
        $writer->save('php://output');
    }, $file, [
        'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ]);
}

}
