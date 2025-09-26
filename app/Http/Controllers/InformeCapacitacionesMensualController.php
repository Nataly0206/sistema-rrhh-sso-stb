<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;


class InformeCapacitacionesMensualController extends Controller
{
    public function index(Request $request)
    {
        // ====== Filtros ======
        $year  = (int)($request->input('anio') ?: date('Y'));
        $month = (int)($request->input('mes')  ?: date('n'));
        $q     = trim((string)$request->input('q', ''));

        // Primer y último día del mes (como DATE)
        $start = Carbon::createFromDate($year, $month, 1)->format('Y-m-d');
        $end   = Carbon::createFromDate($year, $month, 1)->endOfMonth()->format('Y-m-d');

        // Expresión SQL robusta para normalizar fecha_recibida (varchar/varios formatos/serial excel)
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

        // === Agregado de impartidas en el mes ===
        // - Participantes: COUNT(DISTINCT empleado)
        // - Fechas: todas las del mes (DISTINCT, ordenadas)
        // - Número: cuántos días distintos se impartió
        // - Departamentos: desde puesto_trabajo de los asistentes
        $impartidasAgg = DB::table('asistencia_capacitacion as ac')
            ->leftJoin('empleado as e', 'e.id_empleado', '=', 'ac.id_empleado')
            ->leftJoin('puesto_trabajo as pt', 'pt.id_puesto_trabajo', '=', 'e.id_puesto_trabajo')
            ->selectRaw('ac.id_capacitacion_instructor as ci_id')
            ->selectRaw('COUNT(DISTINCT ac.id_empleado)                          as participantes')
            ->selectRaw("GROUP_CONCAT(DISTINCT DATE_FORMAT($fechaExpr, '%Y-%m-%d') ORDER BY $fechaExpr SEPARATOR ', ') as fechas")
            ->selectRaw("COUNT(DISTINCT $fechaExpr)                                as numero")
            ->selectRaw("GROUP_CONCAT(DISTINCT TRIM(pt.departamento) ORDER BY pt.departamento SEPARATOR ', ') as departamentos")
            ->whereBetween(DB::raw($fechaExpr), [$start, $end])
            ->groupBy('ac.id_capacitacion_instructor');

        // === Consulta principal ===
        // Incluir SIEMPRE programadas = 'SI' + las que se impartieron en el mes aunque no sean programadas
        $programadaExpr = "UPPER(TRIM(COALESCE(ci.programada, '')))";

        $rows = DB::table('capacitacion_instructor as ci')
            ->join('capacitacion as c', 'c.id_capacitacion', '=', 'ci.id_capacitacion')
            ->leftJoinSub($impartidasAgg, 'd', function ($j) {
                $j->on('d.ci_id', '=', 'ci.id_capacitacion_instructor');
            })
            ->when($q !== '', function ($qq) use ($q) {
                $qq->where('c.capacitacion', 'like', "%{$q}%");
            })
            ->where(function ($w) use ($programadaExpr) {
                $w->whereRaw("$programadaExpr = 'SI'")->orWhereNotNull('d.ci_id');
            })
            ->selectRaw('ci.id_capacitacion_instructor as ci_id')
            ->selectRaw('c.capacitacion as tema')
            ->selectRaw('ci.duracion')
            ->selectRaw('ci.num_categoria as categoria')
            ->selectRaw('ci.programada')
            ->selectRaw('COALESCE(d.participantes, 0) as participantes')
            ->selectRaw("COALESCE(d.fechas, '') as fechas")
            ->selectRaw('COALESCE(d.numero, 0) as numero')
            ->selectRaw("COALESCE(d.departamentos, '') as departamentos")
            ->orderBy('tema')
            ->get()
            ->map(function ($r) {
                $r->horas_hombre = ((int)$r->participantes * (int)$r->duracion) / 60;
                return $r;
            });

        // Totales para el pie de tabla
        $totales = [
            'participantes' => $rows->sum('participantes'),
            'horas_hombre'  => $rows->sum('horas_hombre'),
            'numero'        => $rows->sum('numero'),
        ];

        // Datos para selects
        $months = [
            1=>'Enero',2=>'Febrero',3=>'Marzo',4=>'Abril',5=>'Mayo',6=>'Junio',
            7=>'Julio',8=>'Agosto',9=>'Septiembre',10=>'Octubre',11=>'Noviembre',12=>'Diciembre'
        ];
        $years = range((int)date('Y'), (int)date('Y') - 10);

        return view('capacitaciones.informe_mensual', [
            'rows'    => $rows,
            'totales' => $totales,
            'months'  => $months,
            'years'   => $years,
            'month'   => $month,
            'year'    => $year,
            'q'       => $q,
            'start'   => $start,
            'end'     => $end,
        ]);
    }

public function export(Request $request)
{
    // ====== Mismos filtros que la vista ======
    $year  = (int)($request->input('anio') ?: date('Y'));
    $month = (int)($request->input('mes')  ?: date('n'));
    $q     = trim((string)$request->input('q', ''));

    $start = \Carbon\Carbon::createFromDate($year, $month, 1)->format('Y-m-d');
    $end   = \Carbon\Carbon::createFromDate($year, $month, 1)->endOfMonth()->format('Y-m-d');

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

    $impartidasAgg = DB::table('asistencia_capacitacion as ac')
        ->leftJoin('empleado as e', 'e.id_empleado', '=', 'ac.id_empleado')
        ->leftJoin('puesto_trabajo as pt', 'pt.id_puesto_trabajo', '=', 'e.id_puesto_trabajo')
        ->selectRaw('ac.id_capacitacion_instructor as ci_id')
        ->selectRaw('COUNT(DISTINCT ac.id_empleado) as participantes')
        ->selectRaw("GROUP_CONCAT(DISTINCT DATE_FORMAT($fechaExpr, '%Y-%m-%d') ORDER BY $fechaExpr SEPARATOR ', ') as fechas")
        ->selectRaw("COUNT(DISTINCT $fechaExpr) as numero")
        ->selectRaw("GROUP_CONCAT(DISTINCT TRIM(pt.departamento) ORDER BY pt.departamento SEPARATOR ', ') as departamentos")
        ->whereBetween(DB::raw($fechaExpr), [$start, $end])
        ->groupBy('ac.id_capacitacion_instructor');

    $programadaExpr = "UPPER(TRIM(COALESCE(ci.programada, '')))";

    $rows = DB::table('capacitacion_instructor as ci')
        ->join('capacitacion as c', 'c.id_capacitacion', '=', 'ci.id_capacitacion')
        ->leftJoinSub($impartidasAgg, 'd', fn($j) => $j->on('d.ci_id', '=', 'ci.id_capacitacion_instructor'))
        ->when($q !== '', fn($qq) => $qq->where('c.capacitacion', 'like', "%{$q}%"))
        ->where(function ($w) use ($programadaExpr) {
            $w->whereRaw("$programadaExpr = 'SI'")->orWhereNotNull('d.ci_id');
        })
        ->selectRaw('ci.id_capacitacion_instructor as ci_id, c.capacitacion as tema, ci.duracion, ci.num_categoria as categoria, ci.programada')
        ->selectRaw('COALESCE(d.participantes, 0) as participantes')
        ->selectRaw("COALESCE(d.fechas, '') as fechas")
        ->selectRaw('COALESCE(d.numero, 0) as numero')
        ->selectRaw("COALESCE(d.departamentos, '') as departamentos")
        ->orderBy('tema')
        ->get()
        ->map(function ($r) {
            $r->horas_hombre = (int)$r->participantes * (int)$r->duracion; // igual que la vista
            $r->isProg = strtoupper(trim($r->programada ?? '')) === 'SI';
            return $r;
        });

    // ====== Armar Excel ======
    $xlsx = new Spreadsheet();
    $sheet = $xlsx->getActiveSheet();
    $sheet->setTitle('Mensual');

    // Título
    $sheet->setCellValue('A1', 'Informe mensual de Capacitaciones');
    $sheet->setCellValue('A2', "Periodo: $start a $end");
    $sheet->mergeCells('A1:J1');
    $sheet->mergeCells('A2:J2');
    $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);

    // Encabezados
    $header = ['N°','Tema','Participantes','Departamento(s)','Fecha(s) del mes','Duración','Categoría','No. Horas-Hombre','Número','Programada'];
    $sheet->fromArray($header, null, 'A4');
    $sheet->getStyle('A4:J4')->getFont()->setBold(true)->getColor()->setARGB('FFFFFFFF');
    $sheet->getStyle('A4:J4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF00B0F0');
    $sheet->getStyle('A4:J4')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_MEDIUM);

    // Datos
    $rowN = 5;
    $i = 1;
    foreach ($rows as $r) {
        $sheet->fromArray([
            $i++,
            $r->tema,
            (int)$r->participantes,
            $r->departamentos,
            $r->fechas,
            (int)$r->duracion.' h',
            $r->categoria ?? '—',
            (int)$r->horas_hombre,
            (int)$r->numero,
            $r->isProg ? 'SI' : 'NO',
        ], null, "A{$rowN}");

        // Pintar fila verde si es programada (igual que la vista)
        if ($r->isProg) {
            $sheet->getStyle("A{$rowN}:J{$rowN}")
                ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFECFDF5'); // verde suave
        }
        $rowN++;
    }

    // Totales (pie)
    $sheet->setCellValue("A{$rowN}", 'Totales:');
    $sheet->mergeCells("A{$rowN}:B{$rowN}");
    $sheet->getStyle("A{$rowN}")->getFont()->setBold(true);
    $sheet->setCellValue("C{$rowN}", array_sum($rows->pluck('participantes')->all()));
    $sheet->setCellValue("H{$rowN}", array_sum($rows->pluck('horas_hombre')->all()));
    $sheet->setCellValue("I{$rowN}", array_sum($rows->pluck('numero')->all()));

    // Ajustes
    foreach (range('A','J') as $col) { $sheet->getColumnDimension($col)->setAutoSize(true); }
    $sheet->freezePane('A5'); // congela encabezado

    // Descargar
    $file = "informe_mensual_capacitaciones_{$year}_{$month}.xlsx";
    $writer = new Xlsx($xlsx);
    return response()->streamDownload(function() use ($writer) {
        $writer->save('php://output');
    }, $file, [
        'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ]);
}

}
