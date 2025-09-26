<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use App\Models\AsistenciaCapacitacion;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

class Capacitaciones extends Controller
{
    public function instructor(Request $request)
    {
        $instructores = DB::table('instructor as i')
            ->select('*')
            ->when($request->search, function ($query, $search) {
                return $query->where('i.instructor', 'like', "%{$search}%");
            })
            ->paginate(10)
            ->appends(['search' => $request->search]);
        return view('capacitaciones.instructor', compact('instructores'));
    }

    public function storeinstructor(Request $request)
    {
        DB::table('instructor')->insert([
            'instructor' => $request->input('instructor'),
        ]);
        return redirect()->back()->with('success', 'instructor registrado correctamente.');
    }

    public function updateinstructor(Request $request, $id)
    {
        DB::table('instructor')
            ->where('id_instructor', $id)
            ->update([
                'instructor' => $request->input('instructor'),
            ]);
        return redirect()->back()->with('success', 'instructor actualizado correctamente');
    }

    public function destroyinstructor($id)
    {
        DB::table('instructor')->where('id_instructor', $id)->delete();
        return redirect()->back()->with('success', 'instructor eliminado correctamente');
    }

    public function capacitacion(Request $request)
    {
        $capacitaciones = DB::table('capacitacion as c')
            ->select('*')
            ->when($request->search, function ($query, $search) {
                return $query->where('c.capacitacion', 'like', "%{$search}%");
            })
            ->paginate(10)
            ->appends(['search' => $request->search]);
        return view('capacitaciones.capacitacion', compact('capacitaciones'));
    }

    public function storecapacitacion(Request $request)
    {
        DB::table('capacitacion')->insert([
            'capacitacion' => $request->input('capacitacion'),
        ]);
        return redirect()->back()->with('success', 'capacitacion registrado correctamente.');
    }

    public function updatecapacitacion(Request $request, $id)
    {
        DB::table('capacitacion')
            ->where('id_capacitacion', $id)
            ->update([
                'capacitacion' => $request->input('capacitacion'),
            ]);
        return redirect()->back()->with('success', 'capacitacion actualizado correctamente');
    }

    public function destroycapacitacion($id)
    {
        DB::table('capacitacion')->where('id_capacitacion', $id)->delete();
        return redirect()->back()->with('success', 'capacitacion eliminado correctamente');
    }

    public function Capinstructor(Request $request)
    {
        $capinstructor = DB::table('capacitacion_instructor as ci')
            ->join('instructor as i', 'ci.id_instructor', '=', 'i.id_instructor')
            ->join('capacitacion as c', 'ci.id_capacitacion', '=', 'c.id_capacitacion')
            ->select('ci.*', 'i.*', 'c.*')
            ->when($request->search, function ($query, $search) {
                return $query->where('i.instructor', 'like', "%{$search}%")
                             ->orWhere('c.capacitacion', 'like', "%{$search}%");
            })
            ->orderBy('c.capacitacion', 'asc')
            ->paginate(10)
            ->appends(['search' => $request->search]);
        $capacitaciones = DB::select('CALL sp_obtener_capacitacion()');
        $instructores = DB::select('CALL sp_obtener_instructor()');
        return view('capacitaciones.detallescapacitacion', compact('capinstructor', 'capacitaciones', 'instructores'));
    }

    public function storecapinstructor(Request $request)
    {
        DB::table('capacitacion_instructor')->insert([
            'id_capacitacion' => $request->input('id_capacitacion'),
            'id_instructor' => $request->input('id_instructor'),
            'duracion' => $request->input('duracion'),
            'num_categoria' => $request->input('num_categoria'),
            'programada' => $request->has('programada') ? 'SI' : 'NO',
        ]);
        return redirect()->back()->with('success', 'Capacitacion con su instructor agregado correctamente');
    }

    public function updatecapinstructor(Request $request, $id)
    {
        DB::table('capacitacion_instructor')
            ->where('id_capacitacion_instructor', $id)
            ->update([
                'id_capacitacion' => $request->input('id_capacitacion'),
                'id_instructor' => $request->input('id_instructor'),
                'duracion' => $request->input('duracion'),
                'num_categoria' => $request->input('num_categoria'),
                'programada' => $request->has('programada') ? 'SI' : 'NO',
            ]);
        return redirect()->back()->with('success', 'capacitacion actualizado correctamente');
    }

        public function destroycapinstructor($id)
    {
        DB::table('capacitacion_instructor')->where('id_capacitacion_instructor', $id)->delete();
        return redirect()->back()->with('success', 'capacitacion_instructor eliminado correctamente');
    }

    public function Asistencia(Request $request)
    {
        $capinstructor = DB::select('CALL sp_obtener_capacitacion_instructor()');
        $empleados = DB::select('CALL sp_obtener_empleados()');
        $asistencia = DB::table('asistencia_capacitacion AS ac')
            ->join('empleado AS e', 'ac.id_empleado', '=', 'e.id_empleado')
            ->join('puesto_trabajo AS pt', 'e.id_puesto_trabajo', '=', 'pt.id_puesto_trabajo')
            ->leftJoin('puesto_trabajo_matriz as ptm', 'e.id_puesto_trabajo_matriz', '=', 'ptm.id_puesto_trabajo_matriz')
            ->join('capacitacion_instructor AS ci', 'ac.id_capacitacion_instructor', '=', 'ci.id_capacitacion_instructor')
            ->join('capacitacion AS c', 'ci.id_capacitacion', '=', 'c.id_capacitacion')
            ->join('instructor AS i', 'ci.id_instructor', '=', 'i.id_instructor')
            ->select(
                'ac.id_asistencia_capacitacion',
                'e.id_empleado',
                'e.nombre_completo',
                'e.codigo_empleado',
                DB::raw('COALESCE(ptm.puesto_trabajo_matriz, pt.puesto_trabajo) AS puesto_trabajo'),
                'pt.departamento',
                'c.capacitacion',
                DB::raw('COALESCE(ac.instructor_temporal, i.instructor) AS instructor'),
                'ac.fecha_recibida'
            )
            ->when($request->search, function ($query, $search) {
                $term = "%{$search}%";
                return $query->where(function ($inner) use ($term) {
                    $inner->where('e.nombre_completo', 'LIKE', $term)
                        ->orWhere('e.codigo_empleado', 'LIKE', $term)
                        ->orWhere('pt.puesto_trabajo', 'LIKE', $term)
                        ->orWhere('ptm.puesto_trabajo_matriz', 'LIKE', $term)
                        ->orWhere('pt.departamento', 'LIKE', $term)
                        ->orWhere('c.capacitacion', 'LIKE', $term)
                        ->orWhere('i.instructor', 'LIKE', $term)
                        ->orWhere('ac.instructor_temporal', 'LIKE', $term);
                });
            })
            ->orderBy('ac.id_asistencia_capacitacion', 'desc')
            ->paginate(10)
            ->appends(['search' => $request->search]);
        return view('capacitaciones.asistencia', compact('capinstructor', 'empleados', 'asistencia'));
    }

    public function store(Request $request)
    {
        $request->validate([
            'empleados' => 'required|array',
            'capacitaciones' => 'required|array',
            'fecha' => 'required|string|max:100',
            'instructor' => 'nullable|string|max:100',
        ]);

        $empleados = (array) $request->empleados;
        $capacitaciones = (array) $request->capacitaciones;
        $fecha = $request->fecha;

        $empleadoMap = DB::table('empleado')
            ->whereIn('id_empleado', $empleados)
            ->pluck('nombre_completo', 'id_empleado');

        $capMap = DB::table('capacitacion_instructor as ci')
            ->join('capacitacion as c', 'ci.id_capacitacion', '=', 'c.id_capacitacion')
            ->whereIn('ci.id_capacitacion_instructor', $capacitaciones)
            ->pluck('c.capacitacion', 'ci.id_capacitacion_instructor');

        $duplicados = [];
        $creados = 0;

        foreach ($empleados as $idEmpleado) {
            foreach ($capacitaciones as $idCapacitacionInstructor) {
                $yaExiste = AsistenciaCapacitacion::query()
                    ->where('id_empleado', $idEmpleado)
                    ->where('id_capacitacion_instructor', $idCapacitacionInstructor)
                    ->where('fecha_recibida', $fecha)
                    ->exists();

                if ($yaExiste) {
                    $duplicados[] = (
                        ($empleadoMap[$idEmpleado] ?? 'Empleado '.$idEmpleado)
                        .' - '.($capMap[$idCapacitacionInstructor] ?? 'Cap '.$idCapacitacionInstructor)
                        .' - '.$fecha
                    );
                    continue;
                }

                AsistenciaCapacitacion::create([
                    'id_empleado' => $idEmpleado,
                    'id_capacitacion_instructor' => $idCapacitacionInstructor,
                    'instructor_temporal' => $request->instructor ?: null,
                    'fecha_recibida' => $fecha,
                ]);
                $creados++;
            }
        }

        $mensajeOk = $creados > 0
            ? "Asistencias registradas correctamente ($creados nuevas)."
            : 'No se registraron nuevas asistencias.';

        $redirect = redirect()->back()->with('success', $mensajeOk);
        if (!empty($duplicados)) {
            $redirect->with('warning', 'Se omitieron registros duplicados: '.implode('; ', $duplicados));
        }
        return $redirect;
    }

    public function destroyasistencia($id)
    {
        DB::table('asistencia_capacitacion')->where('id_asistencia_capacitacion', $id)->delete();
        return redirect()->back()->with('success', 'asistencia a capacitacion eliminado correctamente');
    }

    public function consulta(Request $request)
    {
        $nombre = $request->input('nombre');
        $puesto = $request->input('puesto');
        $fecha = $request->input('fecha');
        $capacitacion = $request->input('capacitacion');

        $query = DB::table('asistencia_capacitacion as asica')
            ->join('empleado as emp', 'asica.id_empleado', '=', 'emp.id_empleado')
            ->join('puesto_trabajo as pt', 'emp.id_puesto_trabajo', '=', 'pt.id_puesto_trabajo')
            ->leftJoin('puesto_trabajo_matriz as ptm', 'emp.id_puesto_trabajo_matriz', '=', 'ptm.id_puesto_trabajo_matriz')
            ->join('capacitacion_instructor as cap', 'asica.id_capacitacion_instructor', '=', 'cap.id_capacitacion_instructor')
            ->join('capacitacion as capa', 'cap.id_capacitacion', '=', 'capa.id_capacitacion')
            ->join('instructor as i', 'cap.id_instructor', '=', 'i.id_instructor')
            ->select(
                'emp.nombre_completo',
                DB::raw('COALESCE(ptm.puesto_trabajo_matriz, pt.puesto_trabajo) AS puesto_trabajo'),
                'pt.departamento',
                'capa.capacitacion',
                'i.instructor',
                'asica.fecha_recibida'
            );

        if ($nombre) { $query->where('emp.nombre_completo', $nombre); }
        if ($puesto) {
            $query->where(function ($q) use ($puesto) {
                $q->where('ptm.puesto_trabajo_matriz', $puesto)
                    ->orWhere('pt.puesto_trabajo', $puesto);
            });
        }
        if ($fecha) { $query->where('asica.fecha_recibida', $fecha); }
        if ($capacitacion) { $query->where('capa.capacitacion', $capacitacion); }

    $consulta = $query->paginate(10)->appends($request->except('page'));

        $empleadosConCap = DB::table('asistencia_capacitacion as asica')
            ->join('empleado as emp', 'asica.id_empleado', '=', 'emp.id_empleado')
            ->select('emp.id_empleado','emp.nombre_completo')
            ->distinct()
            ->get();
        $puesto = DB::table('puesto_trabajo')
            ->where('estado', 1)
            ->pluck('puesto_trabajo');
        $puestosMatriz = DB::table('puesto_trabajo_matriz')
            ->where('estado', 1)
            ->pluck('puesto_trabajo_matriz');
        $puesto = $puesto->merge($puestosMatriz)->unique()->sort()->values();
        $fecha = DB::table('asistencia_capacitacion')->distinct()->pluck('fecha_recibida');
        $capacitacion = DB::table('capacitacion')->pluck('capacitacion');

        return view('capacitaciones.consultas', compact('consulta', 'empleadosConCap', 'puesto', 'fecha', 'capacitacion'));
    }

    public function imprimir(Request $request)
{
    $nombre       = $request->input('nombre');
    $puesto       = $request->input('puesto');
    $fecha        = $request->input('fecha');
    $capacitacion = $request->input('capacitacion');

    $query = DB::table('asistencia_capacitacion as asica')
        ->join('empleado as emp', 'asica.id_empleado', '=', 'emp.id_empleado')
        ->join('puesto_trabajo as pt', 'emp.id_puesto_trabajo', '=', 'pt.id_puesto_trabajo')
        ->leftJoin('puesto_trabajo_matriz as ptm', 'emp.id_puesto_trabajo_matriz', '=', 'ptm.id_puesto_trabajo_matriz')
        ->join('capacitacion_instructor as cap', 'asica.id_capacitacion_instructor', '=', 'cap.id_capacitacion_instructor')
        ->join('capacitacion as capa', 'cap.id_capacitacion', '=', 'capa.id_capacitacion')
        ->join('instructor as i', 'cap.id_instructor', '=', 'i.id_instructor')
        ->select(
            'emp.nombre_completo',
            DB::raw('COALESCE(ptm.puesto_trabajo_matriz, pt.puesto_trabajo) AS puesto_trabajo'),
            'pt.departamento',
            'capa.capacitacion','i.instructor','asica.fecha_recibida'
        );

    if ($nombre)       $query->where('emp.nombre_completo', $nombre);
    if ($puesto) {
        $query->where(function ($q) use ($puesto) {
            $q->where('ptm.puesto_trabajo_matriz', $puesto)
                ->orWhere('pt.puesto_trabajo', $puesto);
        });
    }
    if ($fecha)        $query->where('asica.fecha_recibida',$fecha);
    if ($capacitacion) $query->where('capa.capacitacion',   $capacitacion);

    $capacitaciones = $query->get();

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Consultas');

    // Logo
    foreach ([public_path('img/logo.PNG'), public_path('img/logo.png'),
              public_path('logo.png'), public_path('logo.jpg')] as $logoPath) {
        if (is_file($logoPath)) {
            $drawing = new Drawing();
            $drawing->setName('Logo');
            $drawing->setPath($logoPath);
            $drawing->setHeight(60);
            $drawing->setCoordinates('A1');
            $drawing->setOffsetX(5);
            $drawing->setOffsetY(5);
            $drawing->setWorksheet($sheet);
            break;
        }
    }

    // Encabezado
    $empresa     = 'SERVICE AND TRIDING BUSINESS';
    $descripcion = 'LISTADO DE BUSQUEDA DE CAPACITACIONES';

    $sheet->mergeCells('B1:F2');
    $sheet->setCellValue('B1', $empresa);
    $sheet->getStyle('B1')->getFont()->setBold(true)->setSize(16);
    $sheet->getStyle('B1')->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ->setVertical(Alignment::VERTICAL_CENTER);

    $sheet->mergeCells('B3:F3');
    $sheet->setCellValue('B3', $descripcion);
    $sheet->getStyle('B3')->getFont()->setBold(true)->setSize(16);
    $sheet->getStyle('B3')->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ->setVertical(Alignment::VERTICAL_CENTER);

    // Encabezados de la tabla
    $headerRow = 5;
    $headers = ['Nombre Completo','Puesto de Trabajo','Departamento','Capacitacion','Instructor','Fecha'];
    foreach ($headers as $i => $h) {
        $col = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i+1);
        $sheet->setCellValue($col.$headerRow, $h);
    }
    $headerRange = "A{$headerRow}:F{$headerRow}";
    $sheet->getStyle($headerRange)->getFill()
        ->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF2563EB');
    $sheet->getStyle($headerRange)->getFont()->setBold(true)->getColor()->setARGB('FFFFFFFF');
    $sheet->getStyle($headerRange)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

    // Datos
    $row = $headerRow + 1;
    foreach ($capacitaciones as $item) {
        $sheet->setCellValue("A{$row}", $item->nombre_completo);
        $sheet->setCellValue("B{$row}", $item->puesto_trabajo);
        $sheet->setCellValue("C{$row}", $item->departamento);
        $sheet->setCellValue("D{$row}", $item->capacitacion);
        $sheet->setCellValue("E{$row}", $item->instructor);

        // Fecha como fecha Excel (para que formatee bien)
        if (!empty($item->fecha_recibida)) {
            try {
                $dt = new \DateTime($item->fecha_recibida);
                $sheet->setCellValue("F{$row}", ExcelDate::PHPToExcel($dt));
            } catch (\Throwable $e) {
                $sheet->setCellValue("F{$row}", $item->fecha_recibida);
            }
        } else {
            $sheet->setCellValue("F{$row}", '');
        }

        $row++;
    }

    // Bordes + formato de fecha
    $lastRow = $row - 1;
    if ($lastRow >= $headerRow) {
        $tableRange = "A{$headerRow}:F{$lastRow}";
        $sheet->getStyle($tableRange)->getBorders()->getAllBorders()
              ->setBorderStyle(Border::BORDER_THIN)
              ->getColor()->setARGB('FF9CA3AF');

        $sheet->getStyle("F".($headerRow+1).":F{$lastRow}")
              ->getNumberFormat()->setFormatCode('dd/mm/yyyy');
    }

    // Auto ancho
    foreach (range('A','F') as $col) {
        $sheet->getColumnDimension($col)->setAutoSize(true);
    }

    // ===== Descarga robusta =====
    $fileName = 'consultas_capacitaciones_'.date('Ymd_His').'.xlsx';
    $writer   = new Xlsx($spreadsheet);

    if (function_exists('ini_set')) { @ini_set('zlib.output_compression', 'Off'); }
    while (ob_get_level() > 0) { @ob_end_clean(); }

    $tmp = tempnam(sys_get_temp_dir(), 'cap_');
    $writer->save($tmp);

    return response()->download($tmp, $fileName, [
            'Content-Type'              => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Cache-Control'             => 'max-age=0, no-cache, no-store, must-revalidate',
            'Pragma'                    => 'public',
            'Content-Transfer-Encoding' => 'binary',
        ])->deleteFileAfterSend(true);
}

}
