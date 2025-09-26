<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

class RegistroAsistenciaCapController extends Controller
{
    public function index(Request $request)
    {
        // Catálogo para el <select> de capacitaciones (capacitacion_instructor)
        // Nota: si tu tabla de instructores tiene otro nombre/campo, ajusta el COALESCE() abajo.
        $cis = DB::table('capacitacion_instructor as ci')
            ->join('capacitacion as c', 'c.id_capacitacion', '=', 'ci.id_capacitacion')
            ->leftJoin('instructor as ins', 'ins.id_instructor', '=', 'ci.id_instructor') // opcional
            ->selectRaw('ci.id_capacitacion_instructor as id, c.capacitacion as tema, ci.duracion, ci.programada, ci.id_instructor, ins.instructor as instructor_nombre')
            ->orderBy('c.capacitacion')
            ->get();

        return view('capacitaciones.registro_asistencia', [
            'cis' => $cis,
            'ok'  => session('ok'),
            'err' => session('err'),
        ]);
    }

    // Devuelve JSON con info para prellenar (duración, programada, instructor, tema)
    public function ciInfo(Request $request)
    {
        $id = (int) $request->query('id');
        if ($id <= 0) return response()->json(['ok' => false, 'error' => 'id requerido'], 422);

        $row = DB::table('capacitacion_instructor as ci')
            ->join('capacitacion as c', 'c.id_capacitacion', '=', 'ci.id_capacitacion')
            ->leftJoin('instructor as ins', 'ins.id_instructor', '=', 'ci.id_instructor') // opcional
            ->where('ci.id_capacitacion_instructor', $id)
            ->first([
                'c.capacitacion as tema',
                'ci.duracion',
                'ci.programada',
                'ci.id_instructor',
                DB::raw('COALESCE(ins.nombre, ins.instructor) as instructor_nombre')
            ]);

        if (!$row) return response()->json(['ok'=>false,'error'=>'no encontrado'], 404);

        return response()->json([
            'ok' => true,
            'data' => [
                'tema'               => $row->tema,
                'duracion'           => $row->duracion,         // minutos (como en DB)
                'programada'         => trim(strtoupper($row->programada ?? '')) === 'SI' ? 'SI' : 'NO',
                'id_instructor'      => $row->id_instructor,
                'instructor_nombre'  => $row->instructor_nombre,
            ]
        ]);
    }

    // Autocompletar empleados por nombre o identidad
    public function empleadoSearch(Request $request)
    {
        $q = trim((string)$request->query('q', ''));
        if ($q === '') return response()->json(['ok'=>true,'data'=>[]]);

        $rows = DB::table('empleado as e')
            ->select('e.id_empleado as id','e.nombre_completo','e.identidad')
            ->where(function ($w) use ($q) {
                $w->where('e.nombre_completo', 'like', "%{$q}%")
                  ->orWhere('e.identidad', 'like', "%{$q}%");
            })
            ->where(function ($w) {
                $w->where('e.estado', 1)->orWhereNull('e.estado');
            })
            ->limit(12)
            ->get();

        return response()->json(['ok'=>true,'data'=>$rows]);
    }

    public function store(Request $request)
    {
        $request->validate([
            'id_ci'               => 'required|integer|exists:capacitacion_instructor,id_capacitacion_instructor',
            'fecha_recibida'      => 'required|string|max:50', // varchar libre (p.ej. "Del 12/07/2025 al 15/07/2025")
            'instructor_temporal' => 'nullable|string|max:100',
            'empleados'           => 'required|array|min:1',
            'empleados.*'         => 'integer|exists:empleado,id_empleado',
        ], [
            'empleados.required'  => 'Agrega al menos un empleado.',
        ]);

        $idCi   = (int)$request->input('id_ci');
        $fecha  = trim((string)$request->input('fecha_recibida'));
        $instrT = trim((string)$request->input('instructor_temporal', ''));
        $emps   = $request->input('empleados', []);
        $doPrint = (bool)$request->input('print', false);

        try {
            $rows = [];
            foreach ($emps as $empId) {
                $rows[] = [
                    'id_empleado'                 => (int)$empId,
                    'id_capacitacion_instructor'  => $idCi,
                    'instructor_temporal'         => $instrT !== '' ? $instrT : null,
                    'fecha_recibida'              => $fecha,
                ];
            }
            DB::table('asistencia_capacitacion')->insert($rows);

            return redirect()
                ->route('capacitaciones.registro-asistencia', ['print' => $doPrint ? 1 : 0])
                ->with('ok', 'Asistencia registrada: '.count($rows).' fila(s).');
        } catch (\Throwable $e) {
            return back()->withInput()->with('err', 'No se pudo guardar: '.$e->getMessage());
        }
    }
}
