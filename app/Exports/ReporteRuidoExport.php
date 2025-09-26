<?php

namespace App\Exports;

use Illuminate\Support\Facades\DB;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ReporteRuidoExport
{
    public function __construct(
        public ?int $year = null,
        public ?int $locId = null
    ) {}

    public function build(): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Reporte Ruido');

        // ===== Colores / estilos
        $colorHeader      = '00B0F0';
        $colorHeaderDark  = '0088BC';
        $colorBorder      = '000000';
        $colorWhite       = 'FFFFFF';

        $thinBorders = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color'       => ['rgb' => $colorBorder],
                ],
            ],
        ];

        $centerBold = [
            'font' => ['bold' => true],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical'   => Alignment::VERTICAL_CENTER,
                'wrapText'   => true,
            ],
        ];

        $centerBoldWhite = [
            'font' => ['bold' => true, 'color' => ['rgb' => $colorWhite]],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical'   => Alignment::VERTICAL_CENTER,
                'wrapText'   => true,
            ],
            'fill' => [
                'fillType'   => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $colorHeader],
            ],
        ];

        $centerBoldWhiteDark          = $centerBoldWhite;
        $centerBoldWhiteDark['fill']['startColor']['rgb'] = $colorHeaderDark;

        $left = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical'   => Alignment::VERTICAL_CENTER,
                'wrapText'   => true,
            ],
        ];

        $right = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
                'vertical'   => Alignment::VERTICAL_CENTER,
            ],
        ];

        $center = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical'   => Alignment::VERTICAL_CENTER,
            ],
        ];

        $small = [
            'font' => ['size' => 9],
        ];

        // ===== Anchos
        $sheet->getColumnDimension('A')->setWidth(6);
        $sheet->getColumnDimension('B')->setWidth(36);
        $sheet->getColumnDimension('C')->setWidth(36);
        $sheet->getColumnDimension('D')->setWidth(14);
        $sheet->getColumnDimension('E')->setWidth(14);
        $sheet->getColumnDimension('F')->setWidth(14);
        $sheet->getColumnDimension('G')->setWidth(12);
        $sheet->getColumnDimension('H')->setWidth(12);
        $sheet->getColumnDimension('I')->setWidth(16);
        $sheet->getColumnDimension('J')->setWidth(46);

        // ===== Encabezado
        $row = 1;
        $sheet->mergeCells("A{$row}:J{$row}")->setCellValue("A{$row}", 'SERVICE AND TRADING BUSINESS S.A. DE C.V.');
        $row++;
        $sheet->mergeCells("A{$row}:J{$row}")->setCellValue("A{$row}", 'PROCESO DE SALUD Y SEGURIDAD OCUPACIONAL / OCCUPATIONAL HEALTH AND SAFETY PROCESS');
        $row++;
        $sheet->mergeCells("A{$row}:J{$row}")->setCellValue("A{$row}", 'RESUMEN DE MEDICIONES DE RUIDO / SUMMARY OF NOISE MEASUREMENTS');
        $sheet->getStyle('A1:J3')->applyFromArray($centerBold);
        $sheet->getRowDimension(1)->setRowHeight(22);
        $sheet->getRowDimension(2)->setRowHeight(22);
        $sheet->getRowDimension(3)->setRowHeight(26);

        if ($this->year) {
            $row++;
            $sheet->mergeCells("A{$row}:J{$row}")->setCellValue("A{$row}", 'Anio: ' . $this->year);
            $sheet->getStyle("A{$row}:J{$row}")->applyFromArray($left);
        } else {
            $row++;
        }

        // ===== Localizaciones
        $locNames = DB::table('localizacion')->pluck('localizacion', 'id_localizacion');

        // ===== Datos base + área del puesto
        $records = DB::table('mediciones_ruido as m')
            ->leftJoin('puesto_trabajo_matriz as p', 'p.id_puesto_trabajo_matriz', '=', 'm.id_puesto_trabajo_matriz')
            ->leftJoin('area as a', 'a.id_area', '=', 'p.id_area')
            ->select(
                'm.id_mediciones_ruido as id',
                'm.id_localizacion',
                'm.punto_medicion',
                'p.puesto_trabajo_matriz as puesto',
                'a.area as area_nombre',
                'm.nivel_maximo',
                'm.nivel_minimo',
                'm.nivel_promedio',
                'm.limites_aceptables',
                'm.acciones_correctivas'
            )
            ->when($this->locId, fn($q) => $q->where('m.id_localizacion', $this->locId))
            ->when($this->year, fn($q) => $q->whereRaw('YEAR(COALESCE(m.fecha_realizacion_inicio, m.fecha_realizacion_final)) = ?', [$this->year]))
            ->orderBy('m.id_localizacion')
            ->orderBy('m.punto_medicion')
            ->orderBy('m.id_mediciones_ruido')
            ->get();

        // ===== Cálculos (PROM, NRR, NRE)
        $records = $records->map(function ($r) {
            $max = is_numeric($r->nivel_maximo) ? (float)$r->nivel_maximo : null;
            $min = is_numeric($r->nivel_minimo) ? (float)$r->nivel_minimo : null;

            // PROMEDIO
            if ($max !== null && $min !== null) {
                $prom = ($max + $min) / 2.0;
            } else {
                $prom = is_numeric($r->nivel_promedio) ? (float)$r->nivel_promedio : null;
            }

            // Límite (default 85)
            $lim  = 80.0;

            // Reducción por EPP (si excede límite)
            $nrr  = null;
            $nre  = $prom;

            if ($prom !== null && $prom > $lim) {
                // Área Interna → Orejeras (NRR 13.5); otro caso → Tapones (NRR 11.24)
                $nrr = (strcasecmp((string)$r->area_nombre, 'Area Interna') === 0) ? 22.00 : 27.00;
                $nre = $prom - $nrr;
            }

            $r->calc_promedio = $prom === null ? null : round($prom, 2);
            $r->calc_nrr      = $nrr  === null ? null : round($nrr, 2);
            $r->calc_nre      = $nre  === null ? null : round($nre, 2);
            $r->lim_final     = $lim;

            return $r;
        });

        $groups = $records->groupBy('id_localizacion');

        // ===== Secciones
        $hasRows = false;
        foreach ($locNames as $locId => $locName) {
            $rows = $groups->get($locId);
            if (!$rows || $rows->isEmpty()) continue;

            $hasRows = true;
            $row++;
            $sheet->mergeCells("A{$row}:J{$row}");
            $sheet->setCellValue("A{$row}", Str::upper($locName));
            $sheet->getStyle("A{$row}:J{$row}")->applyFromArray($centerBold);
            $sheet->getStyle("A{$row}:J{$row}")->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB($colorHeader);
            $sheet->getRowDimension($row)->setRowHeight(20);

            // Encabezados
            $row++;
            $firstHeaderRow = $row;
            $sheet->setCellValue("A{$row}", 'No.');
            $sheet->setCellValue("B{$row}", 'ZONA MEDICION');
            $sheet->setCellValue("C{$row}", 'PUESTO DE TRABAJO');
            $sheet->setCellValue("D{$row}", 'NIVEL DE RUIDO');
            $sheet->mergeCells("D{$row}:F{$row}");
            $sheet->setCellValue("G{$row}", 'NRR');
            $sheet->setCellValue("H{$row}", 'NRE');
            $sheet->setCellValue("I{$row}", 'LIMITES ACEPTABLES');
            $sheet->setCellValue("J{$row}", 'ACCIONES CORRECTIVAS');
            $sheet->getStyle("A{$row}:J{$row}")->applyFromArray($centerBoldWhite);
            $sheet->getRowDimension($row)->setRowHeight(22);

            $row++;
            $sheet->setCellValue("D{$row}", 'MAXIMO');
            $sheet->setCellValue("E{$row}", 'MINIMO');
            $sheet->setCellValue("F{$row}", 'PROMEDIO');
            $sheet->getStyle("A{$row}:J{$row}")->applyFromArray($centerBoldWhiteDark);
            $sheet->getRowDimension($row)->setRowHeight(20);
            $sheet->getStyle('A' . $firstHeaderRow . ":J{$row}")->applyFromArray($thinBorders);

            // Filas
            $n = 1;
            foreach ($rows as $data) {
                // Texto de EPP sugerido si excede límite
                $acciones = $data->acciones_correctivas;
                $eppTxt   = null;

                if ($data->calc_promedio !== null && $data->lim_final !== null && $data->calc_promedio > $data->lim_final) {
                    $isInterna = strcasecmp((string)($data->area_nombre ?? ''), 'Area Interna') === 0;

                    if ($acciones === null || trim($acciones) === '') {
                        $acciones = 'Uso obligatorio de protección auditiva: ';
                    }
                }

                $row++;
                $sheet->setCellValue("A{$row}", $n++);
                $sheet->setCellValue("B{$row}", $data->punto_medicion ?: '');
                $sheet->setCellValue("C{$row}", $data->puesto ?: '');
                $sheet->setCellValue("D{$row}", $this->fmt($data->nivel_maximo));
                $sheet->setCellValue("E{$row}", $this->fmt($data->nivel_minimo));
                $sheet->setCellValue("F{$row}", $this->fmt($data->calc_promedio)); // promedio calculado
                $sheet->setCellValue("G{$row}", $this->fmt($data->calc_nrr));       // NRR
                $sheet->setCellValue("H{$row}", $this->fmt($data->calc_nre));       // NRE
                $sheet->setCellValue("I{$row}", $this->fmt($data->lim_final, 0));   // Límite
                $sheet->setCellValue("J{$row}", $acciones ?: '');

                $sheet->getStyle("A{$row}")->applyFromArray($centerBold);
                $sheet->getStyle("B{$row}:C{$row}")->applyFromArray($left);
                $sheet->getStyle("D{$row}:I{$row}")->applyFromArray($right);
                $sheet->getStyle("A{$row}:J{$row}")->applyFromArray($thinBorders);
                $sheet->getRowDimension($row)->setRowHeight(18);
            }

            // Espacio entre localizaciones
            $row++;
        }

        if (!$hasRows) {
            $row++;
            $sheet->mergeCells("A{$row}:J{$row}");
            $sheet->setCellValue("A{$row}", 'Sin mediciones de ruido para los filtros seleccionados.');
            $sheet->getStyle("A{$row}:J{$row}")->applyFromArray($centerBold);
        }

        // ===== Pie de página
        $row += 2;

        $sheet->setCellValue("B{$row}", '1 Copia Archivo');
        $sheet->getStyle("B{$row}")->applyFromArray($small)->applyFromArray($left);
        $row++;

        $sheet->setCellValue("B{$row}", '1 Copia Sistema');
        $sheet->getStyle("B{$row}")->applyFromArray($small)->applyFromArray($left);
        $row++;

        $sheet->setCellValue("B{$row}", 'VERSION 2018');
        $sheet->getStyle("B{$row}")->applyFromArray($small)->applyFromArray($center);

        $sheet->setCellValue("C{$row}", 'STB/SSO/R040');
        $sheet->getStyle("C{$row}")->applyFromArray($small)->applyFromArray($left);

        $sheet->getRowDimension($row-2)->setRowHeight(16);
        $sheet->getRowDimension($row-1)->setRowHeight(16);
        $sheet->getRowDimension($row  )->setRowHeight(16);

        return $spreadsheet;
    }

    private function fmt($value, int $decimals = 2): string
    {
        if ($value === null || $value === '') return '';
        if (!is_numeric($value)) return (string) $value;
        return number_format((float) $value, $decimals);
    }
}
