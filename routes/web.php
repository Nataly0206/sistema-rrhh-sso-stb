<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\Capacitaciones;
use App\Http\Controllers\EPP;
use App\Http\Controllers\Generales;
use App\Http\Controllers\Organigrama;
use App\Http\Controllers\Prestamos;
use App\Http\Controllers\EmpleadoImportController;
use App\Http\Controllers\PuestoImportController;
use App\Http\Controllers\ExportarAsistencia;
use App\Http\Controllers\InicioCapacitaciones;
use App\Http\Controllers\AsistenciaImportController;
use App\Http\Controllers\EPPImportController;
use App\Http\Controllers\PrestamoImportController;
use App\Http\Controllers\ReportePrestamosController;
use App\Http\Controllers\RiesgosController;
use App\Http\Controllers\InicioEPP;
use App\Http\Controllers\QuimicosImportController;
use App\Http\Controllers\EvaluacionRiesgoImportController;
use App\Http\Controllers\RiesgoValorImportController;
use App\Http\Controllers\ExportPrestamosController;
use App\Http\Controllers\QuimicoPuestoImportController;
use App\Http\Controllers\MedidasRiesgoImportController;
Use App\Http\Controllers\MatrizQuimicosController;
Use App\Http\Controllers\DetallesRiesgoImportController;
use App\Http\Controllers\EppObligatoriosConsultaController;
use App\Http\Controllers\VerificacionRiesgosController;
use App\Http\Controllers\AnalisisRiesgosPuestoController;
use App\Http\Controllers\MatrizPuestosEppController;
use App\Http\Controllers\MatrizPuestosCapacitacionController;
use App\Http\Controllers\ComparacionPuestosController;
use App\Http\Controllers\CapacitacionesObligatoriasConsultaController;
use App\Http\Controllers\CapacitacionInstructorImportController;
use App\Http\Controllers\IRFormController;
use App\Http\Controllers\EvaluacionRiesgosExportController;
use App\Http\Controllers\IdentificacionRiesgosExportController;
use App\Http\Controllers\PuestoTrabajoMatrizController;
use App\Http\Controllers\IdentificacionRiesgosController;
use App\Http\Controllers\NotificacionRiesgoExcelTemplateController;
use App\Http\Controllers\AjustesPrestamosController;
use App\Http\Controllers\MedicionesController;
use App\Http\Controllers\EppRequeridosController;
use App\Http\Controllers\PrestamoRenunciaController;
use App\Http\Controllers\PrestamosResumenController;
use App\Http\Controllers\InformeCapacitacionesMensualController;
use App\Http\Controllers\InformeCapacitacionesAnualController;
use App\Http\Controllers\RegistroAsistenciaCapController;

Route::middleware(['auth'])->group(function () {

    Route::get ('/capacitaciones/registro-asistencia',        [RegistroAsistenciaCapController::class, 'index'])->name('capacitaciones.registro-asistencia');
    Route::post('/capacitaciones/registro-asistencia',        [RegistroAsistenciaCapController::class, 'store'])->name('capacitaciones.registro-asistencia.store');

    Route::get ('/capacitaciones/ci-info',                    [RegistroAsistenciaCapController::class, 'ciInfo'])->name('capacitaciones.ci-info');
    Route::get ('/capacitaciones/empleados/search',           [RegistroAsistenciaCapController::class, 'empleadoSearch'])->name('capacitaciones.empleados.search');

    Route::get('/capacitaciones/informe-anual', [InformeCapacitacionesAnualController::class, 'index'])
    ->name('capacitaciones.informe-anual');
    Route::get('/capacitaciones/informe-anual/export', [InformeCapacitacionesAnualController::class, 'export'])
    ->name('capacitaciones.informe-anual.export');

    Route::get('/capacitaciones/informe-mensual', [InformeCapacitacionesMensualController::class, 'index'])
    ->name('capacitaciones.informe-mensual');
    Route::get('/capacitaciones/informe-mensual/export', [InformeCapacitacionesMensualController::class, 'export'])
    ->name('capacitaciones.informe-mensual.export');

    Route::get('/evaluacion-riesgos/puestos', [EvaluacionRiesgosExportController::class, 'puestos'])->name('evaluacion.riesgos.puestos');
    Route::post('/evaluacion-riesgos/export', [EvaluacionRiesgosExportController::class, 'export'])->name('evaluacion.riesgos.export');
    Route::post('/identificacion-riesgos/export', [IdentificacionRiesgosExportController::class, 'export'])->name('identificacion.riesgos.export');
    
    Route::get('/', function () {
        return view('index');
    })->name('home');

    Route::get('/prestamos/reporte/excel', [ExportPrestamosController::class, 'export'])
    ->name('prestamos.reporte.excel');

    Route::get('/Capacitaciones', [InicioCapacitaciones::class, 'Capacitaciones']);

    //Route::get('/epp', function () {
        //return view('epp.inicioepp');
    //})->name('EPP');

    Route::get('/epp', [InicioEPP::class, 'dashboard']);

    //Route::get('/Riesgos', function () {
        //return view('riesgos.inicioriesgos');
    //})->name('Riesgos');

// RUTAS PARA ACCEDER A LOS DATOS GENERALES
    Route::get('/empleado', [Generales::class, 'empleado']);
    Route::get('/puestos', [Generales::class, 'puestos']);
    Route::get('/departamento', [Generales::class, 'departamento']);
    Route::get('/localizacion', [Generales::class, 'localizacion']);
    Route::get('/area', [Generales::class, 'area']);

// RUTAS PARA ACCEDER A LOS DATOS DE LOS EPP
    // Control de Entrega de EPP
    Route::get('/controlentrega', [EPP::class, 'controlentrega']);
    Route::post('/controlentrega', [App\Http\Controllers\EPP::class, 'store'])->name('controlentrega.store');
    Route::delete('/controlentrega/{id}', [App\Http\Controllers\EPP::class, 'destroy'])->name('controlentrega.destroy');
    Route::post('/controlentrega', [App\Http\Controllers\EPP::class, 'store'])->name('controlentrega.store');
    Route::delete('/controlentrega/{id}', [App\Http\Controllers\EPP::class, 'destroy'])->name('controlentrega.destroy');
    Route::post('/controlentrega/importar', [EPPImportController::class, 'import'])->name('controlentrega.import');
    // Ruta para B�squeda
    Route::get('/controlentrega', [EPP::class, 'controlentrega'])->name('controlentrega');
    Route::get('/epp/imprimir', [App\Http\Controllers\EPP::class, 'imprimir'])->name('epp.imprimir');

    // Equipo
    Route::get('/equipo', [EPP::class, 'equipo']);
    Route::post('/equipo/storeequipo', [App\Http\Controllers\EPP::class, 'storeequipo'])->name('equipo.storeequipo');
    Route::put('/equipo/updateequipo/{id}', [App\Http\Controllers\EPP::class, 'updateequipo'])->name('equipo.updateequipo');
    Route::delete('/equipo/{id}', [App\Http\Controllers\EPP::class, 'destroyequipo'])->name('equipo.destroyequipo');
    // Ruta para B�squeda
    Route::get('/equipo', [EPP::class, 'equipo'])->name('equipo');

    // Tipo de Proteccion
    Route::get('/tipoproteccion', [EPP::class, 'tipoproteccion']);
    Route::post('/tipoproteccion/storetipo', [App\Http\Controllers\EPP::class, 'storetipo'])->name('tipoproteccion.storetipo');
    Route::put('/tipoproteccion/updatetipo/{id}', [App\Http\Controllers\EPP::class, 'updatetipo'])->name('tipoproteccion.updatetipo');
    Route::delete('/tipoproteccion/{id}', [App\Http\Controllers\EPP::class, 'destroytipo'])->name('tipoproteccion.destroytipo');
    // Ruta para B�squeda
    Route::get('/tipoproteccion', [EPP::class, 'tipoproteccion'])->name('tipoproteccion');

 // RUTAS PARA ACCEDER A LOS DATOS DEL ORGANIGRAMA
    // Matriz de Puestos
    Route::get('/matrizpuestos', [Organigrama::class, 'matrizpuestos']);
    Route::post('/matrizpuestos/storematrizpuestos', [App\Http\Controllers\Organigrama::class, 'storematrizpuestos'])->name('matrizpuestos.storematrizpuestos');
    Route::put('/matrizpuestos/updatematrizpuestos/{id}', [App\Http\Controllers\Organigrama::class, 'updatematrizpuestos'])->name('matrizpuestos.updatematrizpuestos');
    Route::delete('/matrizpuestos/{id}', [App\Http\Controllers\Organigrama::class, 'destroymatrizpuestos'])->name('matrizpuestos.destroymatrizpuestos');
    Route::get('/matrizpuestos', [Organigrama::class, 'matrizpuestos'])->name('matrizpuestos');

    // Niveles
    Route::get('/niveles', [Organigrama::class, 'niveles']);
    Route::post('/matrizpuestos/storeniveles', [App\Http\Controllers\Organigrama::class, 'storeniveles'])->name('matrizpuestos.storeniveles');
    Route::put('/matrizpuestos/updateniveles/{id}', [App\Http\Controllers\Organigrama::class, 'updateniveles'])->name('matrizpuestos.updateniveles');
    Route::delete('/matrizdepuestos/{id}', [App\Http\Controllers\Organigrama::class, 'destroyniveles'])->name('matrizpuestos.destroyniveles');
    // Ruta para B�squeda
    Route::get('/matrizpuestos', [Organigrama::class, 'matrizpuestos'])->name('matrizpuestos');

 // RUTAS PARA ACCEDER A LOS DATOS DE LOS PRESTAMOS
    Route::get('/empleadosprestamo', [Prestamos::class, 'empleadosprestamo']);
        // Ruta para B�squeda
    Route::get('/empleadosprestamo', [Prestamos::class, 'empleadosprestamo'])->name('empleadosprestamo');

    Route::get('/infoprestamo', [Prestamos::class, 'prestamo']);
            // Ruta para B�squeda
    Route::get('/infoprestamo', [Prestamos::class, 'prestamo'])->name('infoprestamo');
    
    Route::get('/cuotas', [Prestamos::class, 'cuotas']);
            // Ruta para B�squeda
    Route::get('/cuotas', [Prestamos::class, 'cuotas'])->name('cuotas');
    Route::get('/cuotas/rango', [Prestamos::class, 'cuotasPorRango'])->name('cuotas.rango');
    Route::post('/cuotas/import-ajustes', [Prestamos::class, 'importAjustesCuotas'])->name('cuotas.import');
    Route::put('/cuotas/{id}', [Prestamos::class, 'updateCuota'])->name('cuotas.update');
    Route::get('/cuotas-especiales', [Prestamos::class, 'cuotasEspeciales'])->name('cuotas.especiales');
    Route::delete('/cuotas/{id}', [Prestamos::class, 'destroyCuota'])->name('cuotas.destroy');

 // RUTAS PARA ACCEDER A LOS DATOS DE LOS EMPLEADOS
    // Route::post('/empleado', [EmpleadoController::class, 'store'])->name('empleados.store');
    Route::post('/empleados/importar', [EmpleadoImportController::class, 'import'])->name('empleados.import');
    Route::post('/empleados/store', [App\Http\Controllers\Generales::class, 'store'])->name('empleado.store');
    Route::put('/empleados/update/{id}', [App\Http\Controllers\Generales::class, 'update'])->name('empleado.update');
    Route::delete('/empleados/{id}', [App\Http\Controllers\Generales::class, 'destroy'])->name('empleados.destroy');
    // Ruta para B�squeda
    Route::get('/empleados', [Generales::class, 'empleado'])->name('empleados');

// RUTAS PARA ACCEDER A LOS DATOS DE LOS PUESTOS DE TRABAJO
    Route::post('/puestos/import', [PuestoImportController::class, 'import'])->name('puestos.import');
    Route::post('/puestos/storepuestos', [App\Http\Controllers\Generales::class, 'storepuestos'])->name('puestos.storepuestos');
    Route::put('/puestos/updatepuestos/{id}', [App\Http\Controllers\Generales::class, 'updatepuestos'])->name('puestos.updatepuestos');
    Route::delete('/puestos/{id}', [App\Http\Controllers\Generales::class, 'destroypuestos'])->name('puestos.destroypuestos');
    // Ruta para B�squeda
    Route::get('/puestos', [Generales::class, 'puestos'])->name('puestos');

// RUTAS PARA PUESTOS DE TRABAJO DEL SISTEMA
    Route::get('/puestossistemas', [Generales::class, 'puestossistema'])->name('puestossistema');
    // Ruta para B�squeda
    Route::get('/puestossistemas', [Generales::class, 'puestossistema'])->name('puestossistemas');

 // RUTAS PARA ACCEDER A LOS DATOS DE LOS DEPARTAMENTOS
    Route::post('/departamento/storedepartamento', [App\Http\Controllers\Generales::class, 'storedepartamento'])->name('departamento.storedepartamento');
    Route::put('/departamento/updatedepartamento/{id}', [App\Http\Controllers\Generales::class, 'updatedepartamento'])->name('departamento.updatedepartamento');
    Route::delete('/departamento/{id}', [App\Http\Controllers\Generales::class, 'destroydepartamento'])->name('departamento.destroydepartamento');
    // Ruta para B�squeda
    Route::get('/departamento', [Generales::class, 'departamento'])->name('departamento');

 // RUTAS PARA ACCEDER A LOS DATOS DE LAS LOCALIZACIONES
    Route::post('/localizacion/storelocalizacion', [App\Http\Controllers\Generales::class, 'storelocalizacion'])->name('localizacion.storelocalizacion');
    Route::put('/localizacion/updatelocalizacion/{id}', [App\Http\Controllers\Generales::class, 'updatelocalizacion'])->name('localizacion.updatelocalizacion');
    Route::delete('/localizacion/{id}', [App\Http\Controllers\Generales::class, 'destroylocalizacion'])->name('localizacion.destroylocalizacion');
    // Ruta para B�squeda
    Route::get('/localizacion', [Generales::class, 'localizacion'])->name('localizacion');

 // RUTAS PARA ACCEDER A LOS DATOS DE LAS AREAS
    Route::post('/area/storearea', [App\Http\Controllers\Generales::class, 'storearea'])->name('area.storearea');
    Route::put('/area/updatearea/{id}', [App\Http\Controllers\Generales::class, 'updatearea'])->name('area.updatearea');
    Route::delete('/area/{id}', [App\Http\Controllers\Generales::class, 'destroyarea'])->name('area.destroyarea');
    // Ruta para B�squeda
    Route::get('/area', [Generales::class, 'area'])->name('area');

 // RUTAS PARA ACCEDER A LOS DATOS DE LAS CAPACITACIONES
    // Capacitaciones
    Route::get('/capacitacion', [Capacitaciones::class, 'Capacitacion']);
    Route::post('/capacitacion/storecapacitacion', [App\Http\Controllers\Capacitaciones::class, 'storecapacitacion'])->name('capacitacion.storecapacitacion');
    Route::put('/capacitacion/updatecapacitacion/{id}', [App\Http\Controllers\Capacitaciones::class, 'updatecapacitacion'])->name('capacitacion.updatecapacitacion');
    Route::delete('/capacitacion/{id}', [App\Http\Controllers\Capacitaciones::class, 'destroycapacitacion'])->name('capacitacion.destroycapacitacion');
    // Ruta para B�squeda
    Route::get('/capacitacion', [Capacitaciones::class, 'capacitacion'])->name('capacitacion');
    Route::get('/capacitaciones/imprimir', [App\Http\Controllers\Capacitaciones::class, 'imprimir'])->name('capacitaciones.imprimir');

    // Consultas de Capacitaciones
    Route::get('/consultascapacitaciones', [Capacitaciones::class, 'consulta'])->name('capacitaciones.consultas');
    // Consultas de EPP
    Route::get('/consultas', [EPP::class, 'consulta'])->name('epp.consultas');
    Route::get('/epp/consultas/imprimir', [App\Http\Controllers\EPP::class, 'imprimirConsultas'])->name('epp.consultas.imprimir');

 // RUTAS PARA ACCEDER A LOS DATOS DE LOS INSTRUCTORES
    // Instructores
    Route::get('/instructor', [Capacitaciones::class, 'Instructor']);
    Route::post('/instructor/storeinstructor', [App\Http\Controllers\Capacitaciones::class, 'storeinstructor'])->name('instructor.storeinstructor');
    Route::put('/instructor/updateinstructor/{id}', [App\Http\Controllers\Capacitaciones::class, 'updateinstructor'])->name('instructor.updateinstructor');
    Route::delete('/instructor/{id}', [App\Http\Controllers\Capacitaciones::class, 'destroyinstructor'])->name('instructor.destroyinstructor');
    // Ruta para B�squeda
    Route::get('/instructor', [Capacitaciones::class, 'instructor'])->name('instructor');

    // Detalles Capacitaciones
    Route::get('/detallescapacitacion', [Capacitaciones::class, 'Capinstructor']);
    Route::post('/detallescapacitacion/storecapinstructor', [App\Http\Controllers\Capacitaciones::class, 'storecapinstructor'])->name('detallescapacitacion.storecapinstructor');
    Route::put('/detallescapacitacion/updateinstructor/{id}', [App\Http\Controllers\Capacitaciones::class, 'updatecapinstructor'])->name('detallescapacitacion.updatecapinstructor');
    Route::delete('/detallescapacitacion/{id}', [App\Http\Controllers\Capacitaciones::class, 'destroycapinstructor'])->name('detallescapacitacion.destroycapinstructor');
    // Ruta para B�squeda
    Route::get('/detallescapacitacion', [Capacitaciones::class, 'Capinstructor'])->name('detallescapacitacion');
    Route::post('/detallescapacitacion/importar', [CapacitacionInstructorImportController::class, 'import'])->name('detallescapacitacion.import');

    // Asistencia a capacitaciones
    Route::delete('/asistencia/{id}', [App\Http\Controllers\Capacitaciones::class, 'destroyasistencia'])->name('asistencia.destroyasistencia');
    Route::post('/asistencia/importar', [AsistenciaImportController::class, 'import'])->name('asistencia.import');
    Route::get('/asistencia', [Capacitaciones::class, 'Asistencia'])->name('Asistencias');
    Route::post('/asistencias', [Capacitaciones::class, 'store'])->name('asistencias.store');
    // Ruta para B�squeda
    Route::get('/asistencia', [Capacitaciones::class, 'Asistencia'])->name('asistencia');

    Route::post('/infoprestamo/storeprestamo', [App\Http\Controllers\PrestamosNuevo::class, 'storeprestamo'])->name('infoprestamo.storeprestamo');
    Route::post('/infoprestamo/importar', [PrestamoImportController::class, 'import'])->name('infoprestamo.import');
    Route::get('/capacitaciones/export/{idEmpleado}', [ExportarAsistencia::class, 'exportCapacitacion'])->name('capacitaciones.export');

    Route::get('/reportes/prestamos/julio', [ReportePrestamosController::class, 'index'])
        ->name('reportes.prestamos.julio');

    Route::get('/reportes/prestamos/julio/descargar', [ReportePrestamosController::class, 'export'])
        ->name('reportes.prestamos.julio.descargar');

    Route::get('/epp/export', [EPP::class, 'exportFormatoEpp'])->name('epp.export');
    Route::get('/home', [App\Http\Controllers\HomeController::class, 'index'])->name('home');

    // RUTAS PARA RIESGOS
    // Tipo de exposici�n
    Route::get('/tipoexposicion', [RiesgosController::class, 'tipoexposicion']);
    Route::post('/tipoexposicion/storetipoexposicion', [App\Http\Controllers\RiesgosController::class, 'storetipoexposicion'])->name('matrizpuestos.storetipoexposicion');
    Route::put('/tipoexposicion/updatetipoexposicion/{id}', [App\Http\Controllers\RiesgosController::class, 'updatetipoexposicion'])->name('matrizpuestos.updatetipoexposicion');
    Route::delete('/tipoexposicion/{id}', [App\Http\Controllers\RiesgosController::class, 'destroytipoexposicion'])->name('matrizpuestos.destroytipoexposicion');
    // Ruta para B�squeda
    Route::get('/tipoexposicion', [RiesgosController::class, 'tipoexposicion'])->name('tipoexposicion');

    // Quimicos
    Route::get('/quimicos', [RiesgosController::class, 'quimicos']);
    Route::post('/quimicos/storequimicos', [App\Http\Controllers\RiesgosController::class, 'storequimicos'])->name('quimicos.storequimicos');
    Route::put('/quimicos/updatequimicos/{id}', [App\Http\Controllers\RiesgosController::class, 'updatequimicos'])->name('quimicos.updatequimicos');
    Route::delete('/quimicos/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyquimicos'])->name('quimicos.destroyquimicos');
    // Ruta para B�squeda
    Route::get('/quimicos', [RiesgosController::class, 'quimicos'])->name('quimicos');
    Route::post('/quimicos/importar', [QuimicosImportController::class, 'import'])->name('quimicos.import');
    Route::get('/quimicos/export', [\App\Http\Controllers\RiesgosController::class, 'exportQuimicosExcel'])->name('quimicos.export');

    // Quimicos y Tipo de Exposici�n
    Route::get('/quimicotipoexposicion', [RiesgosController::class, 'quimicotipoexposicion']);
    Route::post('/quimicotipoexposicion/storequimicos', [App\Http\Controllers\RiesgosController::class, 'storequimicotipoexposicion'])->name('quimicotipoexposicion.storequimicotipoexposicion');
    Route::delete('/quimicotipoexposicion/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyquimicotipoexposicion'])->name('quimicotipoexposicion.destroyquimicotipoexposicion');
    // Ruta para B�squeda
    Route::get('/quimicotipoexposicion', [RiesgosController::class, 'quimicotipoexposicion'])->name('quimicotipoexposicion');

    // Quimicos por Puesto de Trabajo
    Route::get('/quimicospuestos', [RiesgosController::class, 'quimicospuestos']);
    Route::post('/quimicospuestos/storequimicospuestos', [App\Http\Controllers\RiesgosController::class, 'storequimicospuestos'])->name('quimicospuestos.storequimicospuestos');
    Route::put('/quimicospuestos/updatequimicospuestos/{id}', [App\Http\Controllers\RiesgosController::class, 'updatequimicospuestos'])->name('quimicospuestos.updatequimicospuestos');
    Route::delete('/quimicospuestos/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyquimicospuestos'])->name('quimicospuestos.destroyquimicospuestos');
    // Ruta para B�squeda
    Route::get('/quimicospuestos', [RiesgosController::class, 'quimicospuestos'])->name('quimicospuestos');
    Route::get('/quimicospuestos/export', [\App\Http\Controllers\RiesgosController::class, 'exportQuimicosPuestosExcel'])->name('quimicospuestos.export');

    Route::get('/senalizacion', [RiesgosController::class, 'senalizacion']);
    Route::post('/senalizacion/storesenalizacion', [App\Http\Controllers\RiesgosController::class, 'storesenalizacion'])->name('senalizacion.storesenalizacion');
    Route::put('/senalizacion/updatesenalizacion/{id}', [App\Http\Controllers\RiesgosController::class, 'updatesenalizacion'])->name('senalizacion.updatesenalizacion');
    Route::delete('/senalizacion/{id}', [App\Http\Controllers\RiesgosController::class, 'destroysenalizacion'])->name('senalizacion.destroysenalizacion');
    // Ruta para B�squeda
    Route::get('/senalizacion', [RiesgosController::class, 'senalizacion'])->name('senalizacion');

    Route::get('/otras', [RiesgosController::class, 'otras']);
    Route::post('/otras/storeotras', [App\Http\Controllers\RiesgosController::class, 'storeotras'])->name('otras.storeotras');
    Route::put('/otras/updateotras/{id}', [App\Http\Controllers\RiesgosController::class, 'updateotras'])->name('otras.updateotras');
    Route::delete('/otras/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyotras'])->name('otras.destroyotras');
    // Ruta para B�squeda
    Route::get('/otras', [RiesgosController::class, 'otras'])->name('otras');

    Route::get('/medidasriesgo', [RiesgosController::class, 'medidasriesgo']);
    Route::post('/medidasriesgo/storemedidasriesgo', [App\Http\Controllers\RiesgosController::class, 'storemedidasriesgo'])->name('medidasriesgo.storemedidasriesgo');
    // Ruta para B�squeda
    Route::get('/medidasriesgo', [RiesgosController::class, 'medidasriesgo'])->name('medidasriesgo');
    Route::put('/medidasriesgo/{id_riesgo}/{id_area}', [App\Http\Controllers\RiesgosController::class, 'updatemedidasriesgo'])->name('medidasriesgo.update');
    Route::delete('/medidasriesgo/{id_riesgo}/{id_area}', [App\Http\Controllers\RiesgosController::class, 'destroymedidasriesgo'])->name('medidasriesgo.destroy');

    Route::get('/estandarilu', [RiesgosController::class, 'estandarilu']);
    Route::post('/estandarilu/storeestandarilu', [App\Http\Controllers\RiesgosController::class, 'storeestandarilu'])->name('estandarilu.storeestandarilu');
    Route::put('/estandarilu/updateestandarilu/{id}', [App\Http\Controllers\RiesgosController::class, 'updateestandarilu'])->name('estandarilu.updateestandarilu');
    Route::delete('/estandarilu/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyestandarilu'])->name('estandarilu.destroyestandarilu');
    // Ruta para B�squeda
    Route::get('/estandarilu', [RiesgosController::class, 'estandarilu'])->name('estandarilu');

    Route::get('/estandarruido', [RiesgosController::class, 'estandarruido']);
    Route::post('/estandarruido/storeestandarruido', [App\Http\Controllers\RiesgosController::class, 'storeestandarruido'])->name('estandarruido.storeestandarruido');
    Route::put('/estandarruido/updateestandarruido/{id}', [App\Http\Controllers\RiesgosController::class, 'updateestandarruido'])->name('estandarruido.updateestandarruido');
    Route::delete('/estandarruido/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyestandarruido'])->name('estandarruido.destroyestandarruido');
    // Ruta para B�squeda
    Route::get('/estandarruido', [RiesgosController::class, 'estandarruido'])->name('estandarruido');

    Route::get('/estandartemperatura', [RiesgosController::class, 'estandartemperatura']);
    Route::post('/estandartemperatura/storeestandartemperatura', [App\Http\Controllers\RiesgosController::class, 'storeestandartemperatura'])->name('estandartemperatura.storeestandartemperatura');
    Route::put('/estandartemperatura/updateestandartemperatura/{id}', [App\Http\Controllers\RiesgosController::class, 'updateestandartemperatura'])->name('estandartemperatura.updateestandartemperatura');
    Route::delete('/estandartemperatura/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyestandartemperatura'])->name('estandartemperatura.destroyestandartemperatura');
    // Ruta para B�squeda
    Route::get('/estandartemperatura', [RiesgosController::class, 'estandartemperatura'])->name('estandartemperatura');

    Route::get('/estandares', [RiesgosController::class, 'estandares']);
    // Ruta para B�squeda
    Route::get('/estandares', [RiesgosController::class, 'estandares'])->name('estandares');

    Route::get('/tiporiesgo', [RiesgosController::class, 'tiporiesgo']);
    Route::post('/tiporiesgo/storetiporiesgo', [App\Http\Controllers\RiesgosController::class, 'storetiporiesgo'])->name('tiporiesgo.storetiporiesgo');
    Route::put('/tiporiesgo/updatetiporiesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'updatetiporiesgo'])->name('tiporiesgo.updatetiporiesgo');
    Route::delete('/tiporiesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'destroytiporiesgo'])->name('tiporiesgo.destroytiporiesgo');
    // Ruta para B�squeda
    Route::get('/tiporiesgo', [RiesgosController::class, 'tiporiesgo'])->name('tiporiesgo');

    // routes/web.php
    Route::get('/matrizriesgos/export', [\App\Http\Controllers\RiesgosController::class, 'exportMatrizIdentificacionExcel'])
        ->name('matrizriesgos.export');

    // routes/web.php
    Route::post('/riesgos/import', [\App\Http\Controllers\RiesgoValorImportController::class, 'import'])
    ->name('riesgos.import');


    Route::get('/riesgo', [RiesgosController::class, 'riesgo']);
    Route::post('/riesgo/storeriesgo', [App\Http\Controllers\RiesgosController::class, 'storeriesgo'])->name('riesgo.storeriesgo');
    Route::put('/riesgo/updateriesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'updateriesgo'])->name('riesgo.updateriesgo');
    Route::delete('/riesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyriesgo'])->name('riesgo.destroyriesgo');
    // Ruta para B�squeda
    Route::get('/riesgo', [RiesgosController::class, 'riesgo'])->name('riesgo');

    Route::get('/riesgopuesto', [RiesgosController::class, 'riesgopuesto']);
    Route::post('/riesgopuesto/storeriesgopuesto', [App\Http\Controllers\RiesgosController::class, 'storeriesgopuesto'])->name('riesgopuesto.storeriesgopuesto');
    Route::put('/riesgopuesto/updateriesgopuesto/{id}', [App\Http\Controllers\RiesgosController::class, 'updateriesgopuesto'])->name('riesgopuesto.updateriesgopuesto');
    Route::delete('/riesgopuesto/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyriesgopuesto'])->name('riesgopuesto.destroyriesgopuesto');
    // Ruta para B�squeda
    Route::get('/riesgopuesto', [RiesgosController::class, 'riesgopuesto'])->name('riesgopuesto');

    Route::get('/detallesriesgo', [RiesgosController::class, 'detallesriesgo']);
    Route::post('/detallesriesgo/storedetallesriesgo', [App\Http\Controllers\RiesgosController::class, 'storedetallesriesgo'])->name('detallesriesgo.storedetallesriesgo');
    Route::put('/detallesriesgo/updatedetallesriesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'updatedetallesriesgo'])->name('detallesriesgo.updatedetallesriesgo');
    Route::delete('/detallesriesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'destroydetallesriesgo'])->name('detallesriesgo.destroydetallesriesgo');
    // Ruta para B�squeda
    Route::get('/detallesriesgo', [RiesgosController::class, 'detallesriesgo'])->name('detallesriesgo');

    Route::get('/informacionriesgo', [RiesgosController::class, 'informacionriesgo']);
    // Ruta para B�squeda
    Route::get('/informacionriesgo', [RiesgosController::class, 'informacionriesgo'])->name('informacionriesgo');

    Route::get('/valoracionriesgo', [RiesgosController::class, 'valoracionriesgo']);
    Route::post('/valoracionriesgo/storevaloracionriesgo', [App\Http\Controllers\RiesgosController::class, 'storevaloracionriesgo'])->name('valoracionriesgo.storevaloracionriesgo');
    Route::delete('/valoracionriesgo/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyvaloracionriesgo'])->name('valoracionriesgo.destroyvaloracionriesgo');

    Route::get('/evaluacionriesgos', [RiesgosController::class, 'evaluacionriesgos']);
    Route::put('/evaluacionriesgos/{id}', [App\Http\Controllers\RiesgosController::class, 'updateevaluacionriesgos'])->name('evaluacionriesgos.updateevaluacionriesgos');
    Route::post('/evaluacionriesgos/storeevaluacionriesgos', [App\Http\Controllers\RiesgosController::class, 'storeevaluacionriesgos'])->name('evaluacionriesgos.storeevaluacionriesgos');
    Route::delete('/evaluacionriesgos/{id}', [App\Http\Controllers\RiesgosController::class, 'destroyevaluacionriesgos'])->name('evaluacionriesgos.destroyevaluacionriesgos');
    // Ruta para B�squeda
    Route::get('/evaluacionriesgos', [RiesgosController::class, 'evaluacionriesgos'])->name('evaluacionriesgos');
    // routes/web.php
    Route::post('/evaluacion-riesgos/import', [EvaluacionRiesgoImportController::class, 'import'])
    ->name('evaluacion_riesgos.import');

    Route::get('/evaluacion', [RiesgosController::class, 'matrizEvaluacion'])
    ->name('matrizevaluacion');

    Route::get('/Riesgos', [RiesgosController::class, 'index2'])
    ->name('Riesgos');
    Route::post('/matrizpuestos/importar', [Organigrama::class, 'importMatrizpuestos'])
    ->name('matrizpuestos.importar');

    Route::get('/matrizpuestos/exportar', [Organigrama::class, 'exportMatrizpuestos'])
    ->name('matrizpuestos.exportar');

    // Exportar organigrama (imagen -> Excel)
    Route::post('/organigrama/export-excel', [Organigrama::class, 'exportOrganigramaExcel'])
    ->name('organigrama.export_excel');

    Route::get('/matriz-quimicos', [\App\Http\Controllers\MatrizQuimicosController::class, 'index'])
    ->name('riesgos.matrizquimicos');

    Route::post('/quimicospuestos/import', [QuimicoPuestoImportController::class, 'import'])
    ->name('quimicospuestos.import');

    Route::post('/medidas-por-riesgo/import', [MedidasRiesgoImportController::class, 'import'])
    ->name('medidasriesgo.import');

    Route::get('/matriz-quimicos/export', [\App\Http\Controllers\MatrizQuimicosController::class, 'exportExcel'])
    ->name('quimicos.matriz.export');

    // routes/web.php
    Route::get('/evaluacionriesgos/export', [\App\Http\Controllers\RiesgosController::class, 'exportMatrizEvaluacionExcel'])
     ->name('evaluacionriesgos.export');

    Route::get('/detalles-riesgo/import', [DetallesRiesgoImportController::class, 'showImportForm'])->name('detalles_riesgo.import.form');
    Route::post('/detalles-riesgo/import', [DetallesRiesgoImportController::class, 'import'])->name('detalles_riesgo.import');

    Route::get('/analisis-riesgos', [AnalisisRiesgosPuestoController::class, 'index'])
        ->name('riesgos.analisis');

    Route::get('/verificacion', [VerificacionRiesgosController::class, 'index'])
    ->name('riesgos.verificacion');
    Route::get('/verificacion/export', [VerificacionRiesgosController::class, 'export'])
    ->name('riesgos.verificacion.export');
    Route::get('/verificacion/plan-accion', [VerificacionRiesgosController::class, 'exportPlanAccion'])
    ->name('riesgos.verificacion.plan_accion');
    
    Route::get('/epp-obligatorios', [EppObligatoriosConsultaController::class, 'index'])
    ->name('riesgos.epp.obligatorios');
    Route::get('/epp-obligatorios/export', [EppObligatoriosConsultaController::class, 'export'])
    ->name('riesgos.epp.obligatorios.export');
    
    Route::get('/capacitaciones/requeridos', [\App\Http\Controllers\CapacitacionesRequeridasController::class, 'index'])
    ->name('capacitaciones.requeridos');

    Route::get('/medidas/{id_riesgo}/{id_area}/edit', [MedidasRiesgoPuestoController::class, 'edit'])
        ->name('medidas.edit');

    Route::put('/medidas/{id_riesgo}/{id_area}', [MedidasRiesgoPuestoController::class, 'update'])
        ->name('medidas.update');

    Route::get('/puestos-epp', [MatrizPuestosEppController::class, 'index'])->name('epp.puestos_epp');
    Route::post('/puestos-epp.store', [MatrizPuestosEppController::class, 'store'])->name('epp.puestos_epp.store');

    Route::get('/quimicos-por-puesto', [\App\Http\Controllers\QuimicosPorPuestoController::class, 'index'])
        ->name('riesgos.quimicos.puesto');

    Route::get('/puestos-capacitacion',
    [MatrizPuestosCapacitacionController::class, 'index'])
    ->name('capacitaciones.puestoscapacitacion');

    Route::post('/puestos-capacitacion.store',
    [MatrizPuestosCapacitacionController::class, 'store'])
    ->name('capacitaciones.puestoscapacitacion.store');

    Route::get('/comparacion-puestos', [ComparacionPuestosController::class, 'index'])->name('comparacion_puestos.index');
    Route::post('/comparacion-puestos', [ComparacionPuestosController::class, 'store'])->name('comparacion_puestos.store');
    Route::put('/comparacion-puestos/{id}', [ComparacionPuestosController::class, 'update'])->name('comparacion_puestos.update');
    Route::delete('/comparacion-puestos/{id}', [ComparacionPuestosController::class, 'destroy'])->name('comparacion_puestos.destroy');

    Route::get('/capacitaciones-obligatorias', [CapacitacionesObligatoriasConsultaController::class, 'index'])
    ->name('capacitaciones.capacitaciones.obligatorias');

//Route::get('/identificacion-riesgo/crear', [IRFormController::class,'create'])
    //->name('identificacion-riesgo.create');

//Route::post('/identificacion-riesgo/guardar', [IRFormController::class,'store'])
    //->name('identificacion-riesgo.store');

// (Opcional) si alguien abre /guardar con GET, redirige al form:
//Route::get('/identificacion-riesgo/guardar', fn() => redirect()->route('identificacion-riesgo.create'));

// (Opcional) Si luego lo sirves como Blade para verlo con GET:
// Route::view('/identificacion-riesgo/crear', 'identificacion_riesgo.create')
//     ->name('identificacion-riesgo.create');

Route::get('/evaluacion-riesgos/puestos', [EvaluacionRiesgosExportController::class, 'puestos'])
    ->name('evaluacion.riesgos.puestos');

Route::post('/evaluacion-riesgos/exportar', [EvaluacionRiesgosExportController::class, 'export'])
    ->name('evaluacion.riesgos.export');

    Route::get('/identificacion-riesgos', [IdentificacionRiesgosController::class, 'index'])
     ->name('riesgos.identificacion-riesgos');

     Route::post('/identificacion-riesgos', [IdentificacionRiesgosController::class, 'store'])
    ->name('identificacion.store');

    Route::get('/riesgos-fisico', [\App\Http\Controllers\RiesgosReportController::class, 'fisicoPorPuesto'])
    ->name('riesgos.fisico.puesto');

    Route::get('/notificacion-riesgos-excel', [NotificacionRiesgoExcelTemplateController::class, 'index'])
        ->name('notificacion.excel.index');

    Route::post('/notificacion-riesgos-excel/export', [NotificacionRiesgoExcelTemplateController::class, 'export'])
        ->name('notificacion.excel.export');

    Route::get('/notificacion-riesgos-excel/puestos', [NotificacionRiesgoExcelTemplateController::class, 'puestos'])
        ->name('notificacion.excel.puestos');

    Route::get('/notificacion-riesgos-excel/empleados', [NotificacionRiesgoExcelTemplateController::class, 'empleados'])
        ->name('notificacion.excel.empleados');
    Route::post('/notificacion-riesgos-excel/export/empleado', [NotificacionRiesgoExcelTemplateController::class, 'exportEmpleado'])
        ->name('notificacion.excel.empleado.export');

        Route::post('/prestamo/eliminar', [Prestamos::class, 'eliminarPrestamo'])->name('prestamo.eliminar');

        Route::post('/prestamos/{prestamo}/marcar-pagadas', [PrestamosNuevo::class, 'marcarPagadasHastaFecha'])
    ->name('prestamos.marcar-pagadas');

    Route::post('/prestamos/ajustes/import', [\App\Http\Controllers\AjustesPrestamosController::class, 'importExcel'])
        ->name('prestamos.ajustes.import');

    Route::prefix('mediciones')->group(function () {
        Route::get('/captura', [MedicionesController::class, 'captureBatch'])->name('mediciones.captura');
        //Route::post('/captura', [MedicionesController::class, 'storeBatch'])->name('mediciones.captura.store');
        Route::post('/mediciones/captura/save', [MedicionesController::class, 'saveBatch'])->name('mediciones.captura.save');
        Route::get('/mediciones/captura/prefill', [MedicionesController::class, 'prefill'])->name('mediciones.captura.prefill');
    });

     Route::get('/mediciones/iluminacion/reporte', [MedicionesController::class, 'reporteIluminacion'])
     ->name('mediciones.iluminacion.reporte');

    Route::post('/mediciones/iluminacion/reporte/accion', [MedicionesController::class, 'updateAccionIluminacion'])
        ->name('mediciones.iluminacion.accion');

    Route::get('/mediciones/ruido/reporte', [MedicionesController::class, 'reporteRuido'])
        ->name('mediciones.ruido.reporte');

    Route::post('/mediciones/ruido/reporte/accion', [MedicionesController::class, 'updateAccionRuido'])
        ->name('mediciones.ruido.accion');

    Route::get('/mediciones/ruido/reporte/accion', function () {
        return redirect()->route('mediciones.ruido.reporte');
    });

    Route::get('/mediciones/timeline', [MedicionesController::class, 'timeline'])
    ->name('mediciones.timeline');

Route::get('/matrizriesgos', [RiesgosController::class, 'matriz'])->name('matrizriesgos.index');
Route::get('/riesgos/medida', [RiesgosController::class, 'getMedida'])->name('riesgos.medida.get');
Route::post('/riesgos/medida', [RiesgosController::class, 'saveMedida'])->name('riesgos.medida.save');

Route::get('/identificacion/{id}', [IdentificacionRiesgosController::class, 'fetch'])->name('identificacion.fetch');

     Route::get('/epp/requeridos', [EppRequeridosController::class, 'index'])
    ->name('epp.requeridos');

    Route::post('/mediciones/iluminacion/delete', [MedicionesController::class, 'deleteIluminacion'])
    ->name('mediciones.iluminacion.delete');

Route::post('/mediciones/ruido/delete', [MedicionesController::class, 'deleteRuido'])
    ->name('mediciones.ruido.delete');

    Route::post('/mediciones/iluminacion/update', [MedicionesController::class, 'updateIluminacionRow'])
    ->name('mediciones.iluminacion.update');

Route::post('/mediciones/ruido/update', [MedicionesController::class, 'updateRuidoRow'])
    ->name('mediciones.ruido.update');

    Route::get('/mediciones/graficas', [MedicionesController::class, 'charts'])
     ->name('mediciones.charts');

     Route::get('mediciones/export/iluminacion', [MedicionesController::class, 'exportIluminacion'])
    ->name('mediciones.export.iluminacion');

    Route::get('mediciones/export/ruido', [MedicionesController::class, 'exportRuido'])
    ->name('mediciones.export.ruido');

    Route::get('/prestamos/empleados/export', [\App\Http\Controllers\Prestamos::class, 'exportEmpleadosPrestamo'])
    ->name('empleadosprestamo.export');

    Route::get('mediciones/export/iluminacion-plantilla', [MedicionesController::class, 'exportIluminacionDesdePlantilla'])
    ->name('mediciones.export.iluminacion.plantilla');

    Route::get('/mediciones/timeline/excel', [\App\Http\Controllers\MedicionesController::class, 'timelineExcel'])
    ->name('mediciones.timeline.excel');

    Route::get('nataly/mediciones/export/ruido-plantilla', [MedicionesController::class, 'exportRuidoDesdePlantilla'])
     ->name('mediciones.export.ruido.plantilla');

Route::get('/mediciones/captura/prefill', [MedicionesController::class, 'prefill'])
    ->name('mediciones.captura.prefill');

    Route::get('/prestamos/renuncia', [PrestamoRenunciaController::class, 'form'])
    ->name('prestamos.renuncia.form');

Route::post('/prestamos/renuncia/confirmar', [PrestamoRenunciaController::class, 'confirmar'])
    ->name('prestamos.renuncia.confirmar');

    Route::get('/prestamos/detalle/{id}', [Prestamos::class, 'detallePrestamo'])
    ->whereNumber('id')
    ->name('prestamos.detalle');

    Route::get('/prestamos/resumen-mensual', [PrestamosResumenController::class, 'index'])
    ->name('prestamos.resumen.index');

    Route::post('/prestamos/resumen-mensual/guardar', [PrestamosResumenController::class, 'store'])
        ->name('prestamos.resumen.store');

    Route::post('/prestamos/resumen-mensual/recalcular', [PrestamosResumenController::class, 'recalcular'])
    ->name('prestamos.resumen.recalcular');

    Route::post('/prestamos/resumen-mensual/eliminar', [PrestamosResumenController::class, 'destroy'])
        ->name('prestamos.resumen.eliminar');
        
    Route::get('/prestamos/ajustes', [AjustesPrestamosController::class, 'form'])
        ->name('prestamos.ajustes.form');

    Route::post('/prestamos/ajustes/preview', [AjustesPrestamosController::class, 'previewExcel'])
        ->name('prestamos.ajustes.preview');

    Route::post('/prestamos/ajustes/commit', [AjustesPrestamosController::class, 'commitExcel'])
        ->name('prestamos.ajustes.commit');

Route::get('/epp/detalle', [\App\Http\Controllers\EppRequeridosController::class, 'detalle'])
    ->name('epp.detalle');
    
Route::get('/capacitaciones/detalle', [\App\Http\Controllers\CapacitacionesRequeridasController::class, 'detalle'])
    ->name('cap.detalle');

});



Auth::routes();
