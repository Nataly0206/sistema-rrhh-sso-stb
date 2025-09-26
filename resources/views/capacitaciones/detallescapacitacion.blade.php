@extends('layouts.capacitacion')

@section('title', 'Administrar Asistencia a Capacitaciones')

@section('content')

    <nav class="flex px-5 py-3 text-gray-700 bg-blue-100 rounded-lg" aria-label="Breadcrumb">
    <ol class="inline-flex items-center space-x-1 md:space-x-3">
        <li class="inline-flex items-center">
        <a href="/" class="inline-flex items-center text-sm font-medium text-black hover:text-blue-900">
            <svg class="w-4 h-4 mr-2 text-black" fill="currentColor" viewBox="0 0 20 20">
            <path d="M10 2a1 1 0 00-.707.293l-7 7a1 1 0 001.414 1.414L4 10.414V17a1 1 0 001 1h2a1 1 0 001-1v-3h2v3a1 1 0 001 1h2a1 1 0 001-1v-6.586l.293.293a1 1 0 001.414-1.414l-7-7A1 1 0 0010 2z" />
            </svg>
            Inicio
        </a>
        </li>
        <li class="inline-flex items-center">
        <a href="/Capacitaciones" class="inline-flex items-center text-sm font-medium text-black hover:text-blue-900">
            <svg class="w-4 h-4 mx-2 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clip-rule="evenodd" />
            </svg>
            Registro de Asistencia a Capacitaciones
        </a>
        </li>
        <li>
        <div class="flex items-center">
            <svg class="w-4 h-4 mx-2 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clip-rule="evenodd" />
            </svg>
            <span class="text-sm font-medium text-black">Detalles de Capacitaciones</span>
        </div>
        </li>
    </ol>
    </nav>
    <br>    

    <div class="flex justify-between items-center w-full mb-4">
        <!-- Botones a la izquierda -->
        <div class="flex items-center gap-2">
            <!-- Botón agregar -->
            <div class="inline-flex rounded-md shadow-xs" role="group">
                <button data-modal-target="create-modal" data-modal-toggle="create-modal" type="button"
                    class="inline-flex items-center px-4 py-2 text-sm font-medium text-gray-900 bg-transparent border border-gray-900 rounded-s-lg hover:bg-gray-900 focus:z-10 focus:ring-2 focus:ring-gray-500 focus:bg-gray-900">
                    <svg class="w-6 h-6 mr-1" fill="none" viewBox="0 0 24 24">
                        <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                            d="M5 12h14m-7 7V5" />
                    </svg>
                    Agregar
                </button>
                    <form id="importForm" action="{{ route('detallescapacitacion.import') }}" method="POST" enctype="multipart/form-data" style="display: none;">
                        @csrf
                        <input type="file" id="excelFileInput" name="excel_file" accept=".xls,.xlsx,.xlsm" onchange="document.getElementById('importForm').submit();">
                    </form>
                    <button type="button" 
                        onclick="document.getElementById('excelFileInput').click();" 
                        class="inline-flex items-center px-4 py-2 text-sm font-medium text-gray-900 bg-transparent border border-gray-900 rounded-s-lg hover:bg-gray-900 focus:z-10 focus:ring-2 focus:ring-gray-500 focus:bg-gray-900">
                        <svg class="w-6 h-6 text-gray-800 dark:text-white" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24">
                            <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 12V7.914a1 1 0 0 1 .293-.707l3.914-3.914A1 1 0 0 1 9.914 3H18a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1v-4m5-13v4a1 1 0 0 1-1 1H5m0 6h9m0 0-2-2m2 2-2 2"/>
                        </svg>
                        Subir Lista de Capacitaciones
                    </button>
            </div>
            <!-- Main modal -->
            <div id="create-modal" tabindex="-1" aria-hidden="true" class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
                <div class="relative p-4 w-full max-w-md max-h-full">
                    <!-- Modal content -->
                    <div class="relative bg-white rounded-lg shadow-sm">
                        <!-- Modal header -->
                        <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
                            <h3 class="text-lg font-semibold text-gray-900">
                                Ingresar Nombre de Capacitación
                            </h3>
                            <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-toggle="create-modal">
                            <svg class="w-3 h-3" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 14">
                                <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"/>
                            </svg>
                            <span class="sr-only">Close modal</span>
                            </button>
                        </div>
                        <!-- Modal body -->
                        <form action="{{ route('detallescapacitacion.storecapinstructor') }}" method="POST" class="p-4 md:p-5">
                            @csrf 
                            <div class="grid gap-4 mb-4 grid-cols-2">
                                <div class="col-span-2">
                                    <label for="id_capacitacion" class="block mb-2 text-sm font-medium text-gray-900">Capacitación</label>
                                    <select name="id_capacitacion" id="id_capacitacion" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block w-full p-2.5" required>
                                        <option value="" disabled selected>Seleccione una capacitación</option>
                                        @foreach($capacitaciones as $capacitacion)
                                            <option value="{{ $capacitacion->id_capacitacion }}">{{ $capacitacion->capacitacion }}</option>
                                        @endforeach
                                    </select>
                                </div>
                                <div class="col-span-2">
                                    <label for="id_instructor" class="block mb-2 text-sm font-medium text-gray-900">Instructores</label>
                                    <select name="id_instructor" id="id_instructor" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block w-full p-2.5" required>
                                        <option value="" disabled selected>Seleccione un instructor</option>
                                            @foreach($instructores as $instructor)
                                                <option value="{{ $instructor->id_instructor }}">{{ $instructor->instructor }}</option>
                                            @endforeach
                                    </select>
                                </div>
                                    <div class="col-span-2">
                                        <label for="duracion" class="block mb-2 text-sm font-medium text-gray-900">
                                            Duración de la Capacitación
                                        </label>
                                        <p class="mb-2 text-sm text-gray-600">
                                            Ingrese las horas en minutos
                                        </p>
                                        <div class="flex items-center gap-2">
                                            <input type="text" name="duracion" id="duracion" 
                                                class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block p-2.5 w-full"
                                                placeholder="Ejem: 1 hora y media = 90, 16 horas = 960" required="">
                                            <span class="text-gray-700">minutos</span>
                                        </div>
                                    </div>

                                    <div class="col-span-2">
                                        <label for="num_categoria" class="block mb-2 text-sm font-medium text-gray-900">Número de categoría</label>
                                        <input type="text" name="num_categoria" id="num_categoria" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block p-2.5 w-full" placeholder="Ej: 01">
                                    </div>

                                    <div class="col-span-2">
                                        <div class="flex items-center gap-2">
                                            <input id="programada" name="programada" type="checkbox" value="SI" class="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500">
                                            <label for="programada" class="ml-2 text-sm font-medium text-gray-900">Programada</label>
                                        </div>
                                        <p class="mt-1 text-xs text-gray-500">Marca la casilla si la capacitación está programada.</p>
                                    </div>
                                </div>
                            <button type="submit" class="text-white inline-flex items-center bg-blue-700 hover:bg-blue-800 focus:ring-4 focus:outline-none focus:ring-blue-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center">
                            <svg class="me-1 -ms-1 w-5 h-5" fill="currentColor" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><path fill-rule="evenodd" d="M10 5a1 1 0 011 1v3h3a1 1 0 110 2h-3v3a1 1 0 11-2 0v-3H6a1 1 0 110-2h3V6a1 1 0 011-1z" clip-rule="evenodd"></path></svg>
                                Guardar
                            </button>
                        </form>
                    </div>
                </div>
            </div> 

            <!-- Botones de navegación -->
            <div class="inline-flex rounded-md shadow-xs" role="group">
                <a href="/capacitacion" type="button"
                    class="inline-flex items-center px-4 py-2 text-sm font-medium text-white bg-blue-600 border border-blue-700 rounded-l-lg hover:bg-blue-700 focus:z-10 focus:ring-2 focus:ring-blue-500">
                    <svg class="w-6 h-6 text-white" xmlns="http://www.w3.org/2000/svg" fill="currentColor"
                        viewBox="0 0 24 24">
                        <path fill-rule="evenodd"
                            d="M8 3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1h2a2 2 0 0 1 2 2v15a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h2Zm6 1h-4v2H9a1 1 0 0 0 0 2h6a1 1 0 1 0 0-2h-1V4Zm-3 8a1 1 0 0 1 1-1h3a1 1 0 1 1 0 2h-3a1 1 0 0 1-1-1Zm-2-1a1 1 0 1 0 0 2h.01a1 1 0 1 0 0-2H9Zm2 5a1 1 0 0 1 1-1h3a1 1 0 1 1 0 2h-3a1 1 0 0 1-1-1Zm-2-1a1 1 0 1 0 0 2h.01a1 1 0 1 0 0-2H9Z"
                            clip-rule="evenodd" />
                    </svg>
                    Ver Capacitaciones
                </a>
                <a href="/instructor" type="button"
                    class="inline-flex items-center px-4 py-2 text-sm font-medium text-white bg-blue-600 border border-blue-700 rounded-r-lg hover:bg-blue-700 focus:z-10 focus:ring-2 focus:ring-blue-500">
                    <svg class="w-6 h-6 text-white" xmlns="http://www.w3.org/2000/svg" fill="currentColor"
                        viewBox="0 0 24 24">
                        <path fill-rule="evenodd"
                            d="M8 3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1h2a2 2 0 0 1 2 2v15a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h2Zm6 1h-4v2H9a1 1 0 0 0 0 2h6a1 1 0 1 0 0-2h-1V4Zm-3 8a1 1 0 0 1 1-1h3a1 1 0 1 1 0 2h-3a1 1 0 0 1-1-1Zm-2-1a1 1 0 1 0 0 2h.01a1 1 0 1 0 0-2H9Zm2 5a1 1 0 0 1 1-1h3a1 1 0 1 1 0 2h-3a1 1 0 0 1-1-1Zm-2-1a1 1 0 1 0 0 2h.01a1 1 0 1 0 0-2H9Z"
                            clip-rule="evenodd" />
                    </svg>
                    Ver Instructores
                </a>
            </div>
        </div>

        <!-- Buscador a la derecha -->
        <form action="{{ route('detallescapacitacion') }}" method="GET"
            class="relative w-full max-w-sm bg-white flex items-center">
            <div class="relative w-full">
                <input type="text" name="search" value="{{ request('search') }}" placeholder="Buscar..."
                    oninput="this.form.submit()"
                    class="pl-10 pr-10 py-2 w-full border border-gray-300 rounded-lg shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm" />
                <div class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <svg class="w-5 h-5 text-gray-500" fill="none" stroke="currentColor" stroke-width="2"
                        viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round"
                            d="M21 21l-4.35-4.35M10.5 18a7.5 7.5 0 1 1 0-15 7.5 7.5 0 0 1 0 15z" />
                    </svg>
                </div>
            </div>
        </form>
    </div>

    <div class="relative overflow-x-auto overflow-y-auto max-h-[70vh] shadow-md sm:rounded-lg" style="margin-top: 20px;">
        <table class="w-full text-sm text-left text-gray-500 border-separate border-spacing-0">
            <thead class="sticky top-0 z-20 bg-gray-50">
                <tr>
                    <th scope="col" class="px-6 py-3">
                        Nombre de la Capacitación
                    </th>
                    <th scope="col" class="px-6 py-3">
                        Nombre del Instructor
                    </th>
                    <th scope="col" class="px-6 py-3">
                        Duración de la Capacitación (Minutos)
                    </th>
                    <th scope="col" class="px-6 py-3">
                        Número de categoría
                    </th>
                    <th scope="col" class="px-6 py-3">
                        Programada
                    </th>
                    <th scope="col" class="px-6 py-3">
                        Acción
                    </th>
                </tr>
            </thead>
            <tbody>
                @foreach ($capinstructor as $capinstructores)
                <tr class="bg-white border-b border-gray-200 hover:bg-gray-50">
                    <td class="px-6 py-4">
                        {{ $capinstructores->capacitacion }}
                    </td>
                    <td class="px-6 py-4">
                        {{ $capinstructores->instructor }}
                    </td>
                    <td class="px-6 py-4">
                        {{ $capinstructores->duracion }}
                    </td>
                    <td class="px-6 py-4">
                        {{ $capinstructores->num_categoria }}
                    </td>
                    <td class="px-6 py-4">
                        {{ $capinstructores->programada === 'SI' ? 'SI' : 'NO' }}
                    </td>
                    <td class="flex items-center px-6 py-4">
                        <button data-modal-target="edit-modal-{{ $capinstructores->id_capacitacion_instructor }}" data-modal-toggle="edit-modal-{{ $capinstructores->id_capacitacion_instructor }}" class="font-medium text-blue-600 hover:underline">Editar</button>

                        <!-- Main modal -->
                        <div id="edit-modal-{{ $capinstructores->id_capacitacion_instructor }}" tabindex="-1" aria-hidden="true" class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
                            <div class="relative p-4 w-full max-w-md max-h-full">
                                <!-- Modal content -->
                                <div class="relative bg-white rounded-lg shadow-sm">
                                    <!-- Modal header -->
                                    <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
                                        <h3 class="text-lg font-semibold text-gray-900">
                                            Editar Detalles de Capacitación
                                        </h3>
                                        <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-toggle="edit-modal-{{ $capinstructores->id_capacitacion_instructor }}">
                                            <svg class="w-3 h-3" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 14">
                                                <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"/>
                                            </svg>
                                            <span class="sr-only">Close modal</span>
                                        </button>
                                    </div>
                                    <!-- Modal body -->
                                    <form action="{{ route('detallescapacitacion.updatecapinstructor', $capinstructores->id_capacitacion_instructor) }}" method="POST" class="p-4 md:p-5">
                                        @csrf
                                        @method('PUT')
                                        <div class="grid gap-4 mb-4 grid-cols-2">
                                            <div class="col-span-2">
                                                <label for="id_capacitacion" class="block mb-2 text-sm font-medium text-gray-900">Capacitación</label>
                                                <select name="id_capacitacion" id="id_capacitacion" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block w-full p-2.5" required>
                                                    @foreach($capacitaciones as $capacitacion)
                                                        <option value="{{ $capacitacion->id_capacitacion }}" @if($capinstructores->id_capacitacion == $capacitacion->id_capacitacion) selected @endif>
                                                            {{ $capacitacion->capacitacion }}
                                                        </option>
                                                    @endforeach
                                                </select>
                                            </div>
                                            <div class="col-span-2">
                                                <label for="id_instructor" class="block mb-2 text-sm font-medium text-gray-900">Instructores</label>
                                                <select name="id_instructor" id="id_instructor" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block w-full p-2.5" required>
                                                    @foreach($instructores as $instructor)
                                                        <option value="{{ $instructor->id_instructor}}" @if($capinstructores->id_instructor == $instructor->id_instructor) selected @endif>
                                                            {{ $instructor->instructor }}
                                                        </option>
                                                    @endforeach
                                                </select>
                                            </div>
                                            <div class="col-span-2">
                                                <label for="duracion" class="block mb-2 text-sm font-medium text-gray-900">
                                                    Duración de la Capacitación
                                                </label>
                                                <p class="mb-2 text-sm text-gray-600">
                                                    Ingrese las horas en minutos
                                                </p>
                                                <div class="flex items-center gap-2">
                                                    <input type="text" name="duracion" id="duracion" value="{{ $capinstructores->duracion }}" 
                                                        class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block p-2.5 w-full"
                                                        placeholder="Ejem: 1 hora y media = 90, 16 horas = 960">
                                                    <span class="text-gray-700">minutos</span>
                                                </div>
                                            </div>

                                            <div class="col-span-2">
                                                <label for="num_categoria-{{ $capinstructores->id_capacitacion_instructor }}" class="block mb-2 text-sm font-medium text-gray-900">Número de categoría</label>
                                                <input type="text" name="num_categoria" id="num_categoria-{{ $capinstructores->id_capacitacion_instructor }}" value="{{ $capinstructores->num_categoria }}" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-primary-600 focus:border-primary-600 block p-2.5 w-full" placeholder="Ej: 01">
                                            </div>

                                            <div class="col-span-2">
                                                <div class="flex items-center gap-2">
                                                    <input id="programada-{{ $capinstructores->id_capacitacion_instructor }}" name="programada" type="checkbox" value="SI" @if($capinstructores->programada === 'SI') checked @endif class="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500">
                                                    <label for="programada-{{ $capinstructores->id_capacitacion_instructor }}" class="ml-2 text-sm font-medium text-gray-900">Programada</label>
                                                </div>
                                                <p class="mt-1 text-xs text-gray-500">Marca la casilla si la capacitación está programada.</p>
                                            </div>
                                        </div>
                                        <button type="submit" class="text-white inline-flex items-center bg-blue-700 hover:bg-blue-800 focus:ring-4 focus:outline-none focus:ring-blue-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center">
                                            <svg class="me-1 -ms-1 w-5 h-5" fill="currentColor" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><path fill-rule="evenodd" d="M10 5a1 1 0 011 1v3h3a1 1 0 110 2h-3v3a1 1 0 11-2 0v-3H6a1 1 0 110-2h3V6a1 1 0 011-1z" clip-rule="evenodd"></path></svg>
                                            Guardar
                                        </button>
                                    </form>
                                </div>
                            </div>
                        </div> 

                        <button data-modal-target="popup-modal-{{ $capinstructores->id_capacitacion_instructor }}" data-modal-toggle="popup-modal-{{ $capinstructores->id_capacitacion_instructor }}" class="font-medium text-red-600 hover:underline ms-3">Eliminar</button>

                        <div id="popup-modal-{{ $capinstructores->id_capacitacion_instructor }}" tabindex="-1" class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
                            <div class="relative p-4 w-full max-w-md max-h-full">
                                <div class="relative bg-white rounded-lg shadow-sm dark:bg-gray-700">
                                    <button type="button" class="absolute top-3 end-2.5 text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center dark:hover:bg-gray-600 dark:hover:text-white" data-modal-hide="popup-modal-{{ $capinstructores->id_capacitacion_instructor }}">
                                        <svg class="w-3 h-3" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 14">
                                            <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"/>
                                        </svg>
                                        <span class="sr-only">Close modal</span>
                                    </button>
                                    <div class="p-4 md:p-5 text-center">
                                        <svg class="mx-auto mb-4 text-gray-400 w-12 h-12 dark:text-gray-200" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 20 20">
                                            <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 11V6m0 8h.01M19 10a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z"/>
                                        </svg>
                                        <h3 class="mb-5 text-lg font-normal text-gray-500 dark:text-gray-400">¿Está seguro de eliminar este item?</h3>
                                        <form action="{{ route('detallescapacitacion.destroycapinstructor', $capinstructores->id_capacitacion_instructor) }}" method="POST">
                                            @csrf
                                            @method('DELETE')
                                            <button type="submit" class="text-white bg-red-600 hover:bg-red-800 focus:ring-4 focus:outline-none focus:ring-red-300 dark:focus:ring-red-800 font-medium rounded-lg text-sm inline-flex items-center px-5 py-2.5 text-center">
                                                Sí, estoy seguro
                                            </button>
                                            <button data-modal-hide="popup-modal-{{ $capinstructores->id_capacitacion_instructor }}" type="button" class="py-2.5 px-5 ms-3 text-sm font-medium text-gray-900 focus:outline-none bg-white rounded-lg border border-gray-200 hover:bg-gray-100 hover:text-blue-700 focus:z-10 focus:ring-4 focus:ring-gray-100 dark:focus:ring-gray-700 dark:bg-gray-800 dark:text-gray-400 dark:border-gray-600 dark:hover:text-white dark:hover:bg-gray-700">
                                                No, cancelar
                                            </button>
                                        </form>
                                    </div>
                                </div>
                            </div>
                        </div>

                    </td>
                </tr>
                @endforeach
            </tbody>
        </table>
        {{ $capinstructor->links() }}
    </div>

    <!-- jQuery -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

@endsection
