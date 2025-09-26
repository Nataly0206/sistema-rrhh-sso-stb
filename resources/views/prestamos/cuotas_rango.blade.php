@extends('layouts.prestamos')

@section('title', 'Cuotas por rango')

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
            <li>
                <div class="flex items-center">
                    <svg class="w-4 h-4 mx-2 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clip-rule="evenodd" />
                    </svg>
                    <span class="text-sm font-medium text-black">Cuotas por rango</span>
                </div>
            </li>
        </ol>
    </nav>

    <!-- <div class="mt-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <h2 class="text-xl font-semibold text-gray-800">Filtrar cuotas por fechas</h2>
        <a href="{{ route('cuotas') }}" class="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-blue-700 border border-blue-300 rounded hover:bg-blue-50">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15 19l-7-7 7-7" />
            </svg>
            Volver a historial completo
        </a>
    </div> -->

    @if ($errorMensaje)
        <div class="mt-4 p-3 rounded bg-red-50 text-red-800 border border-red-200">
            {{ $errorMensaje }}
        </div>
    @endif

    @if (session('success'))
        <div class="mt-4 p-3 rounded bg-green-100 text-green-800 border border-green-200 text-sm">
            {{ session('success') }}
        </div>
    @endif

    <form action="{{ route('prestamos.ajustes.preview') }}" method="post" enctype="multipart/form-data" class="mt-4 space-y-3">
        @csrf
        <input type="file" name="archivo" accept=".xlsx,.xls" required class="block">
        <button class="px-4 py-2 bg-blue-600 text-white rounded">Previsualizar ajustes</button>
    </form>


    <form action="{{ route('cuotas.rango') }}" method="GET" class="mt-4 grid gap-4 md:grid-cols-4">
        <div>
            <label for="fecha_inicio" class="block mb-1 text-sm font-medium text-gray-700">Fecha inicio</label>
            <input type="date" id="fecha_inicio" name="fecha_inicio" value="{{ request('fecha_inicio') }}" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
        </div>
        <div>
            <label for="fecha_fin" class="block mb-1 text-sm font-medium text-gray-700">Fecha fin</label>
            <input type="date" id="fecha_fin" name="fecha_fin" value="{{ request('fecha_fin') }}" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
        </div>
        <div>
            <label for="estado" class="block mb-1 text-sm font-medium text-gray-700">Estado</label>
            <select id="estado" name="estado" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
                <option value="todas" {{ request('estado', 'todas') === 'todas' ? 'selected' : '' }}>Pagadas y pendientes</option>
                <option value="pagadas" {{ request('estado') === 'pagadas' ? 'selected' : '' }}>Solo pagadas</option>
                <option value="pendientes" {{ request('estado') === 'pendientes' ? 'selected' : '' }}>Solo pendientes</option>
            </select>
        </div>
        <div>
            <label for="search" class="block mb-1 text-sm font-medium text-gray-700">Buscar</label>
            <input type="text" id="search" name="search" value="{{ request('search') }}" placeholder="Nombre, c�digo o pr�stamo" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
        </div>
        <div class="md:col-span-4 flex flex-wrap items-center gap-2">
            <button type="submit" class="inline-flex items-center px-4 py-2 text-sm font-semibold text-white bg-blue-600 rounded hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-400 focus:ring-offset-1">
                Aplicar filtros
            </button>
            <a href="{{ route('cuotas.rango') }}" class="text-sm text-gray-600 hover:text-gray-900">Limpiar filtros</a>
        </div>
    </form>

    @if ($resumen)
        <div class="mt-6 grid gap-3 sm:grid-cols-3">
            <div class="p-4 bg-white border border-gray-200 rounded shadow-sm">
                <p class="text-xs uppercase text-gray-500">Total de cuotas</p>
                <p class="text-lg font-semibold text-gray-800">{{ number_format($resumen['total']) }}</p>
            </div>
            <div class="p-4 bg-white border border-gray-200 rounded shadow-sm">
                <p class="text-xs uppercase text-gray-500">Pagadas</p>
                <p class="text-lg font-semibold text-emerald-600">{{ number_format($resumen['pagadas']) }}</p>
            </div>
            <div class="p-4 bg-white border border-gray-200 rounded shadow-sm">
                <p class="text-xs uppercase text-gray-500">Pendientes</p>
                <p class="text-lg font-semibold text-amber-600">{{ number_format($resumen['pendientes']) }}</p>
            </div>
        </div>
    @endif

    @if ($ajustePlan)
        @include('prestamos.partials.ajustes-preview', [
            'ajustePlan' => $ajustePlan,
            'ajusteToken' => $ajusteToken,
            'filtros' => $filtros ?? [],
        ])
    @endif

    @if ($cuotas)
        @include('prestamos.partials.cuotas-table', ['cuotas' => $cuotas])
    @elseif (!$errorMensaje && (request()->filled('fecha_inicio') || request()->filled('fecha_fin')))
        <div class="mt-6 p-3 rounded bg-yellow-50 text-yellow-800 border border-yellow-200 text-sm">
            Selecciona ambas fechas y vuelve a intentar para ver las cuotas del rango.
        </div>
    @else
        <div class="mt-6 p-3 rounded bg-blue-50 text-blue-800 border border-blue-200 text-sm">
            Elige un rango de fechas y presiona "Aplicar filtros" para listar las cuotas.
        </div>
    @endif
@endsection
