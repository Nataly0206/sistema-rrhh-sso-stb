@extends('layouts.prestamos')

@section('title', 'Control de Prestamos')

@section('content')
    <!-- Breadcrumb -->
    <nav class="flex px-5 py-3 text-gray-700 bg-blue-100 rounded-lg" aria-label="Breadcrumb">
    <ol class="inline-flex items-center space-x-1 md:space-x-3">
        <li class="inline-flex items-center">
        <a href="/" class="inline-flex items-center text-sm font-medium text-black hover:text-blue-900">
            <!-- Ícono de inicio -->
            <svg class="w-4 h-4 mr-2 text-black" fill="currentColor" viewBox="0 0 20 20">
            <path d="M10 2a1 1 0 00-.707.293l-7 7a1 1 0 001.414 1.414L4 10.414V17a1 1 0 001 1h2a1 1 0 001-1v-3h2v3a1 1 0 001 1h2a1 1 0 001-1v-6.586l.293.293a1 1 0 001.414-1.414l-7-7A1 1 0 0010 2z" />
            </svg>
            Inicio
        </a>
        </li>
        <!-- Separador con flechita -->
        <li>
        <div class="flex items-center">
            <svg class="w-4 h-4 mx-2 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clip-rule="evenodd" />
            </svg>
            <span class="text-sm font-medium text-black">Historial de Cuotas</span>
        </div>
        </li>
    </ol>
    </nav>
    <br>

       @if (session('ajustes_msg'))
  <div class="p-3 mb-4 rounded bg-green-50 text-green-800 border border-green-200">
    {{ session('ajustes_msg') }}
  </div>
@endif

@if (session('error'))
  <div class="p-3 mb-4 rounded bg-red-50 text-red-800 border border-red-200">
    {{ session('error') }}
  </div>
@endif

            <form action="{{ route('cuotas') }}" method="GET" class="relative w-full max-w-sm bg-white flex items-center">
        <div class="relative w-full">
            <input
                type="text"
                name="search"
                value="{{ request('search') }}"
                placeholder="Buscar..."
                oninput="this.form.submit()" {{-- aquí está la magia --}}
                class="pl-10 pr-10 py-2 w-full border border-gray-300 rounded-l-lg shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
            />
            <div class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                <svg class="w-5 h-5 text-gray-500" fill="none" stroke="currentColor" stroke-width="2"
                    viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round"
                        d="M21 21l-4.35-4.35M10.5 18a7.5 7.5 0 1 1 0-15 7.5 7.5 0 0 1 0 15z" />
                </svg>
            </div>
        </div>
    </form>
    <!--  <div class="mt-3">
        <a href="{{ route('cuotas.rango') }}" class="inline-flex items-center gap-2 text-sm text-blue-700 hover:text-blue-900">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M8 12h13M8 6h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
            </svg>
            Ver cuotas por rango de fechas
        </a>
    </div>  -->

    @if (session('success'))
        <div class="mt-3 p-3 rounded bg-green-100 text-green-800 text-sm">{{ session('success') }}</div>
    @endif
    @if (session('error'))
        <div class="mt-3 p-3 rounded bg-red-100 text-red-800 text-sm">{{ session('error') }}</div>
    @endif

    @include('prestamos.partials.cuotas-table', ['cuotas' => $cuotas])

@endsection
