<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>@yield('title', 'Análisis de Riesgos por Puesto')</title>

  <!-- Tailwind config: define brand BEFORE tailwind -->
  <script>
    window.tailwind = window.tailwind || {};
    tailwind.config = {
      theme: { extend: { colors: { brand: '#00B0F0' } } }
    }
  </script>
  <!-- Tailwind CSS (CDN) -->
  <script src="https://cdn.tailwindcss.com"></script>

  <!-- Flowbite (JS) -->
  <script src="https://unpkg.com/flowbite@2.3.0/dist/flowbite.min.js" defer></script>

  <!-- Bootstrap (opcional si lo usas en formularios/tablas) -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">

  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.bunny.net">
  <link href="https://fonts.bunny.net/css?family=instrument-sans:400,500,600" rel="stylesheet" />

  <style>
    html, body { height: 100%; }
    body {
      font-family: 'Instrument Sans', ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'Apple Color Emoji', 'Segoe UI Emoji';
    }
    .scrollbar-thin::-webkit-scrollbar{width:8px}
    .scrollbar-thin::-webkit-scrollbar-thumb{background:#c7c7d1;border-radius:8px}

    /* Transiciones suaves */
    #app-sidebar { transition: transform 200ms ease; }
    .app-main { transition: margin 200ms ease; }

    /* Ocultar sidebar en escritorio (sm+) cuando html tiene .sidebar-hidden */
    @media (min-width: 640px) {
      .sidebar-hidden #app-sidebar { transform: translateX(-100%) !important; }
      .sidebar-hidden .app-main { margin-left: 0 !important; }
    }
  </style>
</head>

<body class="bg-slate-50 text-slate-800">
  <!-- Topbar -->
  <header class="fixed top-0 z-50 w-full bg-white/90 backdrop-blur border-b border-slate-200">
    <div class="max-w-screen-xl mx-auto flex items-center justify-between px-4 py-3">
      <div class="flex items-center gap-2">
        <!-- Hamburger (mobile) -->
        <button
          class="sm:hidden inline-flex items-center justify-center h-10 w-10 rounded-lg border border-slate-300 hover:bg-slate-100"
          type="button"
          data-drawer-target="app-sidebar"
          data-drawer-toggle="app-sidebar"
          aria-controls="app-sidebar"
          aria-label="Abrir menú">
          <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5"/>
          </svg>
        </button>

        <a href="/" class="flex items-center gap-3">
          <img src="{{ asset('img/logo.PNG') }}" alt="Logo" class="h-10 w-auto">
          <div class="leading-tight">
            <p class="text-sm text-slate-500">Sistema de Recursos Humanos</p>
            <p class="text-lg font-semibold">Análisis de Riesgos por Puesto de Trabajo</p>
          </div>
        </a>
      </div>

      <div class="flex items-center gap-2">
        <!-- Toggle ocultar/mostrar sidebar (escritorio) -->
        <button id="sidebarToggle"
                class="hidden sm:inline-flex items-center justify-center h-10 w-10 rounded-lg border border-slate-300 hover:bg-slate-100"
                type="button" aria-label="Ocultar/Mostrar menú lateral" title="Ocultar/Mostrar menú">
          <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6 pointer-events-none" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M3.75 5.25v13.5m4.5-13.5h12v13.5h-12m4.5-6.75H8.25"/>
          </svg>
        </button>

        <button
          data-modal-target="select-modal"
          data-modal-toggle="select-modal"
          class="hidden sm:inline-block text-white bg-brand hover:brightness-95 focus:ring-4 focus:outline-none focus:ring-brand/30 font-medium rounded-lg text-sm px-4 py-2">
          Cambiar de Módulo
        </button>
      </div>
    </div>
  </header>

  <!-- Sidebar (drawer en móvil, fijo en escritorio) -->
  <aside id="app-sidebar"
         class="fixed top-0 left-0 z-40 w-72 h-screen pt-20 transition-transform -translate-x-full sm:translate-x-0 bg-white border-r border-slate-200">
    <div class="h-full flex flex-col">
      <!-- Encabezado del menú -->
      <div class="px-4 py-3 border-b border-slate-200 bg-gradient-to-r from-brand/10 to-transparent">
        <span class="inline-flex items-center gap-2 text-sm font-medium text-slate-600">
          <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5 text-brand" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M3 7h18M5 7v10a2 2 0 002 2h10a2 2 0 002-2V7"/>
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 11h6m-6 4h6"/>
          </svg>
          Riesgos · Menú
        </span>
      </div>

      <!-- NAV -->
      <nav class="flex-1 overflow-y-auto scrollbar-thin px-2 py-3">
        @php
          $activeLink = function($pattern){
            return request()->is($pattern)
              ? 'bg-brand/10 text-brand border border-brand/30'
              : 'text-slate-700 hover:bg-slate-100 border border-transparent';
          };
        @endphp
        <ul class="space-y-1">

        <li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">PASO 1</li>
          <!-- Preparación (collapsible) -->
          <li>
            <button type="button"
                    class="w-full flex items-center justify-between rounded-xl px-3 py-2 text-slate-700 hover:bg-slate-100"
                    data-collapse-toggle="grp-preparacion">
              <span class="inline-flex items-center gap-3">
                <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 7h16M8 7v10m8-10v10M6 17h12"/>
                </svg>
                <span class="text-sm font-medium">Preparación</span>
              </span>
              <svg class="h-4 w-4 transition-transform" data-chevron viewBox="0 0 20 20" fill="currentColor">
                <path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.71a.75.75 0 011.08 1.04l-4.25 4.25a.75.75 0 01-1.06 0L5.21 8.27a.75.75 0 01.02-1.06z" clip-rule="evenodd"/>
              </svg>
            </button>
            <div id="grp-preparacion" class="mt-1 ml-2 hidden space-y-1">
              <a href="/puestos" class="flex items-center gap-3 rounded-lg px-3 py-2 {{ $activeLink('puestos*') }}">
                <span class="text-sm">Puestos de Trabajo</span>
              </a>
              <a href="/quimicos" class="flex items-center gap-3 rounded-lg px-3 py-2">
                <span class="text-sm">Inventario de Químicos</span>
              </a>
              <a href="/medidasriesgo" class="flex items-center gap-3 rounded-lg px-3 py-2 {{ $activeLink('medidasriesgo*') }}">
                <span class="text-sm">Medidas de Prevención por Riesgo</span>
              </a>
            </div>
          </li>
    <li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">PASO 2</li>
          <!-- Identificación de Riesgos (collapsible) -->
          <li>
            <button type="button"
                    class="w-full flex items-center justify-between rounded-xl px-3 py-2 text-slate-700 hover:bg-slate-100"
                    data-collapse-toggle="grp-identificacion">
              <span class="inline-flex items-center gap-3">
                <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M7 4h10M7 8h10M7 12h10M7 16h6"/>
                </svg>
                <span class="text-sm font-medium">Identificación de Riesgos</span>
              </span>
              <svg class="h-4 w-4 transition-transform" data-chevron viewBox="0 0 20 20" fill="currentColor">
                <path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.71a.75.75 0 011.08 1.04l-4.25 4.25a.75.75 0 01-1.06 0L5.21 8.27a.75.75 0 01.02-1.06z" clip-rule="evenodd"/>
              </svg>
            </button>
            <div id="grp-identificacion" class="mt-1 ml-2 hidden space-y-1">
              <a href="#"
                 data-modal-target="modal-identificacion-riesgos"
                 data-modal-toggle="modal-identificacion-riesgos"
                 class="flex items-center gap-3 rounded-lg px-3 py-2 text-slate-700 hover:bg-slate-100">
                <span class="text-sm">Formato de Identificación (Excel)</span>
              </a>
              <a href="/identificacion-riesgos" class="flex items-center gap-3 rounded-lg px-3 py-2 {{ $activeLink('identificacion-riesgos*') }}">
                <span class="text-sm">Hoja de Campo</span>
              </a>
            </div>
          </li>
<li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">PASO 3</li>
          <!-- Evaluación (enlace abre modal) -->
          <li>
            <a href="#"
               data-modal-target="modal-evaluacion-riesgos"
               data-modal-toggle="modal-evaluacion-riesgos"
               class="group flex items-center gap-3 rounded-xl px-3 py-2 text-slate-700 hover:bg-slate-100 border border-transparent">
              <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 4h6M9 8h6M7 4h10a2 2 0 012 2v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6a2 2 0 012-2z"/>
              </svg>
              <span class="text-sm font-medium">Evaluación de Riesgos (Excel)</span>
            </a>
          </li>

          <li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">PASO 4</li>
          <li>
  @php
    $planActive = request()->routeIs('riesgos.verificacion.plan_accion');
  @endphp
  <a href="{{ route('riesgos.verificacion.plan_accion') }}"
     class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $planActive ? 'bg-brand/10 text-brand border border-brand/30' : 'text-slate-700 hover:bg-slate-100 border border-transparent' }}">
    <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5"
            d="M3 7h18M5 7v10a2 2 0 002 2h10a2 2 0 002-2V7M9 12h6m-6 4h4" />
    </svg>
    <span class="text-sm font-medium">Plan de Acción (Control de Riesgos)</span>
  </a>
</li>

<li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">PASO 5</li>
          <!-- Matrices/Resumen (collapsible) -->
          <li>
            <button type="button"
                    class="w-full flex items-center justify-between rounded-xl px-3 py-2 text-slate-700 hover:bg-slate-100"
                    data-collapse-toggle="grp-matrices">
              <span class="inline-flex items-center gap-3">
                <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 6h7v7H4zM13 6h7v4h-7zM13 12h7v6h-7zM4 15h7v3H4z"/>
                </svg>
                <span class="text-sm font-medium">Matrices / Resumen</span>
              </span>
              <svg class="h-4 w-4 transition-transform" data-chevron viewBox="0 0 20 20" fill="currentColor">
                <path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.71a.75.75 0 011.08 1.04l-4.25 4.25a.75.75 0 01-1.06 0L5.21 8.27a.75.75 0 01.02-1.06z" clip-rule="evenodd"/>
              </svg>
            </button>
            <div id="grp-matrices" class="mt-1 ml-2 hidden space-y-1">
              <a href="/matrizriesgos" class="flex items-center gap-3 rounded-lg px-3 py-2 {{ $activeLink('matrizriesgos*') }}">
                <span class="text-sm">Matriz de Identificación</span>
              </a>
              <a href="/evaluacion" class="flex items-center gap-3 rounded-lg px-3 py-2 {{ $activeLink('evaluacion*') }}">
                <span class="text-sm">Matriz de Evaluación</span>
              </a>
              <a href="/matriz-quimicos" class="flex items-center gap-3 rounded-lg px-3 py-2 {{ $activeLink('matriz-quimicos*') }}">
                <span class="text-sm">Matriz de Riesgo Químico</span>
              </a>
            </div>
          </li>
<li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">PASO 6</li>
          <!-- Notificación de Riesgos (collapsible) -->
          <li>
            <button type="button"
                    class="w-full flex items-center justify-between rounded-xl px-3 py-2 text-slate-700 hover:bg-slate-100"
                    data-collapse-toggle="grp-notificacion">
              <span class="inline-flex items-center gap-3">
                <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M15 17h5l-1.4-1.4A2 2 0 0118 14.2V11a6 6 0 10-12 0v3.2a2 2 0 01-.6 1.4L4 17h5m2 0v1a3 3 0 106 0v-1"/>
                </svg>
                <span class="text-sm font-medium">Notificación de Riesgos</span>
              </span>
              <svg class="h-4 w-4 transition-transform" data-chevron viewBox="0 0 20 20" fill="currentColor">
                <path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.71a.75.75 0 011.08 1.04l-4.25 4.25a.75.75 0 01-1.06 0L5.21 8.27a.75.75 0 01.02-1.06z" clip-rule="evenodd"/>
              </svg>
            </button>
            <div id="grp-notificacion" class="mt-1 ml-2 hidden space-y-1">
              <a href="#"
                 data-modal-target="modal-notificacion-empleado"
                 data-modal-toggle="modal-notificacion-empleado"
                 class="flex items-center gap-3 rounded-lg px-3 py-2 text-slate-700 hover:bg-slate-100">
                <span class="text-sm">Formato por Empleado (Excel)</span>
              </a>
              <a href="#"
                 data-modal-target="modal-notificacion-puesto"
                 data-modal-toggle="modal-notificacion-puesto"
                 class="flex items-center gap-3 rounded-lg px-3 py-2 text-slate-700 hover:bg-slate-100">
                <span class="text-sm">Formato por Puesto (Excel)</span>
              </a>
            </div>
          </li>
<li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">Consultar Información</li>
                                        <!-- Ver Informacion -->
          <li>
            <a href="/verificacion" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('verificacion*') }}">
              <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M21 21l-4.35-4.35M10.5 18a7.5 7.5 0 100-15 7.5 7.5 0 000 15z"/>
              </svg>
              <span class="text-sm font-medium">Ver Informacion</span>
            </a>
          </li>
          <li>
            <a href="{{ route('riesgos.analisis') }}" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('analisis-riesgos*') }}">
              <svg class="h-5 w-5" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 19h16M7 11v8m5-14v14m5-6v6M6 5h12"/>
              </svg>
              <span class="text-sm font-medium">Analisis de Riesgos</span>
            </a>
          </li>
        </ul>
      </nav>

      <!-- Botón Cambiar de Módulo (visible también en mobile) -->
      <div class="p-3 border-t border-slate-200">
        <button
          data-modal-target="select-modal"
          data-modal-toggle="select-modal"
          class="w-full inline-flex items-center justify-center gap-2 text-white bg-brand hover:brightness-95 focus:ring-4 focus:outline-none focus:ring-brand/30 font-medium rounded-lg text-sm px-4 py-2">
          <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M10 19l-7-7 7-7M14 5l7 7-7 7"/>
          </svg>
          Cambiar de Módulo
        </button>
      </div>
    </div>
  </aside>

  <!-- Contenido principal -->
  <div id="app-main" class="app-main sm:ml-72 pt-20">
    <main class="max-w-screen-xl mx-auto p-6">
      @yield('content')
    </main>
  </div>

  <!-- MODAL: Cambiar de Módulo -->
  <div id="select-modal" tabindex="-1" aria-hidden="true"
       class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
    <div class="relative p-4 w-full max-w-md max-h-full">
      <div class="relative bg-white rounded-lg shadow-sm">
        <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
          <h3 class="text-lg font-semibold text-gray-900">Cambiar de Módulo</h3>
          <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm h-8 w-8 ms-auto inline-flex justify-center items-center" data-modal-toggle="select-modal" aria-label="Cerrar">
            <svg class="w-3 h-3" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 14">
              <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"/>
            </svg>
          </button>
        </div>
        <div class="p-4 md:p-5">
          <ul class="space-y-3">
            <li>
              <a href="/epp" class="flex items-center justify-between w-full p-4 border border-slate-200 rounded-lg hover:bg-slate-50">
                <div class="font-semibold">Entrega de EPP</div>
                <svg class="w-4 h-4 text-slate-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 10">
                  <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 5h12m0 0L9 1m4 4L9 9"/>
                </svg>
              </a>
            </li>
            <li>
              <a href="/Capacitaciones" class="flex items-center justify-between w-full p-4 border border-slate-200 rounded-lg hover:bg-slate-50">
                <div class="font-semibold">Asistencia a Capacitaciones</div>
                <svg class="w-4 h-4 text-slate-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 10">
                  <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 5h12m0 0L9 1m4 4L9 9"/>
                </svg>
              </a>
            </li>
            <li>
              <a href="/empleadosprestamo" class="flex items-center justify-between w-full p-4 border border-slate-200 rounded-lg hover:bg-slate-50">
                <div class="font-semibold">Control de Préstamos</div>
                <svg class="w-4 h-4 text-slate-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 10">
                  <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 5h12m0 0L9 1m4 4L9 9"/>
                </svg>
              </a>
            </li>
            <li>
              <a href="/matrizpuestos" class="flex items-center justify-between w-full p-4 border border-slate-200 rounded-lg hover:bg-slate-50">
                <div class="font-semibold">Organigrama</div>
                <svg class="w-4 h-4 text-slate-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 10">
                  <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 5h12m0 0L9 1m4 4L9 9"/>
                </svg>
              </a>
            </li>
          </ul>
        </div>
      </div>
    </div>
  </div>

  <!-- =================== TUS MODALES ORIGINALES =================== -->

  <!-- MODAL: Evaluación de Riesgos (Excel) -->
  <div id="modal-evaluacion-riesgos" tabindex="-1" aria-hidden="true"
       class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50
              justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
    <div class="relative p-4 w-full max-w-md max-h-full mx-auto">
      <div class="relative bg-white rounded-lg shadow-sm">
        <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
          <h3 class="text-lg font-semibold text-gray-900">Descargar Formato — Evaluación de Riesgos</h3>
          <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-hide="modal-evaluacion-riesgos" aria-label="Cerrar">
            <svg class="w-3 h-3" viewBox="0 0 14 14" fill="none"><path d="M1 1l12 12M13 1L1 13" stroke="currentColor" stroke-width="2"/></svg>
          </button>
        </div>

        <form id="form-export-evaluacion" action="{{ route('evaluacion.riesgos.export') }}" method="POST" class="p-4 md:p-5">
          @csrf
          <label for="ptm-select" class="block mb-2 text-sm font-medium text-gray-900">Selecciona el Puesto de Trabajo</label>
          <select id="ptm-select" name="ptm_id" class="w-full border border-gray-300 rounded-lg p-2 text-sm focus:ring-1 focus:ring-blue-500" required>
            <option value="">Cargando puestos...</option>
          </select>
          <div class="mt-5 flex justify-end gap-2">
            <button type="button" data-modal-hide="modal-evaluacion-riesgos" class="px-4 py-2 rounded-lg border border-gray-300 text-gray-700">Cancelar</button>
            <button type="submit" class="px-4 py-2 rounded-lg bg-blue-600 text-white hover:bg-blue-700">Descargar Excel</button>
          </div>
        </form>
      </div>
    </div>
  </div>

  <!-- MODAL: Identificación de Riesgos (Excel) -->
  <div id="modal-identificacion-riesgos" tabindex="-1" aria-hidden="true"
       class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
    <div class="relative p-4 w-full max-w-md max-h-full mx-auto">
      <div class="relative bg-white rounded-lg shadow-sm">
        <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
          <h3 class="text-lg font-semibold text-gray-900">Descargar Formato — Identificación de Riesgos</h3>
          <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-hide="modal-identificacion-riesgos" aria-label="Cerrar">
            <svg class="w-3 h-3" viewBox="0 0 14 14" fill="none"><path d="M1 1l12 12M13 1L1 13" stroke="currentColor" stroke-width="2"/></svg>
          </button>
        </div>
        <form id="form-export-identificacion" action="{{ route('identificacion.riesgos.export') }}" method="POST" target="download-frame-riesgos" class="p-4 md:p-5">
          @csrf
          <label for="ptm-select-ident" class="block mb-2 text-sm font-medium text-gray-900">Selecciona el Puesto de Trabajo</label>
          <select id="ptm-select-ident" name="ptm_id" class="w-full border border-gray-300 rounded-lg p-2 text-sm focus:ring-1 focus:ring-blue-500" required>
            <option value="">Cargando puestos...</option>
          </select>
          <div class="mt-5 flex justify-end gap-2">
            <button type="button" data-modal-hide="modal-identificacion-riesgos" class="px-4 py-2 rounded-lg border border-gray-300 text-gray-700">Cancelar</button>
            <button type="submit" class="px-4 py-2 rounded-lg bg-blue-600 text-white hover:bg-blue-700">Descargar Excel</button>
          </div>
        </form>
      </div>
    </div>
  </div>

  <!-- MODAL: Notificación por Empleado -->
  <div id="modal-notificacion-empleado" tabindex="-1" aria-hidden="true"
       class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
    <div class="relative p-4 w-full max-w-md max-h-full mx-auto">
      <div class="relative bg-white rounded-lg shadow-sm">
        <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
          <h3 class="text-lg font-semibold text-gray-900">Descargar Formato — Notificación por Empleado</h3>
          <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-hide="modal-notificacion-empleado" aria-label="Cerrar">
            <svg class="w-3 h-3" viewBox="0 0 14 14" fill="none"><path d="M1 1l12 12M13 1L1 13" stroke="currentColor" stroke-width="2"/></svg>
          </button>
        </div>
        <form id="form-export-notif-empleado" action="{{ route('notificacion.excel.empleado.export') }}" method="POST" class="p-4 md:p-5">
          @csrf
          <label for="ptm-select-emp" class="block mb-2 text-sm font-medium text-gray-900">Selecciona el Puesto de Trabajo</label>
          <select id="ptm-select-emp" name="ptm_id" class="w-full border border-gray-300 rounded-lg p-2 text-sm focus:ring-1 focus:ring-blue-500" required>
            <option value="">Cargando puestos...</option>
          </select>
          <div class="mt-5 flex justify-end gap-2">
            <button type="button" data-modal-hide="modal-notificacion-empleado" class="px-4 py-2 rounded-lg border border-gray-300 text-gray-700">Cancelar</button>
            <button type="submit" class="px-4 py-2 rounded-lg bg-blue-600 text-white hover:bg-blue-700">Descargar Excel</button>
          </div>
        </form>
      </div>
    </div>
  </div>

  <!-- MODAL: Notificación por Puesto -->
  <div id="modal-notificacion-puesto" tabindex="-1" aria-hidden="true"
       class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
    <div class="relative p-4 w-full max-w-md max-h-full mx-auto">
      <div class="relative bg-white rounded-lg shadow-sm">
        <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
          <h3 class="text-lg font-semibold text-gray-900">Descargar Formato — Notificación por Puesto</h3>
          <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-hide="modal-notificacion-puesto" aria-label="Cerrar">
            <svg class="w-3 h-3" viewBox="0 0 14 14" fill="none"><path d="M1 1l12 12M13 1L1 13" stroke="currentColor" stroke-width="2"/></svg>
          </button>
        </div>
        <form id="form-export-notif-puesto" action="{{ route('notificacion.excel.export') }}" method="POST" class="p-4 md:p-5">
          @csrf
          <label for="ptm-select-notif" class="block mb-2 text-sm font-medium text-gray-900">Selecciona el Puesto de Trabajo</label>
          <select id="ptm-select-notif" name="puesto_token" class="w-full border border-gray-300 rounded-lg p-2 text-sm focus:ring-1 focus:ring-blue-500" required>
            <option value="">Cargando puestos...</option>
          </select>
          <div class="mt-5 flex justify-end gap-2">
            <button type="button" data-modal-hide="modal-notificacion-puesto" class="px-4 py-2 rounded-lg border border-gray-300 text-gray-700">Cancelar</button>
            <button type="submit" class="px-4 py-2 rounded-lg bg-blue-600 text-white hover:bg-blue-700">Descargar Excel</button>
          </div>
        </form>
      </div>
    </div>
  </div>

  <!-- Iframe oculto global para descargas -->
  <iframe id="download-frame-riesgos" name="download-frame-riesgos" style="display:none;width:0;height:0;border:0;"></iframe>

  <!-- JS: toggle + persistencia & colapsables & modales -->
  <script>
    // --- Toggle sidebar persistente ---
    (function () {
      const root = document.documentElement;
      const KEY = 'sidebarHidden';
      const btn = document.getElementById('sidebarToggle');
      const saved = localStorage.getItem(KEY);
      if (saved === '1') root.classList.add('sidebar-hidden');
      btn?.addEventListener('click', () => {
        const hidden = root.classList.toggle('sidebar-hidden');
        localStorage.setItem(KEY, hidden ? '1' : '0');
      });
    })();

    // --- Collapsibles del sidebar con persistencia por grupo ---
    (function(){
      const KEY = 'arSidebarGroups';
      const state = JSON.parse(localStorage.getItem(KEY) || '{}'); // {id: true/false}

      document.querySelectorAll('[data-collapse-toggle]').forEach(btn => {
        const id = btn.getAttribute('data-collapse-toggle');
        const panel = document.getElementById(id);
        const chevron = btn.querySelector('[data-chevron]');
        const open = state[id] ?? false;
        if (open) { panel?.classList.remove('hidden'); chevron?.classList.add('rotate-180'); }

        btn.addEventListener('click', () => {
          panel?.classList.toggle('hidden');
          chevron?.classList.toggle('rotate-180');
          state[id] = !panel?.classList.contains('hidden');
          localStorage.setItem(KEY, JSON.stringify(state));
        });
      });
    })();

    // ================= LÓGICA DE TUS MODALES =================

    document.addEventListener('DOMContentLoaded', () => {
      // -------- Evaluación de riesgos --------
      const modalEvalId = 'modal-evaluacion-riesgos';
      const selectEval  = document.getElementById('ptm-select');
      const formEval    = document.getElementById('form-export-evaluacion');
      let loadedEval    = false;

      function loadPuestosEval() {
        if (loadedEval) return;
        fetch('{{ route('evaluacion.riesgos.puestos') }}', { headers: { 'X-Requested-With': 'XMLHttpRequest' }})
          .then(r => r.json())
          .then(rows => {
            selectEval.innerHTML = '<option value="">Seleccione...</option>';
            rows.forEach(r => {
              const opt = document.createElement('option');
              opt.value = r.id_puesto_trabajo_matriz;
              opt.textContent = r.puesto_trabajo_matriz;
              selectEval.appendChild(opt);
            });
            loadedEval = true;
          })
          .catch(() => { selectEval.innerHTML = '<option value="">Error cargando puestos</option>'; });
      }

      document.querySelectorAll('[data-modal-target="'+modalEvalId+'"]').forEach(trigger => {
        trigger.addEventListener('click', () => {
          loadPuestosEval();
          const act = trigger.getAttribute('data-export-action');
          if (act) { formEval.setAttribute('action', act); }
        });
      });

      formEval?.addEventListener('submit', () => {
        document.getElementById(modalEvalId).classList.add('hidden');
      });
      formEval?.addEventListener('submit', () => {
        formEval.setAttribute('target', 'download-frame-riesgos');
        const closeBtn = document.querySelector(`[data-modal-hide="${modalEvalId}"]`);
        if (closeBtn) closeBtn.click();
        setTimeout(() => {
          document.querySelectorAll('div[modal-backdrop], .modal-backdrop').forEach(el => el.remove());
          document.documentElement.classList.remove('overflow-hidden');
          document.body.classList.remove('overflow-hidden');
        }, 100);
        setTimeout(() => { window.location.reload(); }, 1200);
      });

      // -------- Identificación de riesgos --------
      const modalIdentId = 'modal-identificacion-riesgos';
      const selectIdent  = document.getElementById('ptm-select-ident');
      const formIdent    = document.getElementById('form-export-identificacion');
      let loadedIdent    = false;

      function loadPuestosIdent() {
        if (loadedIdent) return;
        fetch('{{ route('evaluacion.riesgos.puestos') }}', { headers: { 'X-Requested-With': 'XMLHttpRequest' }})
          .then(r => r.json())
          .then(rows => {
            selectIdent.innerHTML = '<option value="">Seleccione...</option>';
            rows.forEach(r => {
              const opt = document.createElement('option');
              opt.value = r.id_puesto_trabajo_matriz;
              opt.textContent = r.puesto_trabajo_matriz;
              selectIdent.appendChild(opt);
            });
            loadedIdent = true;
          })
          .catch(() => { selectIdent.innerHTML = '<option value="">Error cargando puestos</option>'; });
      }

      document.querySelectorAll('[data-modal-target="'+modalIdentId+'"]').forEach(trigger => {
        trigger.addEventListener('click', loadPuestosIdent);
      });

      formIdent?.addEventListener('submit', () => {
        formIdent.setAttribute('target', 'download-frame-riesgos');
        const closeBtn = document.querySelector(`[data-modal-hide="${modalIdentId}"]`);
        if (closeBtn) closeBtn.click();
        setTimeout(() => {
          document.querySelectorAll('div[modal-backdrop], .modal-backdrop').forEach(el => el.remove());
          document.documentElement.classList.remove('overflow-hidden');
          document.body.classList.remove('overflow-hidden');
        }, 100);
        setTimeout(() => { window.location.reload(); }, 1200);
      });

      // -------- Notificación por Empleado --------
      const modalEmpId = 'modal-notificacion-empleado';
      const selectEmp  = document.getElementById('ptm-select-emp');
      const formEmp    = document.getElementById('form-export-notif-empleado');
      let loadedEmp    = false;

      function loadPuestosEmp() {
        if (loadedEmp) return;
        fetch('{{ route('notificacion.excel.puestos') }}', { headers: { 'X-Requested-With': 'XMLHttpRequest' }})
          .then(r => r.json())
          .then(rows => {
            selectEmp.innerHTML = '<option value="">Seleccione...</option>';
            rows.forEach(r => {
              if (r.source !== 'matriz') {
                return;
              }
              const tokenParts = (r.token || '').split(':');
              const opt = document.createElement('option');
              opt.value = tokenParts[1] || '';
              opt.textContent = r.label;
              selectEmp.appendChild(opt);
            });
            loadedEmp = true;
          })
          .catch(() => { selectEmp.innerHTML = '<option value="">Error cargando puestos</option>'; });
      }

      document.querySelectorAll('[data-modal-target="'+modalEmpId+'"]').forEach(trigger => {
        trigger.addEventListener('click', loadPuestosEmp);
      });

      formEmp?.addEventListener('submit', () => {
        formEmp.setAttribute('target', 'download-frame-riesgos');
        setTimeout(() => {
          const closeBtn = document.querySelector(`[data-modal-hide="${modalEmpId}"]`);
          if (closeBtn) closeBtn.click();
          setTimeout(() => {
            document.querySelectorAll('div[modal-backdrop], .modal-backdrop').forEach(el => el.remove());
            document.documentElement.classList.remove('overflow-hidden');
            document.body.classList.remove('overflow-hidden');
            selectEmp.value = '';
          }, 100);
        }, 400);
      });

      // -------- Notificación por Puesto --------
      const modalNPId = 'modal-notificacion-puesto';
      const selectNP  = document.getElementById('ptm-select-notif');
      const formNP    = document.getElementById('form-export-notif-puesto');
      let loadedNP    = false;

      function loadPuestosNotif() {
        if (loadedNP) return;
        fetch('{{ route('notificacion.excel.puestos') }}', { headers: { 'X-Requested-With': 'XMLHttpRequest' }})
          .then(r => r.json())
          .then(rows => {
            selectNP.innerHTML = '<option value="">Seleccione...</option>';
            rows.forEach(r => {
              const opt = document.createElement('option');
              opt.value = r.token;
              opt.textContent = r.source === 'sistema' ? r.label + ' (Sistema)' : r.label;
              selectNP.appendChild(opt);
            });
            loadedNP = true;
          })
          .catch(() => { selectNP.innerHTML = '<option value="">Error cargando puestos</option>'; });
      }

      document.querySelectorAll('[data-modal-target="'+modalNPId+'"]').forEach(trigger => {
        trigger.addEventListener('click', loadPuestosNotif);
      });

      formNP?.addEventListener('submit', () => {
        formNP.setAttribute('target', 'download-frame-riesgos');
        const closeBtn = document.querySelector(`[data-modal-hide="${modalNPId}"]`);
        if (closeBtn) closeBtn.click();
        setTimeout(() => {
          document.querySelectorAll('div[modal-backdrop], .modal-backdrop').forEach(el => el.remove());
          document.documentElement.classList.remove('overflow-hidden');
          document.body.classList.remove('overflow-hidden');
        }, 100);
      });
    });
  </script>

  <!-- Agrega este iframe en el layout (una sola vez), por ejemplo al final del <body> -->
<iframe name="download-frame-riesgos" id="download-frame-riesgos" style="display:none;"></iframe>

</body>
</html>
