<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>@yield('title', 'Préstamos')</title>

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

  <!-- Bootstrap (opcional) -->
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
            <p class="text-lg font-semibold">Control de Préstamos</p>
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
          Préstamos · Menú
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
          <li>
            <a href="/empleadosprestamo" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('empleadosprestamo*') }}">
              <!-- users -->
              <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M16 14a4 4 0 10-8 0m8 0a4 4 0 01-8 0m8 0v1a4 4 0 004 4H4a4 4 0 004-4v-1"/>
              </svg>
              <span class="text-sm font-medium">Resumen por Empleado</span>
            </a>
          </li>

          <li>
            <a href="/infoprestamo" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('infoprestamo*') }}">
              <!-- document -->
              <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M7 3h8l4 4v14a2 2 0 01-2 2H7a2 2 0 01-2-2V5a2 2 0 012-2z"/>
              </svg>
              <span class="text-sm font-medium">Préstamos</span>
            </a>
          </li>

          <li>
            <a href="/cuotas" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('cuotas*') }}">
              <!-- list -->
              <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 7h16M4 12h16M4 17h10"/>
              </svg>
              <span class="text-sm font-medium">Historial / Cuotas</span>
            </a>
          </li>

          <li>
            <a href="/cuotas-especiales" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('cuotas-especiales*') }}">
              <!-- cash -->
              <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M3 7h18v10H3zM7 12h.01M17 12h.01M12 12a3 3 0 110-6 3 3 0 010 6z"/>
              </svg>
              <span class="text-sm font-medium">Depósitos y Cobros Extraordinarios</span>
            </a>
          </li>

          <li>
            <a href="/cuotas/rango" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('cuotas-especiales*') }}">
              <!-- cash -->
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M8 12h13M8 6h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
              </svg>
              <span class="text-sm font-medium">Ver cuotas por Rango de Fechas</span>
            </a>
          </li>

          <li class="mt-4 px-2 text-xs font-semibold tracking-wide text-slate-500">Reportes</li>

          <li>
            <a href="/prestamos/resumen-mensual" class="group flex items-center gap-3 rounded-xl px-3 py-2 {{ $activeLink('prestamos/resumen-mensual*') }}">
              <!-- calendar -->
              <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M8 7V3m8 4V3M3 11h18M5 21h14a2 2 0 002-2V7H3v12a2 2 0 002 2z"/>
              </svg>
              <span class="text-sm font-medium">Resumen Mensual</span>
            </a>
          </li>

          <li>
            <!-- Abrir Modal de Descarga -->
            <button id="open-reporte-prestamos"
                    type="button"
                    class="w-full text-left group flex items-center gap-3 rounded-xl px-3 py-2 text-slate-700 hover:bg-slate-100 border border-transparent">
              <!-- download -->
              <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5 flex-none" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M12 4v10m0 0l-3.5-3.5M12 14l3.5-3.5M5 20h14"/>
              </svg>
              <span class="text-sm font-medium">Descargar Planilla de Deducciones…</span>
            </button>
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
              <a href="/verificacion" class="flex items-center justify-between w-full p-4 border border-slate-200 rounded-lg hover:bg-slate-50">
                <div class="font-semibold">Análisis de Riesgos</div>
                <svg class="w-4 h-4 text-slate-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 10">
                  <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 5h12m0 0L9 1m4 4L9 9"/>
                </svg>
              </a>
            </li>
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

  {{-- Modal: Reporte de Préstamos (rango de fechas) --}}
  <div id="modal-reporte-prestamos" class="fixed inset-0 z-50 hidden">
    <!-- Fondo -->
    <div class="absolute inset-0 bg-black/30" data-close></div>

    <!-- Contenido -->
    <div class="relative mx-auto my-10 w-full max-w-md rounded-xl bg-white shadow-lg">
      <div class="p-5 border-b border-gray-200">
        <h3 class="text-lg font-semibold text-gray-900">Rango de fechas</h3>
        <p class="text-sm text-gray-500">Selecciona el rango para la planilla de deducciones.</p>
      </div>

      <form id="form-reporte-prestamos" method="GET" action="{{ route('prestamos.reporte.excel') }}" class="p-5 space-y-4" target="download-frame">
        <div>
          <label class="block mb-2 text-sm font-medium text-gray-900">Fecha Inicio</label>
          <input type="date" name="fecha_inicio" required
                 value="{{ old('fecha_inicio', date('Y-m-01')) }}"
                 class="w-full rounded-lg border border-gray-300 bg-gray-50 p-2.5 focus:ring-brand focus:border-brand"/>
        </div>

        <div>
          <label class="block mb-2 text-sm font-medium text-gray-900">Fecha Final</label>
          <input type="date" name="fecha_final" required
                 value="{{ old('fecha_final', date('Y-m-d')) }}"
                 class="w-full rounded-lg border border-gray-300 bg-gray-50 p-2.5 focus:ring-brand focus:border-brand"/>
        </div>

        <div class="flex justify-end gap-3 pt-2">
          <button type="button" data-close
                  class="rounded-lg border border-gray-300 bg-white px-4 py-2 text-gray-700 hover:bg-gray-50">
            Cancelar
          </button>
          <button type="submit"
                  class="rounded-lg bg-blue-600 px-4 py-2 font-semibold text-white hover:bg-blue-700">
            Descargar
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- Iframe oculto para realizar descargas sin bloquear la vista -->
  <iframe id="download-frame" name="download-frame" style="display:none;width:0;height:0;border:0;"></iframe>

  <!-- JS: toggle + persistencia & modal reporte -->
  <script>
    (function () {
      const root = document.documentElement;
      const KEY = 'sidebarHidden';
      const btn = document.getElementById('sidebarToggle');

      // Aplica preferencia guardada
      const saved = localStorage.getItem(KEY);
      if (saved === '1') root.classList.add('sidebar-hidden');

      // Toggle en escritorio
      btn?.addEventListener('click', () => {
        const hidden = root.classList.toggle('sidebar-hidden');
        localStorage.setItem(KEY, hidden ? '1' : '0');
      });
    })();

    // -------- Modal Reporte Préstamos ----------
    document.addEventListener('DOMContentLoaded', () => {
      const openBtn = document.getElementById('open-reporte-prestamos');
      const modal = document.getElementById('modal-reporte-prestamos');
      const form = document.getElementById('form-reporte-prestamos');
      const dlFrame = document.getElementById('download-frame');

      function openModal(e) {
        e?.preventDefault();
        modal.classList.remove('hidden');
        modal.classList.add('flex', 'items-center', 'justify-center');
      }
      function closeModal() {
        modal.classList.add('hidden');
        modal.classList.remove('flex', 'items-center', 'justify-center');
      }

      openBtn?.addEventListener('click', openModal);

      modal.querySelectorAll('[data-close]').forEach(el => el.addEventListener('click', closeModal));

      document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModal(); });

      // Enviar via iframe (no bloquea vista)
      form?.addEventListener('submit', (e) => {
        // Permitimos que el form se envíe al target="download-frame"
        // para no bloquear la vista ni navegar. Luego cerramos y refrescamos.
        setTimeout(closeModal, 100);
        setTimeout(() => { window.location.reload(); }, 1200);
      });
    });
  </script>
</body>
</html>
