@extends('layouts.mediciones')
@section('title', 'Captura de mediciones (Ruido + Iluminación)')

@section('content')

<!-- ==== Fallback de estilos (por si Tailwind no está activo) ==== -->
<style>
  .control{ width:100%; padding:.5rem .75rem; border:1px solid #cbd5e1; border-radius:12px; box-sizing:border-box; background:#fff; }
  .control--xs{ width:8rem; } .control--sm{ width:10rem; } .control--md{ width:14rem; } .control--lg{ width:16rem; } .control--right{ text-align:right; }
  .btn{ padding:.5rem .75rem; border-radius:12px; border:1px solid #e5e7eb; background:#f9fafb; cursor:pointer; }
  .btn-primary{ background:#00B0F0; color:#fff; border-color:#00B0F0 } .btn:hover{ filter:brightness(0.98); }
  .thead-brand th{ background:#00B0F0; color:#fff; font-weight:600 } table{ width:100%; border-collapse:collapse; } th,td{ border-bottom:1px solid #e5e7eb; }
  .section-card{ border:1px solid #e5e7eb; border-radius:16px; padding:1rem; background:#fff; }
  .muted{ font-size:.8rem; color:#6b7280; }

  /* ====== Layout forzado: aside IZQ, form DER ====== */
  .layout-2col{
    display:grid !important;
    grid-template-columns: 260px minmax(0,1fr) !important;
    grid-template-areas: "aside main" !important;
    gap:1.5rem;
    align-items:start;
  }
  .layout-aside{ grid-area:aside !important; width:260px; position:sticky; top:1rem; }
  .layout-main{  grid-area:main  !important; min-width:0; }

  /* 🔽 En móvil apilar: form arriba, aside abajo */
  @media (max-width:768px){
    .layout-2col{
      grid-template-columns: 1fr !important;
      grid-template-areas:
        "main"
        "aside" !important;
    }
    .layout-aside{ width:100%; position:static; }
  }
</style>

<div class="p-6">
  <div class="layout-2col gap-6">

    <!-- ======= MAIN: Formulario ======= -->
    <div class="layout-main space-y-6">
      @if($errors->any())
        <div class="rounded-xl border border-rose-200 bg-rose-50 text-rose-800 px-4 py-3">
          <ul class="list-disc ml-5">@foreach($errors->all() as $e)<li>{{ $e }}</li>@endforeach</ul>
        </div>
      @endif
      @if(session('ok'))
        <div class="rounded-xl border border-green-200 bg-green-50 text-green-800 px-4 py-3">{{ session('ok') }}</div>
      @endif

      <h1 class="text-2xl font-bold">Captura de mediciones</h1>

      <form method="POST" action="{{ route('mediciones.captura.store') }}" class="space-y-8">
        @csrf

        {{-- CABECERA COMÚN --}}
        <div class="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
          <div>
            <label class="block text-sm font-medium text-gray-700">Departamento *</label>
            <input name="departamento" value="{{ old('departamento') }}" class="control">
          </div>

          <div>
            <label class="block text-sm font-medium text-gray-700">Localización *</label>
            <input list="dl-localizaciones" id="localizacion_txt" class="control" placeholder="Escribe y elige…" value="{{ old('localizacion_txt') }}">
            <datalist id="dl-localizaciones">
              @foreach($localizaciones as $lo)
                <option value="{{ $lo->localizacion }}" data-id="{{ $lo->id_localizacion }}"></option>
              @endforeach
            </datalist>
            <input type="hidden" name="id_localizacion" id="id_localizacion" value="{{ old('id_localizacion') }}">
            <small id="loc_help" class="muted"></small>
            @error('id_localizacion')<p class="text-red-600 text-sm">{{ $message }}</p>@enderror
          </div>

          <div>
            <label class="block text-sm font-medium text-gray-700">Observador(a) *</label>
            <input name="nombre_observador" value="{{ old('nombre_observador') }}" class="control">
          </div>

          <div>
            <label class="block text-sm font-medium text-gray-700">Fecha inicio</label>
            <input type="date" name="fecha_realizacion_inicio" value="{{ old('fecha_realizacion_inicio') }}" class="control">
          </div>

          <div>
            <label class="block text-sm font-medium text-gray-700">Fecha fin</label>
            <input type="date" name="fecha_realizacion_final" value="{{ old('fecha_realizacion_final') }}" class="control">
          </div>
        </div>

        {{-- Datalist global de Puestos (toda la lista, sin filtrar) --}}
        <datalist id="dl-puestos">
          @foreach($puestos as $p)
            <option value="{{ $p->puesto_trabajo_matriz }}" data-id="{{ $p->id_puesto_trabajo_matriz }}"></option>
          @endforeach
        </datalist>

        {{-- ================= RUIDO ================= --}}
        <div class="section-card" style="background:#F5F8FA;">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:1rem;">
            <h2 class="text-xl font-semibold text-gray-800" style="margin:0;">Ruido</h2>
            <button type="button" id="btn-add-ruido" class="btn">+ Punto de ruido</button>
          </div>

          <div class="grid md:grid-cols-4 gap-4 mt-3">
            <div>
              <label class="block text-sm font-medium text-gray-700">Instrumento *</label>
              <input name="instrumento_ruido" class="control" maxlength="150" value="{{ old('instrumento_ruido', 'Sonómetro') }}">
            </div>
            <div>
              <label class="block text-sm font-medium text-gray-700">Serie</label>
              <input name="serie_ruido" class="control" maxlength="200" value="{{ old('serie_ruido', '2017050C2100') }}">
            </div>
            <div>
              <label class="block text-sm font-medium text-gray-700">Marca</label>
              <input name="marca_ruido" class="control" maxlength="100" value="{{ old('marca_ruido', 'CE modelo SKU 161-600001-32') }}">
            </div>
            <div>
              <label class="block text-sm font-medium text-gray-700">NRR</label>
              <input type="number" step="0.01" name="nrr" value="{{ old('nrr') }}" class="control">
            </div>
          </div>

          <div class="overflow-x-auto mt-4">
            <table id="tbl-ruido">
              <thead class="thead-brand">
                <tr>
                  <th class="px-2 py-2 text-left">Punto *</th>
                  <th class="px-2 py-2 text-left">Puesto de trabajo *</th>
                  <th class="px-2 py-2 text-right">Máx</th>
                  <th class="px-2 py-2 text-right">Mín</th>
                  <th class="px-2 py-2 text-right">Prom</th>
                  <th class="px-2 py-2 text-right">NRE</th>
                  <th class="px-2 py-2 text-right">Límite *</th>
                  <th class="px-2 py-2">Obs</th>
                  <th class="px-2 py-2"></th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </div>

        {{-- ============== ILUMINACIÓN ============== --}}
        <div class="section-card" style="background:#F5F8FA;">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:1rem;">
            <h2 class="text-xl font-semibold text-gray-800" style="margin:0;">Iluminación</h2>
            <button type="button" id="btn-add-lux" class="btn">+ Punto de iluminación</button>
          </div>

          <div class="grid md:grid-cols-3 gap-4 mt-3">
            <div>
              <label class="block text-sm font-medium text-gray-700">Instrumento *</label>
              <input name="instrumento_lux" class="control" maxlength="150" value="{{ old('instrumento_lux', 'Luxómetro') }}">
            </div>
            <div>
              <label class="block text-sm font-medium text-gray-700">Serie</label>
              <input name="serie_lux" class="control" maxlength="200" value="{{ old('serie_lux', 'X0005MAD2') }}">
            </div>
            <div>
              <label class="block text-sm font-medium text-gray-700">Marca</label>
              <input name="marca_lux" class="control" maxlength="100" value="{{ old('marca_lux', 'Meter modelo LX1330B') }}">
            </div>
          </div>

          <div class="overflow-x-auto mt-4">
            <table id="tbl-lux">
              <thead class="thead-brand">
                <tr>
                  <th class="px-2 py-2 text-left">Punto *</th>
                  <th class="px-2 py-2 text-left">Puesto de trabajo *</th>
                  <th class="px-2 py-2 text-right">Prom (lux) *</th>
                  <th class="px-2 py-2 text-right">Límite (lux) *</th>
                  <th class="px-2 py-2">Obs</th>
                  <th class="px-2 py-2"></th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </div>

        <div style="display:flex;gap:.75rem;">
          <button type="submit" class="btn btn-primary">Guardar mediciones</button>
          <button type="reset" class="btn">Limpiar</button>
        </div>
      </form>

      {{-- ===== Datos para JS (puestos) ===== --}}
      <script>
        window.PUESTOS = @json($puestos);
        window.PUESTO_BY_ID = {};
        for (const p of window.PUESTOS) {
          window.PUESTO_BY_ID[p.id_puesto_trabajo_matriz] = p.puesto_trabajo_matriz;
        }
      </script>

      {{-- ===== JS: filas dinámicas + datalist de localización/puestos (versión limpia) ===== --}}
      <script>
      (function(){
        // ---------- Referencias base ----------
        const form     = document.querySelector('form[action="{{ route('mediciones.captura.store') }}"]') || document.querySelector('form');
        const inputTxt = document.getElementById('localizacion_txt');
        const hiddenId = document.getElementById('id_localizacion');
        const help     = document.getElementById('loc_help');
        const locOpts  = Array.from(document.querySelectorAll('#dl-localizaciones option'));

        const dlPuestos = document.getElementById('dl-puestos');
        const ruidoBody = document.querySelector('#tbl-ruido tbody');
        const luxBody   = document.querySelector('#tbl-lux tbody');

        let idxR = 0, idxL = 0;
        const prefillUrl = @json(route('mediciones.captura.prefill'));

        // ---------- Helpers ----------
        function syncPuestoInput(txtEl, hiddenEl) {
          const val = (txtEl.value || '').trim();
          const opt = Array.from(dlPuestos.options).find(o => o.value === val);
          hiddenEl.value = opt ? opt.dataset.id : '';
        }
        function q(sel){ return document.querySelector(sel); }
        function setIf(elSelector, val){
          const el = q(elSelector);
          if (el && val != null && val !== '') el.value = val;
        }

        // ---------- Filas dinámicas ----------
        function addRuidoRow() {
          const rowIdx = idxR++;
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td class="px-2 py-1"><input name="ruido_puntos[${rowIdx}][punto_medicion]" class="control control--md"></td>
            <td class="px-2 py-1">
              <input list="dl-puestos" id="ruido_puesto_txt_${rowIdx}" class="control control--lg" placeholder="Escribe y elige puesto…">
              <input type="hidden" name="ruido_puntos[${rowIdx}][id_puesto_trabajo_matriz]" id="ruido_puesto_id_${rowIdx}">
            </td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="ruido_puntos[${rowIdx}][nivel_maximo]" class="control control--xs control--right"></td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="ruido_puntos[${rowIdx}][nivel_minimo]" class="control control--xs control--right"></td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="ruido_puntos[${rowIdx}][nivel_promedio]" class="control control--xs control--right"></td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="ruido_puntos[${rowIdx}][nre]" class="control control--xs control--right"></td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="ruido_puntos[${rowIdx}][limites_aceptables]" class="control control--xs control--right" value="85"></td>
            <td class="px-2 py-1"><input name="ruido_puntos[${rowIdx}][observaciones]" class="control control--lg"></td>
            <td class="px-2 py-1" style="text-align:center;"><button type="button" class="rm-row btn" aria-label="Eliminar fila">✕</button></td>
          `;
          ruidoBody.appendChild(tr);
          const txt = document.getElementById(`ruido_puesto_txt_${rowIdx}`);
          const hid = document.getElementById(`ruido_puesto_id_${rowIdx}`);
          txt.addEventListener('change', () => syncPuestoInput(txt, hid));
          txt.addEventListener('blur',   () => syncPuestoInput(txt, hid));
          return { rowIdx, tr };
        }

        function addLuxRow() {
          const rowIdx = idxL++;
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td class="px-2 py-1"><input name="iluminacion_puntos[${rowIdx}][punto_medicion]" class="control control--md"></td>
            <td class="px-2 py-1">
              <input list="dl-puestos" id="lux_puesto_txt_${rowIdx}" class="control control--lg" placeholder="Escribe y elige puesto…">
              <input type="hidden" name="iluminacion_puntos[${rowIdx}][id_puesto_trabajo_matriz]" id="lux_puesto_id_${rowIdx}">
            </td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="iluminacion_puntos[${rowIdx}][promedio]" class="control control--xs control--right"></td>
            <td class="px-2 py-1 text-right"><input type="number" step="0.01" name="iluminacion_puntos[${rowIdx}][limites_aceptables]" class="control control--xs control--right"></td>
            <td class="px-2 py-1"><input name="iluminacion_puntos[${rowIdx}][observaciones]" class="control control--lg"></td>
            <td class="px-2 py-1" style="text-align:center;"><button type="button" class="rm-row btn" aria-label="Eliminar fila">✕</button></td>
          `;
          luxBody.appendChild(tr);
          const txt = document.getElementById(`lux_puesto_txt_${rowIdx}`);
          const hid = document.getElementById(`lux_puesto_id_${rowIdx}`);
          txt.addEventListener('change', () => syncPuestoInput(txt, hid));
          txt.addEventListener('blur',   () => syncPuestoInput(txt, hid));
          return { rowIdx, tr };
        }

        function syncAllPuestos() {
          for (let i=0;i<idxR;i++) {
            const txt = document.getElementById(`ruido_puesto_txt_${i}`);
            const hid = document.getElementById(`ruido_puesto_id_${i}`);
            if (txt && hid) syncPuestoInput(txt, hid);
          }
          for (let i=0;i<idxL;i++) {
            const txt = document.getElementById(`lux_puesto_txt_${i}`);
            const hid = document.getElementById(`lux_puesto_id_${i}`);
            if (txt && hid) syncPuestoInput(txt, hid);
          }
        }

        // ---------- Prefill ----------
        function applyTemplate(data){
          if (!data || !data.ok) return;

          // Cabecera sugerida
          if (data.base) {
            setIf('input[name="departamento"]',      data.base.departamento);
            setIf('input[name="instrumento_ruido"]', data.base.instrumento_ruido);
            setIf('input[name="serie_ruido"]',       data.base.serie_ruido);
            setIf('input[name="marca_ruido"]',       data.base.marca_ruido);
            setIf('input[name="nrr"]',               data.base.nrr);
            setIf('input[name="instrumento_lux"]',   data.base.instrumento_lux);
            setIf('input[name="serie_lux"]',         data.base.serie_lux);
            setIf('input[name="marca_lux"]',         data.base.marca_lux);
          }

          // Limpiar tablas
          ruidoBody.innerHTML = ''; luxBody.innerHTML = '';
          idxR = 0; idxL = 0;

          // Rellenar RUIdo
          (data.ruido_rows || []).forEach(row => {
            const { rowIdx } = addRuidoRow();
            const puestoName = (window.PUESTO_BY_ID && window.PUESTO_BY_ID[row.id_puesto_trabajo_matriz]) || '';
            setIf(`input[name="ruido_puntos[${rowIdx}][punto_medicion]"]`, row.punto_medicion || '');
            setIf(`#ruido_puesto_id_${rowIdx}`, row.id_puesto_trabajo_matriz || '');
            setIf(`#ruido_puesto_txt_${rowIdx}`, puestoName);
            if (row.limites_aceptables != null) {
              setIf(`input[name="ruido_puntos[${rowIdx}][limites_aceptables]"]`, row.limites_aceptables);
            }
          });

          // Rellenar ILUMINACIÓN
          (data.lux_rows || []).forEach(row => {
            const { rowIdx } = addLuxRow();
            const puestoName = (window.PUESTO_BY_ID && window.PUESTO_BY_ID[row.id_puesto_trabajo_matriz]) || '';
            setIf(`input[name="iluminacion_puntos[${rowIdx}][punto_medicion]"]`, row.punto_medicion || '');
            setIf(`#lux_puesto_id_${rowIdx}`, row.id_puesto_trabajo_matriz || '');
            setIf(`#lux_puesto_txt_${rowIdx}`, puestoName);
            if (row.limites_aceptables != null) {
              setIf(`input[name="iluminacion_puntos[${rowIdx}][limites_aceptables]"]`, row.limites_aceptables);
            }
          });

          // Fallback: si no hay filas en la plantilla, deja una vacía de cada tipo
          if ((data.ruido_rows || []).length === 0) addRuidoRow();
          if ((data.lux_rows  || []).length === 0) addLuxRow();

          // Ayuda visual
          if (help) {
            const yrR = data.year_ruido ? `Ruido: ${data.year_ruido}` : 'Ruido: sin datos';
            const yrL = data.year_lux  ? `Iluminación: ${data.year_lux}` : 'Iluminación: sin datos';
            help.textContent = `Plantilla aplicada (${yrR} / ${yrL}).`;
          }
        }

        async function fetchAndApplyPrefill(){
          if (!hiddenId || !hiddenId.value) return;
          try{
            const url = `${prefillUrl}?id_localizacion=${encodeURIComponent(hiddenId.value)}`;
            const res = await fetch(url, { headers: { 'Accept':'application/json' }});
            if (!res.ok) return;
            const data = await res.json();
            applyTemplate(data);
          }catch(err){
            console.error('Prefill error', err);
          }
        }

        // ---------- Localización: datalist -> hidden + prefill ----------
        function syncLoc() {
          const val = (inputTxt.value || '').trim();
          const opt = locOpts.find(o => o.value === val);
          if (opt) {
            hiddenId.value = opt.dataset.id;
            help.textContent = '';
            fetchAndApplyPrefill(); // cargar plantilla al tener ID válido
          } else {
            hiddenId.value = '';
            help.textContent = '⚠️ Selecciona una opción de la lista para registrar el ID.';
          }
        }
        inputTxt.addEventListener('change', syncLoc);
        inputTxt.addEventListener('blur',   syncLoc);
        // Si el aside cambia el hidden (dispatchEvent('change')), también prefilleamos
        hiddenId.addEventListener('change', fetchAndApplyPrefill);

        // ---------- Form submit: asegurar hidden de puestos ----------
        form.addEventListener('submit', function(e){
          syncLoc();
          syncAllPuestos();
          if (!hiddenId.value) {
            e.preventDefault();
            inputTxt.focus();
            help.textContent = '⚠️ Debes elegir una localización de la lista.';
          }
        });

        // ---------- Botones agregar fila ----------
        document.getElementById('btn-add-ruido').addEventListener('click', addRuidoRow);
        document.getElementById('btn-add-lux').addEventListener('click', addLuxRow);

        // ---------- Eliminar fila ----------
        document.addEventListener('click', function(e){
          if (e.target.classList.contains('rm-row')) {
            e.target.closest('tr')?.remove();
          }
        });

        // ---------- Estado inicial ----------
        if (hiddenId && hiddenId.value) {
          // Si viene de old('id_localizacion'), intenta prefillear
          fetchAndApplyPrefill();
        } else {
          // Si no hay selección, deja una fila de cada tipo
          addRuidoRow();
          addLuxRow();
        }
      })();
      </script>
    </div>
    <!-- ======= /MAIN ======= -->

    <!-- ======= ASIDE ======= -->
    <aside class="layout-aside space-y-4">
      <div class="section-card">
        <h2 class="text-lg font-semibold mb-2">Localizaciones (año {{ $year ?? now()->year }})</h2>
        <p class="muted">Marcadas las que ya tienen mediciones de este año.</p>
        <div class="mt-3">
          <style>
            .dot{display:inline-block;width:.6rem;height:.6rem;border-radius:9999px;margin-right:.4rem;vertical-align:middle}
            .dot.ok{background:#16a34a}
            .dot.no{background:#d1d5db}
            .loc-item{display:flex;align-items:center;justify-content:space-between;padding:.35rem .5rem;border-radius:.5rem; cursor:pointer}
            .loc-item.active{background:#eff6ff}
            .badge{font-size:.7rem; background:#f1f5f9; border:1px solid #e5e7eb; border-radius:.4rem; padding:.05rem .35rem; margin-left:.25rem}
          </style>
          <ul style="list-style:none;padding:0;margin:0;max-height:70vh;overflow:auto">
            @foreach(($locStatus ?? collect()) as $ls)
              <li class="loc-item" data-id="{{ $ls->id }}" data-name="{{ $ls->nombre }}">
                <span>
                  <i class="dot {{ $ls->has ? 'ok' : 'no' }}"></i>
                  {{ $ls->nombre }}
                </span>
                <span class="muted">
                  <span class="badge">Lx: {{ $ls->cnt_lux }}</span>
                  <span class="badge">Rd: {{ $ls->cnt_ruido }}</span>
                </span>
              </li>
            @endforeach
          </ul>
        </div>
      </div>

      <div class="section-card">
        <h3 class="font-semibold mb-1">Tip</h3>
        <p class="muted">Selecciona la localización en el campo del formulario; aquí se resalta si ya tiene datos.</p>
      </div>

      <script>
        // Resalta la localización seleccionada en el listado y permite clickear para seleccionarla
        (function(){
          const hiddenId = document.getElementById('id_localizacion');
          const inputTxt = document.getElementById('localizacion_txt');
          const items = document.querySelectorAll('.loc-item');
          function highlight(){
            const sel = hiddenId && hiddenId.value;
            items.forEach(li => {
              if (!sel) { li.classList.remove('active'); return; }
              li.classList.toggle('active', li.getAttribute('data-id') === sel);
            });
          }
          // Al hacer clic en una localización, asignarla al formulario
          items.forEach(li => {
            li.addEventListener('click', () => {
              const id = li.getAttribute('data-id');
              const name = li.getAttribute('data-name');
              if (inputTxt) inputTxt.value = name || '';
              if (hiddenId) { hiddenId.value = id || ''; hiddenId.dispatchEvent(new Event('change')); }
              highlight();
            });
          });
          if (hiddenId) {
            const t = setInterval(highlight, 300);
            setTimeout(() => clearInterval(t), 5000); // primeras actualizaciones
            hiddenId.addEventListener('change', highlight);
          }
          highlight();
        })();
      </script>
    </aside>
    <!-- ======= /ASIDE ======= -->

  </div> <!-- /.layout-2col -->
</div>   <!-- /.p-6 -->

@endsection
