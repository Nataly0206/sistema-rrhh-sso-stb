@extends('layouts.capacitacion')
@section('title', 'Registro de Asistencia a Capacitaciones')

@section('content')
<style>
  .card{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:1rem}
  .btn{padding:.55rem .85rem;border-radius:12px;border:1px solid #e5e7eb;background:#f9fafb;cursor:pointer}
  .btn-primary{background:#00B0F0;color:#fff;border-color:#00B0F0}
  .btn-success{background:#10b981;color:#fff;border-color:#10b981}
  .badge{display:inline-block;padding:.15rem .45rem;border-radius:.5rem;font-weight:700;font-size:.72rem}
  .b-ok{background:#dcfce7;color:#166534}
  .input{border:1px solid #cbd5e1;border-radius:10px;padding:.45rem .65rem}
  .muted{color:#64748b}
  .list{border:1px solid #e5e7eb;border-radius:10px;overflow:hidden}
  .list li{display:flex;align-items:center;justify-content:space-between;padding:.35rem .55rem;border-top:1px solid #f1f5f9}
  .list li:first-child{border-top:0}
  .del{color:#dc2626;cursor:pointer}

  /* ============ PREVIEW IMPRESIÓN (A4) ============ */
  /* … (deja tu css existente) … */

  /* ============ PREVIEW IMPRESIÓN (A4, estilo STB/RRHH/R003) ============ */
  .preview{background:#f8fafc;border:1px dashed #e2e8f0;border-radius:12px;padding:12px}
  .a4{background:#fff;border:1px solid #cbd5e1;border-radius:8px;margin:12px auto;max-width:794px;padding:18px}
  .hdr{display:grid;grid-template-columns:120px 1fr;align-items:center;gap:8px;margin-bottom:8px}
  .hdr img{max-width:100%;height:auto;opacity:.9}
  .hdr-center{text-align:center;line-height:1.15}
  .hdr-center .l1{font-weight:700}
  .hdr-center .l2,.hdr-center .l3,.hdr-center .l4{font-size:.92rem}
  .fname{font-weight:700;margin:8px 0 0 0}
  .line{border-bottom:1px solid #2b2b2b;display:inline-block;min-width:200px}

  .meta{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:10px}
  .meta .row{display:flex;gap:6px;align-items:center}
  .meta label{min-width:80px}
  .yn{display:flex;gap:16px;align-items:center}
  .box{width:18px;height:18px;border:1px solid #111;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:.9rem}
  .box.on::after{content:'X';}

  .table-wrap{margin-top:10px}
  .tbl-r003{width:100%;border-collapse:collapse}
  .tbl-r003 th,.tbl-r003 td{border:1px solid #111;padding:.35rem .45rem;font-size:.9rem}
  .tbl-r003 th{background:#e2e8f0;text-align:center}
  .tbl-r003 td{min-height:26px}
  .col-n{width:38px;text-align:center}
  .col-id{width:190px}
  .col-tel{width:140px}
  .col-puesto{width:130px}
  .col-firma{width:140px;text-align:center}

  .foot{display:flex;justify-content:space-between;font-size:.8rem;margin-top:8px}
  .foot .left{display:flex;gap:18px}
  .foot .right{display:flex;gap:20px}

/* ====== PRINT: solo el formato (.a4) ====== */
@page { size: A4; margin: 12mm; }  /* ajusta margen a tu gusto */

@media print {
  /* Oculta TODO */
  body * { visibility: hidden !important; }

  /* Muestra SOLO el preview (y su contenido) */
  #preview, #preview * { visibility: visible !important; }

  /* Saca el preview del card y quita marcos al imprimir */
  #preview {
    position: absolute !important;
    left: 0; top: 0;
    width: 100%;
    margin: 0 !important; padding: 0 !important; border: 0 !important;
    background: #fff !important;
  }

  /* Cada hoja .a4 en página separada, sin bordes */
  .a4 {
    border: 0 !important;
    box-shadow: none !important;
    page-break-after: always;
    margin: 0 auto !important;
  }
  .a4:last-child { page-break-after: auto; }

  /* Asegura colores de encabezados/tabla en impresión */
  .tbl-r003 th {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
    background: #e2e8f0 !important;
  }

  /* Cualquier UI general */
  .no-print, header, nav, .card, .btn, form, .p-6 > *:not(.card):not(#preview) { display: none !important; }
}
</style>

<div class="p-6 space-y-4">

  {{-- Mensajes --}}
  @if($ok)
    <div class="card" style="border-color:#bbf7d0;background:#ecfdf5">
      <div class="badge b-ok">Éxito</div>
      <div>{{ $ok }}</div>
    </div>
  @endif
  @if($err)
    <div class="card" style="border-color:#fecaca;background:#fef2f2">
      <div class="badge" style="background:#fecaca;color:#7f1d1d">Error</div>
      <div>{{ $err }}</div>
    </div>
  @endif

  <form id="form-asistencia" method="POST" action="{{ route('capacitaciones.registro-asistencia.store') }}">
    @csrf

    <div class="card">
      <h1 class="text-xl font-bold mb-3">1) Datos de la capacitación</h1>
      <div class="grid md:grid-cols-2 gap-3">
        <div>
          <label class="block text-sm muted">Capacitación (capacitacion_instructor)</label>
          <select id="ci-select" name="id_ci" class="input w-full" required>
            <option value="">-- Selecciona --</option>
            @foreach($cis as $ci)
              <option value="{{ $ci->id }}"
                data-tema="{{ $ci->tema }}"
                data-duracion="{{ (int)$ci->duracion }}"
                data-programada="{{ strtoupper(trim($ci->programada ?? ''))==='SI' ? 'SI' : 'NO' }}"
                data-instructor="{{ $ci->instructor_nombre ?? ('ID '.$ci->id_instructor) }}"
              >{{ $ci->tema }}</option>
            @endforeach
          </select>
        </div>
        <div>
          <label class="block text-sm muted">Fecha recibida (texto libre)</label>
          <input type="text" name="fecha_recibida" id="fecha_recibida" class="input w-full" placeholder='Ej. "Del 12/07/2025 al 15/07/2025"' required>
        </div>
        <div>
          <label class="block text-sm muted">Duración</label>
          <input type="text" id="duracion_show" class="input w-full bg-gray-50" readonly placeholder="min">
        </div>
        <div>
          <label class="block text-sm muted">Programada</label>
          <input type="text" id="programada_show" class="input w-full bg-gray-50" readonly>
        </div>
        <div>
          <label class="block text-sm muted">Instructor (referencia)</label>
          <input type="text" id="instructor_show" class="input w-full bg-gray-50" readonly>
        </div>
        <div>
          <label class="block text-sm muted">Instructor temporal (opcional - se guarda)</label>
          <input type="text" name="instructor_temporal" id="instructor_temporal" class="input w-full" placeholder="Nombre del instructor que firmará la planilla">
        </div>
        <div class="md:col-span-2">
          <label class="block text-sm muted">Lugar (solo para impresión)</label>
          <input type="text" id="lugar_print" class="input w-full" placeholder="Bodega, Planta, Sala de juntas... (no se guarda)">
        </div>
        <div class="md:col-span-2 muted text-sm">
          * El instructor mostrado es de <code>capacitacion_instructor</code>. “Instructor temporal” se guarda tal cual en <code>asistencia_capacitacion.instructor_temporal</code>.
        </div>
      </div>
    </div>

    <div class="card">
      <h2 class="text-lg font-bold mb-3">2) Agregar asistentes (Nombre o Identidad)</h2>

      <div class="flex gap-2 items-end no-print">
        <div class="grow">
          <label class="block text-sm muted">Buscar empleado</label>
          <input type="text" id="emp-q" class="input w-full" placeholder="Ej: Ana López / 0801-...">
          <ul id="emp-suggest" class="list" style="display:none; position:absolute; z-index:20; max-height:240px; overflow:auto; width:360px; background:#fff"></ul>
        </div>
        <div>
          <button type="button" id="clear-all" class="btn">Limpiar lista</button>
        </div>
      </div>

      <div class="mt-3">
        <label class="block text-sm muted">Seleccionados</label>
        <ul id="emp-selected" class="list"></ul>
      </div>

      <!-- inputs ocultos -->
      <div id="hidden-inputs"></div>
    </div>

    <div class="flex gap-2 no-print">
      <button type="submit" class="btn btn-primary">Guardar</button>
      <button type="button" id="save-print" class="btn btn-success">Guardar e Imprimir</button>
    </div>
  </form>

  <!-- ============ PREVIEW IMPRESIÓN ============ -->
  <div class="card">
    <h2 class="text-lg font-bold">3) Previsualización para imprimir</h2>
    <p class="muted text-sm">Se generan hojas de 15 registros cada una, replicando los datos de la capacitación.</p>
    <div id="preview" class="preview"></div>
  </div>

</div>

<script>
const routeCiInfo   = "{{ route('capacitaciones.ci-info') }}";
const routeEmpSrch  = "{{ route('capacitaciones.empleados.search') }}";
const empSuggest    = document.getElementById('emp-suggest');
const empSelected   = document.getElementById('emp-selected');
const hiddenInputs  = document.getElementById('hidden-inputs');
const previewEl     = document.getElementById('preview');

let selected = []; // {id, nombre, identidad}

function renderSelected(){
  empSelected.innerHTML = '';
  hiddenInputs.innerHTML = '';
  selected.forEach((e,i)=>{
    const li = document.createElement('li');
    li.innerHTML = `<div><strong>${i+1}.</strong> ${e.nombre} <span class="muted">(${e.identidad})</span></div>
                    <div class="del" data-id="${e.id}">Quitar</div>`;
    empSelected.appendChild(li);

    const h = document.createElement('input');
    h.type = 'hidden'; h.name = 'empleados[]'; h.value = e.id;
    hiddenInputs.appendChild(h);
  });
  bindDelete();
  renderPreview();
}
function bindDelete(){
  empSelected.querySelectorAll('.del').forEach(a=>{
    a.addEventListener('click', ev=>{
      const id = parseInt(ev.currentTarget.dataset.id,10);
      selected = selected.filter(x=>x.id!==id);
      renderSelected();
    });
  });
}

// PREVIEW: 15 por página
function headerHTML({tema,dur,inst,fecha,lugar,prog}){
  const isSI  = (prog||'').toUpperCase().trim()==='SI';
  const isNO  = !isSI;
  return `
    <div class="hdr">
      <div><img src="{{ asset('img/logo.PNG') }}" alt="LOGO"></div>
      <div class="hdr-center">
        <div class="l1">SERVICE AND TRADING BUSINESS S.A. DE C.V.</div>
        <div class="l2">PROCESO RECURSOS HUMANOS/ HUMAN RESOURCES PROCESS</div>
        <div class="l3">REGISTRO DE ASISTENCIA DE CAPACITACIÓN</div>
        <div class="l4"><em>REGISTER OF TRAINING ASSISTANCE</em></div>
      </div>
    </div>

    <div class="fname">Nombre de la Capacitación:
      <span class="line" style="min-width:500px">${tema || '&nbsp;'}</span>
    </div>

    <div class="meta">
      <div class="row"><label>Duración:</label> <span class="line">${dur? dur+' min':'&nbsp;'}</span></div>
      <div class="row"><label>Instructor:</label> <span class="line">${inst || '&nbsp;'}</span></div>
      <div class="row"><label>Fecha:</label> <span class="line">${fecha || '&nbsp;'}</span></div>
      <div class="row"><label>Lugar:</label> <span class="line">${lugar || '&nbsp;'}</span></div>
      <div class="row yn" style="grid-column:1 / span 2">
        <span>Programa en el plan de Capacitación:</span>
        <span>Si <span class="box ${isSI?'on':''}"></span></span>
        <span>No <span class="box ${isNO?'on':''}"></span></span>
      </div>
    </div>
  `;
}

function tableHTML(slice, offset){
  // 6 columnas EXACTAS del formato: N°, Nombre, N° Identidad, N° Teléfono, Puesto de Trabajo, Firma empleado
  const rows = slice.length ? slice : Array.from({length:15},()=>({}));
  return `
    <div class="table-wrap">
      <table class="tbl-r003">
        <thead>
          <tr>
            <th class="col-n">N°</th>
            <th>Nombre del Empleado</th>
            <th class="col-id">N° de Identidad</th>
            <th class="col-tel">N° de Teléfono</th>
            <th class="col-puesto">Puesto de Trabajo</th>
            <th class="col-firma">Firma empleado</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((r,idx)=>{
            const num = slice.length ? (offset + idx + 1) : '';
            const nom = r.nombre || '';
            const ide = r.identidad || '';
            // Teléfono y Puesto quedan en blanco (el cliente pidió ignorarlos en captura)
            return `
              <tr>
                <td class="col-n">${num}</td>
                <td>${nom || '&nbsp;'}</td>
                <td class="col-id">${ide || '&nbsp;'}</td>
                <td class="col-tel">&nbsp;</td>
                <td class="col-puesto">&nbsp;</td>
                <td class="col-firma">&nbsp;</td>
              </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function footerHTML(){
  return `
    <div class="foot">
      <div class="left">
        <span>1 copia Archivo</span>
        <span>1 copia sistema</span>
      </div>
      <div class="right">
        <span><strong>3 VERSION&nbsp;&nbsp;2017</strong></span>
        <span><strong>STB/RRHH/R003</strong></span>
      </div>
    </div>
  `;
}

function renderPreview(){
  const ciSel = document.getElementById('ci-select');
  const tema = ciSel.selectedOptions[0]?.dataset.tema || '';
  const dur  = ciSel.selectedOptions[0]?.dataset.duracion || '';
  const prog = ciSel.selectedOptions[0]?.dataset.programada || '';
  const inst = ciSel.selectedOptions[0]?.dataset.instructor || '';
  const fecha = document.getElementById('fecha_recibida').value || '';
  const lugar = document.getElementById('lugar_print').value || '';

  previewEl.innerHTML = '';
  const chunk = 15;
  const len = selected.length || 0;
  const pages = Math.max(1, Math.ceil(len / chunk));

  for (let p=0; p<pages; p++){
    const from = p*chunk;
    const to   = from + chunk;
    const slice = selected.slice(from, to);

    const page = document.createElement('div');
    page.className = 'a4';
    page.innerHTML = headerHTML({tema,dur,inst,fecha,lugar,prog}) + tableHTML(slice, from) + footerHTML();
    previewEl.appendChild(page);
  }
}

// ---- Autocompletar empleados
let srchTimer=null;
document.getElementById('emp-q').addEventListener('input', (ev)=>{
  const v = ev.target.value.trim();
  clearTimeout(srchTimer);
  if (!v){ empSuggest.style.display='none'; empSuggest.innerHTML=''; return; }
  srchTimer = setTimeout(async ()=>{
    const r = await fetch(routeEmpSrch + '?q=' + encodeURIComponent(v));
    const j = await r.json();
    empSuggest.innerHTML = '';
    (j.data||[]).forEach(e=>{
      const li = document.createElement('li');
      li.innerHTML = `<div>${e.nombre_completo} <span class="muted">(${e.identidad||'—'})</span></div>
                      <div class="badge">Agregar</div>`;
      li.style.cursor='pointer';
      li.addEventListener('click', ()=>{
        if (!selected.find(x=>x.id===e.id)){
          selected.push({id:e.id, nombre:e.nombre_completo, identidad:e.identidad});
          renderSelected();
        }
        empSuggest.style.display='none'; empSuggest.innerHTML=''; ev.target.value='';
      });
      empSuggest.appendChild(li);
    });
    empSuggest.style.display = (empSuggest.children.length>0) ? 'block' : 'none';
  }, 250);
});

// Limpiar lista
document.getElementById('clear-all').addEventListener('click', ()=>{
  selected=[]; renderSelected();
});

// Prellenar al cambiar capacitación
document.getElementById('ci-select').addEventListener('change', async (ev)=>{
  const opt = ev.target.selectedOptions[0];
  if (!opt) return;
  // Prefill desde data-* (rápido)
  document.getElementById('duracion_show').value   = (opt.dataset.duracion||'') ? (opt.dataset.duracion+' min') : '';
  document.getElementById('programada_show').value = opt.dataset.programada || '';
  document.getElementById('instructor_show').value = opt.dataset.instructor || '';
  // Sugerir instructor temporal con el nombre mostrado
  document.getElementById('instructor_temporal').value = opt.dataset.instructor || '';
  renderPreview();
});

// Propaga cambios a preview
['fecha_recibida','lugar_print'].forEach(id=>{
  document.getElementById(id).addEventListener('input', renderPreview);
});

// Guardar e imprimir
document.getElementById('save-print').addEventListener('click', ()=>{
  const f = document.getElementById('form-asistencia');
  let h = document.createElement('input');
  h.type='hidden'; h.name='print'; h.value='1';
  f.appendChild(h);
  f.submit();
});

// Imprimir automático tras guardar (si ?print=1)
@if(request()->boolean('print') && $ok)
  window.addEventListener('load', ()=> setTimeout(()=>window.print(), 400));
@endif

// Inicial
renderPreview();
</script>
@endsection
