<div class="relative overflow-x-auto overflow-y-auto max-h-[100vh] shadow-md sm:rounded-lg" style="margin-top: 20px;">
    <table class="w-full text-sm text-left text-gray-500 border-separate border-spacing-0">
        <thead class="sticky top-0 z-20 bg-gray-50">
            <tr>
                <th scope="col" class="px-6 py-3">
                    No. Cuota
                </th>
                <th scope="col" class="px-6 py-3">
                    No. Préstamo
                </th>
                <th scope="col" class="px-6 py-3">
                    Nombre del Empleado
                </th>
                <th scope="col" class="px-6 py-3">
                    Fecha de Préstamo
                </th>
                <th scope="col" class="px-6 py-3">
                    Fecha de Deducción
                </th>
                <th scope="col" class="px-6 py-3">
                    Abono a Capital
                </th>
                <th scope="col" class="px-6 py-3">
                    Abono a Intereses
                </th>
                <th scope="col" class="px-6 py-3">
                    Cuota Mensual
                </th>
                <th scope="col" class="px-6 py-3">
                    Cuota Quincenal
                </th>
                <th scope="col" class="px-6 py-3">
                    Estado
                </th>
                <th scope="col" class="px-6 py-3">Observaciones</th>
                <th scope="col" class="px-6 py-3">Acciones</th>
            </tr>
        </thead>
        <tbody>
        @forelse ($cuotas as $cuota)
            @php
                $isPagado = (int) ($cuota->pagado ?? 0) === 1;
                $rowClass = $isPagado ? 'bg-red-50' : 'bg-white';
                $pagadoText = $isPagado ? 'PAGADO' : 'PENDIENTE';
                $pagadoCellClass = $isPagado ? 'bg-red-200 text-red-900 font-semibold' : '';
            @endphp

            <tr class="{{ $rowClass }} border-b border-gray-200 hover:bg-gray-50">
                <td class="px-6 py-4 bg-blue-100 font-semibold">
                    {{ $cuota->num_cuota }}
                </td>
                <td class="px-6 py-4">
                    {{ $cuota->num_prestamo }}
                </td>
                <td class="px-6 py-4">
                    {{ $cuota->nombre_completo }}
                </td>
                <td class="px-6 py-4">
                    {{ $cuota->fecha_deposito_prestamo }}
                </td>
                <td class="px-6 py-4" style="background-color: #90D4D0">
                    {{ $cuota->fecha_programada }}
                </td>
                <td class="px-6 py-4" style="background-color: #D8ABE0">
                    {{ $cuota->abono_capital }}
                </td>
                <td class="px-6 py-4" style="background-color: #D8ABE0">
                    {{ $cuota->abono_intereses }}
                </td>
                <td class="px-6 py-4" style="background-color: #E3C598">
                    {{ $cuota->cuota_mensual }}
                </td>
                <td class="px-6 py-4">
                    {{ $cuota->cuota_quincenal }}
                </td>
                <td class="px-6 py-4 {{ $pagadoCellClass }}">
                    {{ $pagadoText }}
                </td>

                <td class="px-6 py-4 {{ (stripos($cuota->observaciones ?? '', 'cobro extraordinario') !== false) ? 'bg-yellow-100 text-yellow-900 font-medium' : ((stripos($cuota->observaciones ?? '', 'dep�sito') !== false || stripos($cuota->observaciones ?? '', 'deposito') !== false) ? 'bg-green-100 text-green-900 font-medium' : '') }}">
                    {{ $cuota->observaciones }}
                </td>
                <td class="px-6 py-4">
                    <button type="button"
                            data-modal-target="edit-cuota-{{ $cuota->id_historial_cuotas }}"
                            data-modal-toggle="edit-cuota-{{ $cuota->id_historial_cuotas }}"
                            class="text-blue-600 hover:underline">Editar</button>

                    <div id="edit-cuota-{{ $cuota->id_historial_cuotas }}" tabindex="-1" aria-hidden="true" class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
                        <div class="relative p-4 w-full max-w-md max-h-full">
                            <div class="relative bg-white rounded-lg shadow-sm">
                                <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
                                    <h3 class="text-lg font-semibold text-gray-900">Editar cuota</h3>
                                    <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-toggle="edit-cuota-{{ $cuota->id_historial_cuotas }}">
                                        <svg class="w-3 h-3" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 14"><path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"/></svg>
                                        <span class="sr-only">Close modal</span>
                                    </button>
                                </div>
                                <form action="{{ route('cuotas.update', $cuota->id_historial_cuotas) }}" method="POST" class="p-4 md:p-5">
                                    @csrf
                                    @method('PUT')
                                    <div class="grid gap-4 mb-4 grid-cols-2">
                                        <div class="col-span-2">
                                            <label class="block mb-2 text-sm font-medium text-gray-900">Nombre</label>
                                            <input type="text" value="{{ $cuota->nombre_completo }}" class="w-full bg-gray-100 border border-gray-200 text-gray-700 text-sm rounded-lg p-2.5" disabled>
                                        </div>
                                        <div class="col-span-2">
                                            <label class="block mb-2 text-sm font-medium text-gray-900">No. Prestamo</label>
                                            <input type="text" value="{{ $cuota->num_prestamo }}" class="w-full bg-gray-100 border border-gray-200 text-gray-700 text-sm rounded-lg p-2.5" disabled>
                                        </div>
                                        <div>
                                            <label for="num_cuota_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">No. Cuota</label>
                                            <input type="number" id="num_cuota_{{ $cuota->id_historial_cuotas }}" name="num_cuota" value="{{ old('num_cuota', $cuota->num_cuota) }}" min="0" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5" required>
                                        </div>
                                        <div>
                                            <label for="fecha_programada_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Fecha programada</label>
                                            <input type="date" id="fecha_programada_{{ $cuota->id_historial_cuotas }}" name="fecha_programada" value="{{ old('fecha_programada', $cuota->fecha_programada ? \Illuminate\Support\Carbon::parse($cuota->fecha_programada)->format('Y-m-d') : '') }}" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5" required>
                                        </div>
                                        <div>
                                            <label for="abono_capital_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Abono a capital</label>
                                            <input type="number" step="0.01" min="0" id="abono_capital_{{ $cuota->id_historial_cuotas }}" name="abono_capital" value="{{ old('abono_capital', $cuota->abono_capital) }}" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5" required>
                                        </div>
                                        <div>
                                            <label for="abono_intereses_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Abono a intereses</label>
                                            <input type="number" step="0.01" min="0" id="abono_intereses_{{ $cuota->id_historial_cuotas }}" name="abono_intereses" value="{{ old('abono_intereses', $cuota->abono_intereses) }}" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5" required>
                                        </div>
                                        <div>
                                            <label for="cuota_mensual_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Cuota mensual</label>
                                            <input type="number" step="0.01" min="0" id="cuota_mensual_{{ $cuota->id_historial_cuotas }}" name="cuota_mensual" value="{{ old('cuota_mensual', $cuota->cuota_mensual) }}" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5" required>
                                        </div>
                                        <div>
                                            <label for="cuota_quincenal_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Cuota quincenal</label>
                                            <input type="number" step="0.01" min="0" id="cuota_quincenal_{{ $cuota->id_historial_cuotas }}" name="cuota_quincenal" value="{{ old('cuota_quincenal', $cuota->cuota_quincenal) }}" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5" required>
                                        </div>
                                        <div class="col-span-2">
                                            <label for="observaciones_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Observaciones</label>
                                            <textarea id="observaciones_{{ $cuota->id_historial_cuotas }}" name="observaciones" rows="3" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5">{{ old('observaciones', $cuota->observaciones) }}</textarea>
                                        </div>
                                        <div class="col-span-2">
                                            <label for="estado_{{ $cuota->id_historial_cuotas }}" class="block mb-2 text-sm font-medium text-gray-900">Estado</label>
                                            <select id="estado_{{ $cuota->id_historial_cuotas }}" name="estado" class="w-full bg-white border border-gray-300 text-gray-900 text-sm rounded-lg p-2.5">
                                                <option value="0" {{ (int) ($cuota->pagado ?? 0) === 0 ? 'selected' : '' }}>PENDIENTE</option>
                                                <option value="1" {{ (int) ($cuota->pagado ?? 0) === 1 ? 'selected' : '' }}>PAGADO</option>
                                            </select>
                                        </div>
                                    </div>
                                    <button type="submit" class="text-white inline-flex items-center bg-blue-700 hover:bg-blue-800 focus:ring-4 focus:outline-none focus:ring-blue-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center">Guardar</button>
                                </form>
                            </div>
                        </div>
                    </div>

                    <button type="button"
                        data-modal-target="delete-cuota-{{ $cuota->id_historial_cuotas }}"
                        data-modal-toggle="delete-cuota-{{ $cuota->id_historial_cuotas }}"
                        class="text-red-600 hover:underline ml-3">Eliminar</button>

                    <div id="delete-cuota-{{ $cuota->id_historial_cuotas }}" tabindex="-1" aria-hidden="true" class="hidden overflow-y-auto overflow-x-hidden fixed top-0 right-0 left-0 z-50 justify-center items-center w-full md:inset-0 h-[calc(100%-1rem)] max-h-full">
                        <div class="relative p-4 w-full max-w-md max-h-full">
                            <div class="relative bg-white rounded-lg shadow-sm">
                                <div class="flex items-center justify-between p-4 md:p-5 border-b rounded-t border-gray-200">
                                    <h3 class="text-lg font-semibold text-gray-900">Eliminar cuota</h3>
                                    <button type="button" class="text-gray-400 bg-transparent hover:bg-gray-200 hover:text-gray-900 rounded-lg text-sm w-8 h-8 ms-auto inline-flex justify-center items-center" data-modal-toggle="delete-cuota-{{ $cuota->id_historial_cuotas }}">
                                        <svg class="w-3 h-3" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 14 14"><path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"/></svg>
                                        <span class="sr-only">Close modal</span>
                                    </button>
                                </div>
                                <div class="p-4 md:p-5 text-center">
                                    <svg class="mx-auto mb-4 text-gray-400 w-12 h-12" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 20 20">
                                        <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 11V6m0 8h.01M19 10a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z"/>
                                    </svg>
                                    <h3 class="mb-5 text-lg font-normal text-gray-500">�Est� seguro de eliminar esta cuota?</h3>
                                    <form action="{{ route('cuotas.destroy', $cuota->id_historial_cuotas) }}" method="POST">
                                        @csrf
                                        @method('DELETE')
                                        <button type="submit" class="text-white bg-red-600 hover:bg-red-800 focus:ring-4 focus:outline-none focus:ring-red-300 font-medium rounded-lg text-sm inline-flex items-center px-5 py-2.5 text-center">
                                            S�, eliminar
                                        </button>
                                        <button data-modal-hide="delete-cuota-{{ $cuota->id_historial_cuotas }}" type="button" class="py-2.5 px-5 ms-3 text-sm font-medium text-gray-900 focus:outline-none bg-white rounded-lg border border-gray-200 hover:bg-gray-100 hover:text-blue-700 focus:z-10 focus:ring-4 focus:ring-gray-100">
                                            No, cancelar
                                        </button>
                                    </form>
                                </div>
                            </div>
                        </div>
                    </div>
                </td>
            </tr>
        @empty
            <tr>
                <td colspan="12" class="px-6 py-4 text-center text-sm text-gray-500">No hay cuotas para mostrar.</td>
            </tr>
        @endforelse
        </tbody>
    </table>
    @if($cuotas instanceof \Illuminate\Contracts\Pagination\Paginator)
        {{ $cuotas->links() }}
    @endif
</div>

