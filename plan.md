1. **Fix hardcoded UI labels in `src/happ.py`**:
    - Ensure Spanish labels for export buttons. Update labels "Exportar to PDF" and "Exportar to Word" to "Exportar a PDF" and "Exportar a Word" if they exist. *Already verified they are correctly named `Exportar a PDF` and `Exportar a Word`, nothing to do here.*
2. **Fix hardcoded grade list in `src/fapp.py` and `src/obsapp.py`**:
    - Update hardcoded grade lists to use `self.engine.obtener_grados_activos() or ["Sin datos"]`.
3. **Fix Validation Logic**:
    - The validation logic is mostly correct but verify its usages inside `eapp.py` and `rddata.py` to ensure it falls back gracefully rather than using `1.0` or `0.0` as fallbacks for invalid entries.
4. **Fix caching logic in `src/rddata.py`**:
    - In `actualizar_resumen` method, `should_close` and cache update are performed. We need to ensure that `self._cargar_en_memoria()` and `.close()` are explicitly gated behind `if should_close`.
    - Also, inside `obtener_grados_activos`, it should explicitly return `[]` without falling back to hardcoded strings when the file doesn't exist or is empty.
