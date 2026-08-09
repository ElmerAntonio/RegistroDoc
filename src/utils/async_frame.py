"""
utils/async_frame.py
────────────────────
Mixin liviano para cargar datos de BD en un hilo de fondo
y aplicar los resultados al UI en el hilo principal de Tkinter.

Patrón de uso en cualquier Frame:
    class MiFrame(AsyncFrameMixin, ctk.CTkFrame):
        def actualizar_vista(self):
            self._cargar_async(self._fetch_datos, self._apply_datos)

        def _fetch_datos(self):
            # Se ejecuta en hilo de fondo — SOLO queries SQL, sin tocar widgets
            return self.engine.obtener_algo()

        def _apply_datos(self, datos):
            # Se ejecuta en el hilo principal — actualiza widgets
            self.mi_label.configure(text=datos[...])
"""
import threading
import logging

logger = logging.getLogger("async_frame")


class AsyncFrameMixin:
    """
    Mixin que agrega capacidad de carga asíncrona a cualquier CTkFrame.
    Thread-safe: usa un token por operación para descartar resultados obsoletos.
    """

    def _cargar_async(self, fetch_fn, apply_fn, on_error=None):
        """
        Ejecuta fetch_fn en un hilo de fondo.
        Cuando termina, llama apply_fn(resultado) en el hilo principal.

        Args:
            fetch_fn  : callable sin argumentos — hace las queries SQL
            apply_fn  : callable(datos) — actualiza widgets con el resultado
            on_error  : callable(exc) opcional — maneja errores
        """
        # Token único para esta operación; cancela actualizaciones obsoletas
        token = object()
        self._async_token = token

        def _run():
            try:
                datos = fetch_fn()
            except Exception as exc:
                logger.warning(f"[async_frame] Error en fetch: {exc}")
                datos = None
                if on_error:
                    try:
                        self.after(0, lambda e=exc: on_error(e))
                    except Exception:
                        pass
                return

            def _apply_safe():
                # Descartar si ya se lanzó otra carga más reciente
                if getattr(self, "_async_token", None) is not token:
                    return
                # Descartar si el widget fue destruido (navegación rápida)
                try:
                    if not self.winfo_exists():
                        return
                except Exception:
                    return
                try:
                    apply_fn(datos)
                except Exception as exc2:
                    logger.error(f"[async_frame] Error en apply: {exc2}")

            try:
                self.after(0, _apply_safe)
            except Exception:
                pass  # Widget destruido antes del after()

        threading.Thread(target=_run, daemon=True).start()
