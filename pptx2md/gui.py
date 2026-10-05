"""Local desktop window. Launch with python -m pptx2md.gui."""

from copy import deepcopy
import os
from pathlib import Path
import sys
from threading import Lock, Thread

from pptx2md.batch import BatchOptions, markdown_name, run_batch


class Api:
    def __init__(self):
        self._window = None
        self._lock = Lock()
        self._sources = []
        self._output_dir = (Path.cwd() / 'output').resolve()
        self._state = {'running': False, 'total': 0, 'processed': 0, 'succeeded': 0, 'failed': 0,
                       'current': None, 'items': [], 'combined_output': None, 'error': None,
                       'output_dir': str(self._output_dir)}
        self._busy = False

    def get_state(self):
        with self._lock:
            return dict(selection=[dict(name=p.name, path=str(p)) for p in self._sources],
                        output_dir=str(self._output_dir), batch=deepcopy(self._state), busy=self._busy)

    def add_files(self):
        import webview
        with self._lock:
            if self._busy:
                return {'error': 'La conversión está en curso.'}
            self._busy = True
        try:
            selected = self._window.create_file_dialog(
                webview.FileDialog.OPEN, allow_multiple=True,
                file_types=('Documentos compatibles (*.pptx;*.pdf)',))
            with self._lock:
                for name in selected or []:
                    path = Path(name).resolve()
                    if path.suffix.lower() in ('.pptx', '.pdf') and path not in self._sources:
                        self._sources.append(path)
        finally:
            with self._lock:
                self._busy = False
        return self.get_state()

    def remove_file(self, path):
        with self._lock:
            if self._busy:
                return {'error': 'La conversión está en curso.'}
            self._sources = [p for p in self._sources if str(p) != path]
        return self.get_state()

    def choose_output_dir(self):
        import webview
        with self._lock:
            if self._busy:
                return {'error': 'La conversión está en curso.'}
            self._busy = True
        try:
            selected = self._window.create_file_dialog(webview.FileDialog.FOLDER, directory=str(self._output_dir))
            if selected:
                with self._lock:
                    self._output_dir = Path(selected[0]).resolve()
        finally:
            with self._lock:
                self._busy = False
        return self.get_state()

    def start_conversion(self, settings):
        with self._lock:
            if self._busy:
                return {'error': 'Ya hay una operación en curso.'}
            if not self._sources:
                return {'error': 'Agrega al menos un archivo para convertir.'}
            mode = settings.get('mode', 'individual')
            if mode not in ('individual', 'combined'):
                return {'error': 'Selecciona un modo de salida válido.'}
            name = settings.get('combined_name', 'combinado.md')
            try:
                if mode == 'combined':
                    name = markdown_name(name)
            except ValueError as exc:
                return {'error': str(exc)}
            options = BatchOptions(self._output_dir, mode, name,
                                   bool(settings.get('include_images')), bool(settings.get('include_notes')))
            sources = list(self._sources)
            self._busy = True
            self._state = dict(running=True, total=len(sources), processed=0, succeeded=0, failed=0,
                               current=None, items=[], combined_output=None, error=None,
                               output_dir=str(self._output_dir))
        Thread(target=self._run, args=(sources, options), daemon=False).start()
        return self.get_state()

    def _run(self, sources, options):
        try:
            run_batch(sources, options, self._update)
        except Exception as exc:
            with self._lock:
                self._state.update(running=False, error=str(exc))
        finally:
            with self._lock:
                self._busy = False

    def _update(self, state):
        with self._lock:
            self._state = deepcopy(state)

    def open_result(self, path):
        with self._lock:
            allowed = {i['output'] for i in self._state['items'] if i.get('output')}
        if path not in allowed or not Path(path).is_file():
            return {'error': 'El resultado ya no está disponible.'}
        return self._open(Path(path))

    def open_output_dir(self):
        with self._lock:
            path = Path(self._state['output_dir'])
        if not path.is_dir():
            return {'error': 'La carpeta de salida todavía no existe.'}
        return self._open(path)

    @staticmethod
    def _open(path):
        try:
            if sys.platform == 'win32':
                os.startfile(str(path))
            else:
                import subprocess
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', str(path)])
            return {'ok': True}
        except OSError as exc:
            return {'error': f'No se pudo abrir: {exc}'}

    def _can_close(self):
        # No cancellation: finish the current batch before closing the window.
        with self._lock:
            return not self._busy


def main():
    try:
        import webview
    except ImportError:
        raise SystemExit('Instala la interfaz con: python -m pip install --editable ".[gui]"')
    api = Api()
    html = Path(__file__).with_name('gui.html').read_text(encoding='utf-8')
    window = webview.create_window('PPTX2MD · Conversor de documentos', html=html, js_api=api,
                                   width=1020, height=780, min_size=(720, 580), text_select=True,
                                   background_color='#f5f7fb')
    api._window = window
    window.events.closing += api._can_close
    try:
        webview.start(gui='edgechromium' if sys.platform == 'win32' else None)
    except Exception as exc:
        raise SystemExit(f'No se pudo iniciar la interfaz: {exc}. En Windows se requiere Microsoft Edge WebView2 Runtime.')


if __name__ == '__main__':
    main()
