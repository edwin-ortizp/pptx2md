"""Windows integration smoke: real hidden WebView2 and JS/Python bridge.

Run: python tests/gui_smoke.py
Dialog selections and opening external applications are stubbed; the window,
frontend controls, worker and conversions are real.
"""

import json
from pathlib import Path
import sys
import tempfile
import time
from unittest.mock import patch

import webview

from test_batch import BatchTests
from pptx2md import gui


def main():
    fixture = BatchTests()
    fixture.setUp()
    sources = [fixture.pptx(), fixture.pdf()]
    report = {}
    opened = []
    real_create = webview.create_window
    real_start = webview.start
    holder = {}

    def create(*args, **kwargs):
        kwargs['hidden'] = True
        window = real_create(*args, **kwargs)
        holder.update(window=window, api=kwargs['js_api'])
        return window

    def wait_for(predicate, seconds=15):
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            if predicate():
                return
            time.sleep(.1)
        raise AssertionError('Timed out waiting for the GUI')

    def exercise():
        window, api = holder['window'], holder['api']
        try:
            report['stage'] = 'load'
            if not window.events.loaded.wait(15):
                raise AssertionError('WebView2 did not load')
            wait_for(lambda: window.evaluate_js("typeof connected !== 'undefined' && connected"))
            report['stage'] = 'files'
            with patch.object(window, 'create_file_dialog', return_value=[str(p) for p in sources]) as dialog:
                window.run_js("document.getElementById('add').click()")
                wait_for(lambda: len(api.get_state()['selection']) == 2)
                wait_for(lambda: window.evaluate_js("document.getElementById('file-count').textContent === '2'"))
                assert dialog.call_args.kwargs['allow_multiple']
            wait_for(lambda: window.evaluate_js('!calling'))
            report['stage'] = 'folder'
            with patch.object(window, 'create_file_dialog', return_value=[str(fixture.out)]):
                window.run_js("document.getElementById('choose-folder').click()")
                wait_for(lambda: api.get_state()['output_dir'] == str(fixture.out))
                wait_for(lambda: window.evaluate_js("!calling && document.getElementById('choose-folder').disabled === false"))
            report['stage'] = 'convert'
            window.run_js("document.querySelector('[value=combined]').click(); document.getElementById('images').click(); document.getElementById('notes').click(); document.getElementById('convert').click()")
            wait_for(lambda: api.get_state()['batch']['total'] == 2)
            wait_for(lambda: not api.get_state()['busy'])
            report['stage'] = 'results'
            state = api.get_state()['batch']
            assert state['succeeded'] == 2, state
            output = Path(state['combined_output'])
            fixture.assert_images(output)
            assert 'Nota exclusiva' in output.read_text(encoding='utf-8')
            wait_for(lambda: window.evaluate_js("document.getElementById('counter').textContent === '2 de 2 procesados'"))
            wait_for(lambda: window.evaluate_js('!calling'))
            report['stage'] = 'open'
            with patch.object(api, '_open', side_effect=lambda path: opened.append(str(path)) or {'ok': True}):
                window.run_js("document.getElementById('open-combined').click()")
                wait_for(lambda: len(opened) == 1)
                wait_for(lambda: window.evaluate_js('!calling'))
                window.run_js("document.getElementById('open-folder').click()")
                wait_for(lambda: len(opened) == 2)
                wait_for(lambda: window.evaluate_js('!calling'))
            report.update(ok=True, engine='edgechromium', processed=state['processed'],
                          summary=window.evaluate_js("document.getElementById('summary').textContent"),
                          output_links_exist=True, image_links_exist=True, native_dialogs='stubbed',
                          open_actions=opened, overflow=window.evaluate_js('document.documentElement.scrollWidth > innerWidth'))
        except Exception as exc:
            report.update(ok=False, error=str(exc))
        finally:
            window.destroy()

    def start(*args, **kwargs):
        kwargs['func'] = exercise
        return real_start(*args, **kwargs)

    try:
        with patch.object(webview, 'create_window', side_effect=create), patch.object(webview, 'start', side_effect=start):
            gui.main()
    finally:
        fixture.doCleanups()
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report.get('ok') else 1


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
