"""Sequential GUI batches, using the public converter without a GUI dependency."""

from collections import Counter
from dataclasses import dataclass
import os
from pathlib import Path
import re
import shutil
import tempfile

from pptx2md.entry import convert
from pptx2md.types import ConversionConfig


@dataclass(frozen=True)
class BatchOptions:
    output_dir: Path
    mode: str = 'individual'
    combined_name: str = 'combinado.md'
    include_images: bool = False
    include_notes: bool = False


def markdown_name(value):
    value = value.strip()
    if (not value or value in ('.', '..') or re.search(r'[<>:"/\\|?*\x00-\x1f]', value)
            or value.endswith('.') or re.fullmatch(r'(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])',
                                                  value.split('.')[0], re.I)):
        raise ValueError('Escribe un nombre de archivo válido, sin carpetas ni caracteres especiales.')
    if not Path(value).suffix:
        value += '.md'
    if Path(value).suffix.lower() != '.md':
        raise ValueError('El nombre del documento combinado debe terminar en .md.')
    return value


def unique_path(folder, name, used, asset_root=None):
    original = Path(name)
    candidate = folder / original.name
    number = 2
    while (str(candidate).casefold() in used or candidate.exists()
           or (asset_root is not None and (asset_root / candidate.stem).exists())):
        candidate = folder / f'{original.stem}_{number}{original.suffix}'
        number += 1
    used.add(str(candidate).casefold())
    return candidate


def _publish(temporary, destination):
    # Exclusive creation protects existing files even if another process creates
    # the planned destination after preflight. Never remove somebody else's file.
    with destination.open('x', encoding='utf-8') as target:
        try:
            target.write(temporary.read_text(encoding='utf-8'))
        except BaseException:
            target.close()
            destination.unlink(missing_ok=True)
            raise


def _cleanup_assets(path, output_dir):
    if path is not None and path.exists():
        resolved = path.resolve()
        expected_root = (output_dir / 'img').resolve()
        if resolved != expected_root and expected_root in resolved.parents:
            shutil.rmtree(resolved)


def _error_message(exc):
    text = str(exc)
    if 'password protected' in text:
        return 'El PDF está protegido con contraseña. Utiliza una copia sin protección.'
    if 'no extractable text' in text:
        match = re.search(r'page (\d+)', text)
        page = f' (página {match[1]})' if match else ''
        return f'El PDF contiene una página sin texto extraíble{page}. Esta versión no incluye OCR.'
    if isinstance(exc, PermissionError):
        return 'No se pudo leer el archivo o escribir la salida. Revisa los permisos y si está en uso.'
    if isinstance(exc, FileNotFoundError):
        return 'El archivo de origen ya no está disponible.'
    if 'could not read PDF' in text or 'Package not found' in text or 'BadZipFile' in type(exc).__name__:
        return 'No se pudo leer el documento. Comprueba que el archivo no esté dañado.'
    return text or 'No se pudo convertir el documento.'


def combine_section(name, markdown):
    # The converter emits ATX headings. Respect fenced blocks when shifting them.
    lines = []
    fence = None
    for line in markdown.splitlines():
        marker = re.match(r'^\s*(`{3,}|~{3,})', line)
        if marker:
            token = marker[1]
            if fence is None:
                fence = token
            elif token[0] == fence[0] and len(token) >= len(fence):
                fence = None
        elif fence is None:
            line = re.sub(r'^(#{1,6})(\s+)', lambda m: '#' * min(6, len(m[1]) + 1) + m[2], line)
        lines.append(line)
    escaped_name = re.sub(r'([\\`*_{}\[\]<>#!])', r'\\\1', name)
    return f'# {escaped_name}\n\n' + '\n'.join(lines).strip() + '\n\n'


def run_batch(sources, options, on_update=None):
    """Return JSON-ready results; callback receives snapshots throughout a run."""
    if options.mode not in ('individual', 'combined'):
        raise ValueError('Selecciona un modo de salida válido.')
    sources = list(dict.fromkeys(Path(p).resolve() for p in sources))
    if not sources:
        raise ValueError('Agrega al menos un archivo para convertir.')
    combined_name = markdown_name(options.combined_name) if options.mode == 'combined' else None
    output_dir = options.output_dir.resolve()
    state = dict(running=True, total=len(sources), processed=0, succeeded=0, failed=0,
                 current=None, output_dir=str(output_dir), combined_output=None, error=None,
                 items=[dict(name=p.name, path=str(p), status='pending', output=None, error=None)
                        for p in sources])

    def update():
        if on_update:
            # Never expose live dictionaries while the worker updates them.
            from copy import deepcopy
            on_update(deepcopy(state))

    update()
    used = set()
    combined_temp = None
    combined_assets = None
    owns_combined_assets = False
    sections = []
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        if options.mode == 'combined':
            combined_output = unique_path(output_dir, combined_name, used, output_dir / 'img')
            combined_assets = output_dir / 'img' / combined_output.stem
            if options.include_images and any(p.suffix.lower() == '.pptx' for p in sources):
                combined_assets.mkdir(parents=True, exist_ok=False)
                owns_combined_assets = True
            fd, temp_name = tempfile.mkstemp(prefix='.pptx2md-', suffix='.md', dir=output_dir)
            os.close(fd)
            combined_temp = Path(temp_name)
        counts = Counter(p.stem.casefold() for p in sources)
        for index, source in enumerate(sources):
            item = state['items'][index]
            item['status'] = 'converting'
            state['current'] = source.name
            update()
            temporary = None
            asset_dir = None
            owns_assets = False
            try:
                if source.suffix.lower() not in ('.pptx', '.pdf'):
                    raise ValueError('Formato no admitido. Selecciona archivos .pptx o .pdf; los .ppt antiguos no son compatibles.')
                name = (source.name if counts[source.stem.casefold()] > 1 else source.stem) + '.md'
                folder = output_dir if options.mode == 'individual' else combined_assets
                destination = unique_path(folder, name, used, output_dir / 'img' if options.mode == 'individual' else None)
                if options.include_images and source.suffix.lower() == '.pptx':
                    asset_dir = (output_dir / 'img' / destination.stem if options.mode == 'individual'
                                 else combined_assets / destination.stem)
                    asset_dir.mkdir(parents=True, exist_ok=False)
                    owns_assets = True
                fd, temp_name = tempfile.mkstemp(prefix='.pptx2md-', suffix='.md', dir=output_dir)
                os.close(fd)
                temporary = Path(temp_name)
                convert(ConversionConfig(pptx_path=source, output_path=temporary, image_dir=asset_dir,
                                         disable_image=not options.include_images,
                                         disable_notes=not options.include_notes, disable_color=True))
                if options.mode == 'combined':
                    sections.append(combine_section(source.name, temporary.read_text(encoding='utf-8')))
                else:
                    _publish(temporary, destination)
                    item['output'] = str(destination)
                item['status'] = 'completed'
                state['succeeded'] += 1
            except Exception as exc:
                if owns_assets:
                    _cleanup_assets(asset_dir, output_dir)
                item['status'] = 'error'
                item['error'] = _error_message(exc)
                state['failed'] += 1
            finally:
                if temporary is not None:
                    temporary.unlink(missing_ok=True)
            state['processed'] += 1
            update()
        if combined_temp is not None and state['succeeded']:
            combined_temp.write_text(''.join(sections), encoding='utf-8')
            _publish(combined_temp, combined_output)
            state['combined_output'] = str(combined_output)
            for item in state['items']:
                if item['status'] == 'completed':
                    item['output'] = str(combined_output)
        elif owns_combined_assets:
            _cleanup_assets(combined_assets, output_dir)
    except Exception as exc:
        state['error'] = _error_message(exc)
        if owns_combined_assets:
            # Successful entries are only usable once the combined file is saved.
            _cleanup_assets(combined_assets, output_dir)
        for item in state['items']:
            if item['status'] in ('pending', 'converting') or (options.mode == 'combined' and item['status'] == 'completed'):
                item.update(status='error', error=state['error'], output=None)
        state['succeeded'] = sum(i['status'] == 'completed' for i in state['items'])
        state['failed'] = sum(i['status'] == 'error' for i in state['items'])
        state['processed'] = state['total']
    finally:
        if combined_temp is not None:
            combined_temp.unlink(missing_ok=True)
        state['running'] = False
        state['current'] = None
        update()
    return state
