# Copyright 2024 Liu Siyao
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import argparse
import logging
from pathlib import Path

from tqdm import tqdm

from pptx2md.entry import convert
from pptx2md.log import setup_logging
from pptx2md.types import ConversionConfig

setup_logging(compat_tqdm=True)
logger = logging.getLogger(__name__)


def parse_args():
    arg_parser = argparse.ArgumentParser(description='Convert pptx to markdown')
    arg_parser.add_argument('pptx_path',
                            type=Path,
                            nargs='?',
                            default=Path('.'),
                            help='path to the pptx file to be converted')
    arg_parser.add_argument('-t', '--title', type=Path, help='path to the custom title list file')
    arg_parser.add_argument('-o', '--output', type=Path, help='path of the output file')
    arg_parser.add_argument('-i', '--image-dir', type=Path, help='where to put images extracted')
    arg_parser.add_argument('--all', action="store_true", help='convert all pptx files in the target folder')
    arg_parser.add_argument('--image-width', type=int, help='maximum image with in px')
    arg_parser.add_argument('--disable-image', action="store_true", help='disable image extraction')
    arg_parser.add_argument('--disable-wmf',
                            action="store_true",
                            help='keep wmf formatted image untouched(avoid exceptions under linux)')
    arg_parser.add_argument('--disable-color', action="store_true", help='do not add color HTML tags')
    arg_parser.add_argument('--disable-escaping',
                            action="store_true",
                            help='do not attempt to escape special characters')
    arg_parser.add_argument('--disable-notes', action="store_true", help='do not add presenter notes')
    arg_parser.add_argument('--enable-slides', action="store_true", help='deliniate slides `\n---\n`')
    arg_parser.add_argument('--try-multi-column', action="store_true", help='try to detect multi-column slides')
    arg_parser.add_argument('--wiki', action="store_true", help='generate output as wikitext(TiddlyWiki)')
    arg_parser.add_argument('--mdk', action="store_true", help='generate output as madoko markdown')
    arg_parser.add_argument('--qmd', action="store_true", help='generate output as quarto markdown presentation')
    arg_parser.add_argument('--min-block-size',
                            type=int,
                            default=15,
                            help='the minimum character number of a text block to be converted')
    arg_parser.add_argument("--page", type=int, default=None, help="only convert the specified page")
    arg_parser.add_argument(
        "--keep-similar-titles",
        action="store_true",
        help="keep similar titles (allow for repeated slide titles - One or more - Add (cont.) to the title)")

    args = arg_parser.parse_args()

    return args


def get_output_path(args, pptx_path: Path) -> Path:
    extension = '.tid' if args.wiki else '.qmd' if args.qmd else '.md'
    if args.output is None:
        return pptx_path.with_suffix(extension)
    if args.all:
        return args.output / pptx_path.with_suffix(extension).name
    return args.output


def get_image_dir(args, output_path: Path) -> Path:
    return args.image_dir or output_path.parent / 'img'


def make_config(args, pptx_path: Path) -> ConversionConfig:
    output_path = get_output_path(args, pptx_path)
    return ConversionConfig(
        pptx_path=pptx_path,
        output_path=output_path,
        image_dir=get_image_dir(args, output_path),
        title_path=args.title,
        image_width=args.image_width,
        disable_image=args.disable_image,
        disable_wmf=args.disable_wmf,
        disable_color=args.disable_color,
        disable_escaping=args.disable_escaping,
        disable_notes=args.disable_notes,
        enable_slides=args.enable_slides,
        try_multi_column=args.try_multi_column,
        is_wiki=args.wiki,
        is_mdk=args.mdk,
        is_qmd=args.qmd,
        min_block_size=args.min_block_size,
        page=args.page,
        keep_similar_titles=args.keep_similar_titles,
    )


def main():
    args = parse_args()

    if args.all:
        target_dir = args.pptx_path
        if target_dir.is_file():
            target_dir = target_dir.parent
        pptx_files = sorted(target_dir.glob('*.pptx'))
        if not pptx_files:
            raise FileNotFoundError(f'no pptx files found in {target_dir}')
        for pptx_path in tqdm(pptx_files, desc='Converting files'):
            convert(make_config(args, pptx_path))
        return

    convert(make_config(args, args.pptx_path))


if __name__ == '__main__':
    main()
