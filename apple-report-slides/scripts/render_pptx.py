#!/usr/bin/env python3
"""Render local PPTX slides to PNG; native app first, LibreOffice fallback."""

import sys
sys.dont_write_bytecode = True

import argparse
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

import pymupdf
from PIL import Image, ImageOps


UNAVAILABLE_APPS = {}


def _native_export(source: Path, pdf: Path, app: str) -> None:
    # Paths are argv, and the temporary filename is unique to this render.
    if app == 'powerpoint':
        running = subprocess.run(
            ['osascript', '-e', 'application "Microsoft PowerPoint" is running'],
            check=True, capture_output=True, text=True, timeout=10).stdout.strip() == 'true'
        quit_handler = '''on quitIfUnused(wasRunning)
  if not wasRunning then
    tell application "Microsoft PowerPoint"
      with timeout of 8 seconds
        if (count of presentations) is 0 then
          try
            quit
          on error
            error "PowerPoint is showing a dialog; close it by hand"
          end try
        end if
      end timeout
    end tell
  end if
end quitIfUnused
'''
        script = quit_handler + '''on run argv
set sourceFile to POSIX file (item 1 of argv)
set outputFile to POSIX file (item 2 of argv)
set temporaryName to item 3 of argv
set wasRunning to (item 4 of argv is "true")
tell application "Microsoft PowerPoint"
  try
    with timeout of 90 seconds
      open sourceFile
      delay 3
      set d to active presentation
      if name of d is not temporaryName then error "Active presentation is not the temporary deck"
      save d in outputFile as save as PDF
      close d saving no
    end timeout
  on error errorText number errorNumber
    try
      with timeout of 8 seconds
        repeat with i from (count of presentations) to 1 by -1
          set candidate to presentation i
          if name of candidate is temporaryName then close candidate saving no
        end repeat
      end timeout
    end try
    my quitIfUnused(wasRunning)
    error errorText number errorNumber
  end try
  my quitIfUnused(wasRunning)
end tell
end run'''
        cleanup = quit_handler + '''on run argv
set temporaryName to item 3 of argv
set wasRunning to (item 4 of argv is "true")
with timeout of 8 seconds
  tell application "Microsoft PowerPoint"
    repeat with i from (count of presentations) to 1 by -1
      set candidate to presentation i
      if name of candidate is temporaryName then close candidate saving no
    end repeat
  end tell
end timeout
my quitIfUnused(wasRunning)
end run'''
    else:
        script = '''on run argv
set sourceFile to POSIX file (item 1 of argv)
set outputFile to POSIX file (item 2 of argv)
set openedDeck to missing value
tell application "Keynote"
  try
    with timeout of 100 seconds
      set openedDeck to open sourceFile
      export openedDeck to outputFile as PDF
    end timeout
    with timeout of 8 seconds
      close openedDeck saving no
    end timeout
  on error errorText number errorNumber
    if openedDeck is not missing value then
      try
        with timeout of 8 seconds
          close openedDeck saving no
        end timeout
      end try
    end if
    error errorText number errorNumber
  end try
end tell
end run'''
        cleanup = '''on run argv
set sourcePath to item 1 of argv
with timeout of 8 seconds
  tell application "Keynote"
    repeat with candidate in documents
      if (POSIX path of (file of candidate)) is sourcePath then
        close candidate saving no
        exit repeat
      end if
    end repeat
  end tell
end timeout
end run'''
    command = ['osascript', '-', str(source), str(pdf)]
    if app == 'powerpoint':
        command.extend([source.name, str(running).lower()])
    try:
        # Reserve ten seconds for exact-document cleanup inside the 120s budget.
        subprocess.run(command, input=script, check=True, capture_output=True,
                       text=True, timeout=100)
    except subprocess.TimeoutExpired as exc:
        try:
            subprocess.run(command, input=cleanup, check=True, capture_output=True,
                           text=True, timeout=10)
        except (OSError, subprocess.SubprocessError) as cleanup_error:
            if app == 'powerpoint':
                print('PowerPoint is showing a dialog; close it by hand', flush=True)
            raise RuntimeError('Native export timed out; temporary document close unverified: '
                               + str(cleanup_error)) from exc
        raise RuntimeError('Native export timed out after 100s; temporary document cleanup completed') from exc
    if not pdf.is_file():
        raise RuntimeError(f'{app} completed without producing a preview')


def native_export(source: Path, pdf: Path, app: str) -> None:
    if app != 'powerpoint':
        _native_export(source, pdf, app)
        return
    container_tmp = Path.home() / 'Library/Containers/com.microsoft.Powerpoint/Data/tmp'
    container_tmp.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='slides-render-', dir=container_tmp) as tmp:
        work = Path(tmp)
        local_source = work / source.name
        local_pdf = local_source.with_suffix('.pdf')
        shutil.copyfile(source, local_source)
        _native_export(local_source, local_pdf, app)
        shutil.move(str(local_pdf), str(pdf))


def libreoffice_export(source: Path, work: Path) -> Path:
    soffice = shutil.which('soffice')
    if not soffice:
        candidate = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        if candidate.is_file():
            soffice = str(candidate)
    if not soffice:
        raise RuntimeError('LibreOffice unavailable on PATH or in /Applications/LibreOffice.app')
    env = os.environ.copy()
    env['XDG_CACHE_HOME'] = str(work / 'cache')
    executable = Path(soffice).resolve()
    font_configs = [executable.parent.parent / 'Resources/fontconfig/fonts.conf']
    if len(executable.parents) >= 3:
        font_configs.append(executable.parents[2] / 'native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fontconfig/fonts.conf')
    for config in font_configs:
        if config.is_file():
            env['FONTCONFIG_FILE'] = str(config)
            break
    destination = work / 'libreoffice'
    destination.mkdir()
    export = 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}'
    subprocess.run([soffice, f'-env:UserInstallation={(work / "profile").as_uri()}',
                    '--headless', '--convert-to', export, '--outdir', str(destination), str(source)],
                   check=True, capture_output=True, text=True, timeout=120, env=env)
    preview = destination / (source.stem + '.pdf')
    if not preview.is_file():
        raise RuntimeError('LibreOffice completed without producing a preview')
    return preview


def render(pptx: Path, output: Path, app: str = 'powerpoint') -> tuple[int, str]:
    pptx = pptx.expanduser().resolve(strict=True)
    output = output.expanduser().resolve()
    with zipfile.ZipFile(pptx) as deck:
        root = ET.fromstring(deck.read('ppt/presentation.xml'))
    count = len(root.findall('{*}sldIdLst/{*}sldId'))
    if not count:
        raise ValueError('The deck contains no slides')
    expected = {f'S{number:02d}.png' for number in range(1, count + 1)}
    stale = sorted(path.name for path in output.glob('S[0-9]*.png')
                   if path.stem[1:].isdigit() and path.name not in expected)
    if stale:
        raise ValueError('Output contains images outside this deck: ' + ', '.join(stale))
    renderer = {'powerpoint':'PowerPoint', 'keynote':'Keynote', 'libreoffice':'LibreOffice'}[app]
    with tempfile.TemporaryDirectory(prefix='slides-render-') as tmp:
        work = Path(tmp)
        source = work / f'{work.name}.pptx'
        shutil.copyfile(pptx, source)
        pdf = source.with_suffix('.pdf')
        if app != 'libreoffice' and app not in UNAVAILABLE_APPS:
            try:
                native_export(source, pdf, app)
            except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
                reason = (getattr(exc, 'stderr', None) or str(exc)).strip()
                UNAVAILABLE_APPS[app] = reason
                print(f'{renderer} unavailable: {reason}\nFalling back to LibreOffice.', flush=True)
                renderer = 'LibreOffice'
        elif app in UNAVAILABLE_APPS:
            renderer = 'LibreOffice'
        if renderer == 'LibreOffice':
            pdf = libreoffice_export(source, work)
        with pymupdf.open(pdf) as document:
            if len(document) != count:
                raise RuntimeError(f'{renderer} exported {len(document)} pages for {count} slides')
            output.mkdir(parents=True, exist_ok=True)
            for number, page in enumerate(document, 1):
                scale = min(1920/page.rect.width, 1080/page.rect.height)
                pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
                raster = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                fitted = ImageOps.contain(raster, (1920, 1080), Image.Resampling.LANCZOS)
                canvas = Image.new('RGB', (1920, 1080), 'white')
                canvas.paste(fitted, ((1920-fitted.width)//2, (1080-fitted.height)//2))
                canvas.save(output / f'S{number:02d}.png')
        print(f'Renderer: {renderer} · {count} slides · {output}', flush=True)
    return count, renderer


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--app', choices=['powerpoint', 'keynote', 'libreoffice'], default='powerpoint')
    args = parser.parse_args()
    try:
        render(args.pptx, args.output, args.app)
    except (OSError, ValueError, RuntimeError, KeyError, zipfile.BadZipFile,
            ET.ParseError, subprocess.SubprocessError) as exc:
        detail = getattr(exc, 'stderr', None) or str(exc)
        parser.exit(1, f'Render failed: {detail}\n')


if __name__ == '__main__':
    main()
