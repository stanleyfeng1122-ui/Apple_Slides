#!/usr/bin/env python3
"""Screenshot every page of a local slide deck with installed, offline Chrome."""

import sys
sys.dont_write_bytecode = True

import argparse
import glass
from pathlib import Path
import re
from urllib.parse import urlparse

from playwright.sync_api import sync_playwright


def render(html: Path, output: Path, chrome: Path) -> int:
    html = html.expanduser().resolve(strict=True)
    chrome = chrome.expanduser().resolve(strict=True)
    output = output.expanduser().resolve()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            executable_path=str(chrome), headless=True,
            args=['--disable-background-networking', '--disable-component-update',
                  '--no-first-run'],
        )
        try:
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080}, device_scale_factor=1,
                offline=True, service_workers='block',
            )
            context.route('**/*', lambda route: route.abort()
                          if urlparse(route.request.url).scheme in ('http', 'https', 'ws', 'wss')
                          else route.continue_())
            context.add_init_script("window.GLASS_DEFER_BUILD = true")
            page = context.new_page()
            page.goto(html.as_uri(), wait_until='load')
            page.evaluate('document.fonts.ready')
            glass.prepare(page)
            pages = page.locator('.page')
            ids = pages.evaluate_all('(pages) => pages.map(p => p.id)')
            if not ids:
                raise ValueError('No .page elements found')
            if len(ids) != len(set(ids)) or any(
                not re.fullmatch(r'[A-Za-z0-9_-]+', page_id) for page_id in ids
            ):
                raise ValueError('Each page needs a unique filename-safe id')
            output.mkdir(parents=True, exist_ok=True)
            for page_id in ids:
                page.evaluate('(id) => showPage(id)', page_id)
                glass.calibrate(page, page.locator(f'[id="{page_id}"]'))
                glass.screenshot(page).save(output / f'{page_id}.png')
        finally:
            browser.close()
    return len(ids)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('html', type=Path)
    parser.add_argument('output', type=Path, help='Directory for one PNG per page ID')
    parser.add_argument('--chrome', type=Path, default=Path(
        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'))
    args = parser.parse_args()
    count = render(args.html, args.output, args.chrome)
    print(f'Rendered {count} HTML pages (1920×1080): {args.output}')


if __name__ == '__main__':
    main()
