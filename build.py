"""Rebuild index.html from src/ (page head, page body, fixer engine). Run: python3 build.py"""
import pathlib
root = pathlib.Path(__file__).parent
head, body, fixer = [(root / 'src' / n).read_text() for n in ('page_head.html', 'page_body.html', 'fixer.js')]
assert '/*FIXER*/' in body
page = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'
        + head + '</head>\n<body>\n' + body.replace('/*FIXER*/', fixer) + '\n</body>\n</html>\n')
(root / 'index.html').write_text(page)
print('index.html rebuilt')
