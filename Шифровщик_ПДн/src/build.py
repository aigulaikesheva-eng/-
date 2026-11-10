"""Сборка Шифровщик_ПДн.html: встраивает библиотеки в шаблон, чтобы файл работал офлайн.
Запуск: python build.py <путь к xlsx.full.min.js> <путь к jszip.min.js>"""
import sys, pathlib
here = pathlib.Path(__file__).parent
tpl = (here / 'shifr.template.html').read_text(encoding='utf-8')
xlsx = pathlib.Path(sys.argv[1]).read_text(encoding='utf-8')
jszip = pathlib.Path(sys.argv[2]).read_text(encoding='utf-8')
for lib in (xlsx, jszip):
    assert '</script' not in lib.lower()
out = tpl.replace('/*__XLSX__*/', xlsx, 1).replace('/*__JSZIP__*/', jszip, 1)
(here.parent / 'Шифровщик_ПДн.html').write_text(out, encoding='utf-8')
print('OK', len(out))
