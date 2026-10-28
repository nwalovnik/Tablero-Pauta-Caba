"""Partida principal "Publicidad y propaganda" del presupuesto CABA, 2013 en adelante.

Usa los lectores del repo del tablero de presupuesto (build_navegador.py), que unifican los
esquemas de los CSV de ejecución de BA Data. Escribe data/presupuesto_publicidad.json:
por año sancionado, vigente y devengado (el último año: sanción y ejecución parcial),
y el devengado por jurisdicción.
"""
import os, sys, json, re, unicodedata, importlib.util
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_PRES = os.environ.get('REPO_PRESUPUESTO', os.path.join(HERE, '..', '..', '_repo_tablero_presupuesto'))
spec = importlib.util.spec_from_file_location('nav', os.path.join(REPO_PRES, 'data', 'build_navegador.py'))
nav = importlib.util.module_from_spec(spec)
spec.loader.exec_module(nav)


def es_publicidad(t):
    t = ''.join(c for c in unicodedata.normalize('NFD', t or '') if unicodedata.category(c) != 'Mn').lower()
    return 'publicidad' in t


def main():
    out = {'fuente': 'Ejecución presupuestaria GCBA (BA Data), partida principal 3.6 Publicidad y propaganda',
           # IPC promedio anual del tablero de presupuesto (IPCBA; el año en curso es estimado)
           'ipc_prom': {str(k): v for k, v in nav.bb.IPC_PROMEDIO_ANUAL.items()},
           'anios': {}}
    for y in nav.YEARS:
        rows, per = nav.load(y)
        tot = {'s': 0.0, 'v': 0.0, 'd': 0.0}
        jur = defaultdict(lambda: [0.0, 0.0])  # devengado, sancionado
        for r in rows:
            x = nav.rec(r, y)
            if not x or not es_publicidad(x['ppal_d']):
                continue
            for k in 'svd': tot[k] += x[k]
            j = nav.JUR_NOMBRE.get(x['jur']) or nav.pretty(x['jur_d'])
            jur[j][0] += x['d']; jur[j][1] += x['s']
        out['anios'][str(y)] = {'periodo': per, **{k: round(v) for k, v in tot.items()},
                                'jur': sorted([[k, round(v[0]), round(v[1])] for k, v in jur.items()], key=lambda z: -max(z[1], z[2]))}
        print(f'{y} ({per}): sanc {tot["s"]/1e6:10,.0f}  vig {tot["v"]/1e6:10,.0f}  dev {tot["d"]/1e6:10,.0f}  | {len(jur)} jurisdicciones', flush=True)
    with open(os.path.join(HERE, 'presupuesto_publicidad.json'), 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
