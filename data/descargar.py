"""Baja todos los recursos del dataset Pauta Publicitaria de BA Data a data/raw/."""
import os, re, sys, html, time, urllib.parse
import requests

BASE = 'https://data.buenosaires.gob.ar'
DS = BASE + '/dataset/pauta-publicitaria'
RAW = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'raw')
UA = {'User-Agent': 'Mozilla/5.0'}


def recursos():
    s = requests.get(DS, headers=UA, timeout=60).text
    out = []
    for m in re.finditer(r'resource/([0-9a-f-]{36})/download".*?<h3>(.*?)</h3>', s, flags=re.S):
        out.append((m.group(1), html.unescape(re.sub(r'\s+', ' ', re.sub('<[^>]+>', '', m.group(2))).strip())))
    return out


def main():
    os.makedirs(RAW, exist_ok=True)
    rs = recursos()
    print(len(rs), 'recursos')
    with open(os.path.join(RAW, '_recursos.tsv'), 'w', encoding='utf-8') as f:
        for rid, t in rs:
            f.write(f'{rid}\t{t}\n')
    ya = {f[:8] for f in os.listdir(RAW)}
    for rid, t in rs:
        if rid[:8] in ya:
            continue
        for intento in range(5):
            try:
                r = requests.get(f'{DS}/resource/{rid}/download', headers=UA, timeout=120, allow_redirects=True)
                break
            except requests.RequestException as e:
                print('  reintento', intento + 1, t, e.__class__.__name__, flush=True)
                time.sleep(5 * (intento + 1))
        else:
            print('  FALLÓ', t); continue
        name = urllib.parse.unquote(r.url.rstrip('/').split('/')[-1]) or rid
        cd = r.headers.get('content-disposition', '')
        m = re.search(r'filename="?([^";]+)', cd)
        if m: name = m.group(1)
        if name == 'download' or '.' not in name: name = rid + '.bin'
        path = os.path.join(RAW, rid[:8] + '_' + name)
        with open(path, 'wb') as f:
            f.write(r.content)
        print(r.status_code, len(r.content), t, '->', os.path.basename(path), flush=True)
        time.sleep(1)


if __name__ == '__main__':
    main()
