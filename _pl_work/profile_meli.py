from pathlib import Path
from time import perf_counter
import os, sys, pandas as pd
log=Path(r"C:\Users\PauloMendonça\OneDrive - Redefrete\Documentos\DRE\_pl_work\profile_meli.log")
def say(msg):
    with log.open('a', encoding='utf-8') as f:
        f.write(msg+'\n')
        f.flush()
say('start')
sys.path.insert(0, str(Path(r"C:\Users\PauloMendonça\OneDrive - Redefrete\Documentos\DRE") / '_pl_work'))
say('before import script')
import build_meli_month as b
say('after import script')
folder=Path(r"C:\Users\PauloMendonça\OneDrive - Redefrete\Área de Trabalho\Balanço\DashBoard_P&L\Meli\022026")
say(f'folder {folder.exists()} csv={len(list(folder.glob("*.csv")))} xlsx={len(list(folder.glob("*.xlsx")))}')
t=perf_counter(); csv_frames=[]
for p in sorted(folder.glob('*.csv')):
    say('read csv '+p.name)
    csv_frames.append(b.read_meli_csv(p))
say(f'csv done {perf_counter()-t:.2f} rows={sum(len(x) for x in csv_frames)}')
t=perf_counter(); csv_base=b.enrich_csv_rows(pd.concat(csv_frames, ignore_index=True)); say(f'enrich done {perf_counter()-t:.2f} rows={len(csv_base)}')
t=perf_counter();
for p in sorted(folder.glob('*.xlsx')):
    if p.name.startswith('~$') or p.name.startswith('Base_Meli_Consolidada_'):
        say('skip xlsx '+p.name)
        continue
    say('test open '+p.name)
    xl=pd.ExcelFile(p, engine='openpyxl')
    say('sheets '+p.name+' '+str(xl.sheet_names))
say(f'xlsx open all {perf_counter()-t:.2f}')
t=perf_counter(); rotas,carga=b.read_operational_files(folder); say(f'read operational {perf_counter()-t:.2f} rotas={len(rotas)} carga={len(carga)}')
say('done')
