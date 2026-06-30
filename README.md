# Project SelfMonitoring — Iker Martin Saladich

Portafoli personal amb sistema PLM integrat per al seguiment de projectes.

## Web pública

[iker-martin-saladich.github.io/PAP](https://iker-martin-saladich.github.io/PAP/)

## Estructura

```
site/
├── index.html                    ← Pàgina principal
    index_future.html             ← Pàgina principal futura
├── actualitzar.py                ← Llegeix els Excel i actualitza els HTML
├── web_personal/
│   └── plm_web.html
├── tfm_quantica/
│   ├── plm_tesi.html
│   └── PLM_TFM_Quantica.xlsx
├── elevador_plats/
│   ├── plm_elevador.html
│   └── PLM_Elevador_Plats.xlsx
└── frances/
    └── plm_frances.html
```

## Flux de treball

1. Modifica qualsevol Excel amb el progrés actualitzat
2. `git add . && git commit -m "Actualitzar progrés" && git push`
3. GitHub Actions executa `actualitzar.py` automàticament
4. La web pública s'actualitza en ~1 minut

## Ús local

```bash
pip install openpyxl
python actualitzar.py
```

## Web pública

[iker-martin-saladich.github.io/Project-SelfMonitoring](https://iker-martin-saladich.github.io/Project-SelfMonitoring/)
