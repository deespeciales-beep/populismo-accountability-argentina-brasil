"""Calcula el índice de populismo y accountability horizontal a partir del
instrumento de relevamiento (datos/Instrumentos de relevamiento.xlsx).

Reproduce el Cuadro 4 (niveles y puntajes) y el Cuadro 5 (sensibilidad) del
trabajo, y permite recalcular el índice cambiando el nivel de cualquier
indicador o el código de cualquier pieza del corpus.

Uso:
    python calcular_indice.py
    python calcular_indice.py --nivel AR:C1=bajo
    python calcular_indice.py --nivel BR:B1=medio --nivel AR:B3=nulo
    python calcular_indice.py --codigo BR-06:antagonismo=1

Requiere openpyxl (pip install openpyxl).
"""

import argparse
import os
import sys

import openpyxl

RUTA = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'datos',
                    'Instrumentos de relevamiento.xlsx')

FRACCION = {'nulo': 0.0, 'bajo': 1 / 3, 'medio': 2 / 3, 'alto': 1.0}

# Pesos de las subdimensiones de A, en el orden de las columnas de la hoja Corpus.
SUBDIM_A = ('apelacion', 'antagonismo', 'conduccion', 'deslegitimacion')
PESOS_A = {'apelacion': 4, 'antagonismo': 6, 'conduccion': 6, 'deslegitimacion': 4}


def f1(x):
    """Formato con coma decimal y un decimal."""
    return ('%.1f' % x).replace('.', ',')


# ---------------------------------------------------------------------------
# Hoja Índice: niveles por indicador
# ---------------------------------------------------------------------------

def leer_indice(wb):
    """Devuelve una lista de indicadores con peso y nivel por país."""
    ws = wb['Índice']
    filas = []
    for r in range(5, 19):
        ind = ws.cell(r, 3).value or ''
        codigo = ind.split('.')[0].strip() if ind[:2] in ('B1', 'B2', 'B3', 'B4',
                                                            'B5', 'B6', 'B7', 'C1',
                                                            'C2', 'C3') else None
        if codigo is None:
            sub = (ws.cell(r, 2).value or '').lower()
            codigo = {'apelación directa': 'A1', 'antagonismo': 'A2',
                      'conducción personal': 'A3'}.get(sub, 'A4')
        filas.append({
            'codigo': codigo,
            'nombre': ind if codigo[0] != 'A' else ws.cell(r, 2).value,
            'peso': float(ws.cell(r, 4).value),
            'AR': (ws.cell(r, 7).value or '').strip().lower(),
            'BR': (ws.cell(r, 10).value or '').strip().lower(),
        })
    return filas


def puntajes(filas, niveles_ar=None, niveles_br=None):
    """Suma A, B, populismo y C por país, con niveles opcionalmente sustituidos."""
    niveles_ar = niveles_ar or {}
    niveles_br = niveles_br or {}
    res = {}
    for pais, cambios in (('AR', niveles_ar), ('BR', niveles_br)):
        a = b = c = 0.0
        for fila in filas:
            nivel = cambios.get(fila['codigo'], fila[pais])
            if nivel not in FRACCION:
                raise ValueError('nivel desconocido %r en %s %s' % (nivel, pais, fila['codigo']))
            p = fila['peso'] * FRACCION[nivel]
            if fila['codigo'][0] == 'A':
                a += p
            elif fila['codigo'][0] == 'B':
                b += p
            else:
                c += p
        res[pais] = {'A': a, 'B': b, 'populismo': a + b, 'accountability': c}
    return res


# ---------------------------------------------------------------------------
# Hoja Corpus: códigos por pieza
# ---------------------------------------------------------------------------

def leer_corpus(wb):
    """Devuelve dos diccionarios {codigo_pieza: {pais, subdim: 0/1}} para los
    dos corpus (filas 5 a 14 y 30 a 38 de la hoja Corpus)."""
    ws = wb['Corpus']
    corpus = {1: {}, 2: {}}
    for num, filas in ((1, range(5, 15)), (2, range(30, 39))):
        for r in filas:
            cod = ws.cell(r, 1).value
            if not cod:
                continue
            corpus[num][cod] = {
                'pais': 'AR' if cod.startswith('AR') else 'BR',
                'apelacion': int(ws.cell(r, 7).value),
                'antagonismo': int(ws.cell(r, 8).value),
                'conduccion': int(ws.cell(r, 9).value),
                'deslegitimacion': int(ws.cell(r, 10).value),
            }
    return corpus


def nivel_cantidad(k):
    """Regla del primer corpus (Cuadro 1): 0 nulo; 1 bajo; 2 a 3 medio; 4 o más alto."""
    if k == 0:
        return 'nulo'
    if k == 1:
        return 'bajo'
    if k <= 3:
        return 'medio'
    return 'alto'


def nivel_proporcional(k, n):
    """Regla del segundo corpus y del agregado, con cortes equivalentes a los
    del primero: nulo sin piezas; bajo hasta un quinto; medio hasta tres
    quintos; alto más de tres quintos."""
    if k == 0:
        return 'nulo'
    f = k / n
    if f <= 0.2:
        return 'bajo'
    if f <= 0.6:
        return 'medio'
    return 'alto'


def dimension_a(piezas, regla):
    """Puntaje de A por país a partir de un conjunto de piezas codificadas."""
    out = {}
    for pais in ('AR', 'BR'):
        sel = [p for p in piezas.values() if p['pais'] == pais]
        n = len(sel)
        total = 0.0
        detalle = {}
        for s in SUBDIM_A:
            k = sum(p[s] for p in sel)
            nivel = regla(k) if regla is nivel_cantidad else regla(k, n)
            detalle[s] = (k, n, nivel)
            total += PESOS_A[s] * FRACCION[nivel]
        out[pais] = {'total': total, 'detalle': detalle}
    return out


# ---------------------------------------------------------------------------
# Cuadro 5
# ---------------------------------------------------------------------------

VARIANTES = [
    ('Reglas fijadas', {}, {}),
    ('La remisión resuelta por el Congreso en 2005 no se cuenta como efecto (C1 AR bajo)', {'C1': 'bajo'}, {}),
    ('La comunicación de la Auditoría al juzgado no se cuenta como efecto (C3 AR bajo)', {'C3': 'bajo'}, {}),
    ('El dictamen de 2006 sobre la resolución 82/06 cuenta con efecto (C1 AR alto)', {'C1': 'alto'}, {}),
    ('Cada salvedad del Tribunal de Cuentas se cuenta como un asunto (C3 BR alto)', {}, {'C3': 'alto'}),
    ('La sentencia del Supremo de agosto de 2007 se cuenta dentro del período (C2 BR medio)', {}, {'C2': 'medio'}),
    ('Las bancas del CNJ indicadas por las Cámaras se atribuyen al oficialismo (B4 BR medio)', {}, {'B4': 'medio'}),
    ('Las autorizaciones presupuestarias brasileñas se cuentan como delegación (B1 BR medio)', {}, {'B1': 'medio'}),
    ('El conflicto de competencias con la Procuración del Tesoro se cuenta en B7 (B7 AR alto)', {'B7': 'alto'}, {}),
]


def imprimir_cuadro4(filas, res):
    print('Cuadro 4. Niveles y puntajes por indicador')
    print('%-6s %-6s %-6s %-8s %-6s %-8s' % ('Ind.', 'Peso', 'AR', 'ptos', 'BR', 'ptos'))
    for fila in filas:
        par = fila['peso'] * FRACCION[fila['AR']]
        pbr = fila['peso'] * FRACCION[fila['BR']]
        print('%-6s %-6s %-6s %-8s %-6s %-8s' % (fila['codigo'], f1(fila['peso']), fila['AR'],
                                                f1(par), fila['BR'], f1(pbr)))
    for k, nombre in (('A', 'A (sobre 20)'), ('B', 'B (sobre 30)'),
                      ('populismo', 'Populismo (sobre 50)'),
                      ('accountability', 'Accountability horizontal (sobre 50)')):
        print('%-40s AR %6s   BR %6s' % (nombre, f1(res['AR'][k]), f1(res['BR'][k])))


def imprimir_cuadro5(filas, corpus):
    print()
    print('Cuadro 5. Sensibilidad del resultado a las decisiones de codificación')
    print('%-95s %6s %6s %6s %6s' % ('Variante', 'PopAR', 'AccAR', 'PopBR', 'AccBR'))
    base = puntajes(filas)
    for nombre, ar, br in VARIANTES:
        r = puntajes(filas, ar, br)
        print('%-95s %6s %6s %6s %6s' % (nombre, f1(r['AR']['populismo']), f1(r['AR']['accountability']),
                                          f1(r['BR']['populismo']), f1(r['BR']['accountability'])))
    # Variante del corpus agregado
    agregado = dimension_a({**corpus[1], **corpus[2]}, nivel_proporcional)
    pop_ar = base['AR']['B'] + agregado['AR']['total']
    pop_br = base['BR']['B'] + agregado['BR']['total']
    print('%-95s %6s %6s %6s %6s' % ('Los dos corpus agregados con la regla proporcional (A AR %s; A BR %s)'
                                      % (f1(agregado['AR']['total']), f1(agregado['BR']['total'])),
                                      f1(pop_ar), f1(base['AR']['accountability']),
                                      f1(pop_br), f1(base['BR']['accountability'])))
    comb_ar = puntajes(filas, {'B3': 'nulo'}, {})['AR']['B'] + agregado['AR']['total']
    comb_br = puntajes(filas, {}, {'B1': 'medio'})['BR']['B'] + agregado['BR']['total']
    print('%-95s %6s %6s %6s %6s' % ('Escenario combinado: corpus agregado, B1 BR contado y B3 AR descontado',
                                      f1(comb_ar), f1(base['AR']['accountability']),
                                      f1(comb_br), f1(base['BR']['accountability'])))


def imprimir_corpus(corpus):
    print()
    print('Dimensión A por corpus')
    for num, regla, nombre in ((1, nivel_cantidad, 'Primer corpus (regla de cantidad fija)'),
                               (2, nivel_proporcional, 'Segundo corpus, ONU (regla proporcional)')):
        d = dimension_a(corpus[num], regla)
        print(nombre)
        for pais in ('AR', 'BR'):
            det = '; '.join('%s %d/%d %s' % (s, k, n, niv) for s, (k, n, niv) in d[pais]['detalle'].items())
            print('  %s  A = %s   (%s)' % (pais, f1(d[pais]['total']), det))
    d = dimension_a({**corpus[1], **corpus[2]}, nivel_proporcional)
    print('Agregado de los dos corpus (regla proporcional)')
    for pais in ('AR', 'BR'):
        det = '; '.join('%s %d/%d %s' % (s, k, n, niv) for s, (k, n, niv) in d[pais]['detalle'].items())
        print('  %s  A = %s   (%s)' % (pais, f1(d[pais]['total']), det))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--archivo', default=RUTA, help='ruta del Excel de instrumentos')
    ap.add_argument('--nivel', action='append', default=[],
                    help='sustituye un nivel: PAIS:INDICADOR=nivel, por ejemplo AR:C1=bajo')
    ap.add_argument('--codigo', action='append', default=[],
                    help='sustituye el código de una pieza: PIEZA:subdimension=0|1, por ejemplo BR-06:antagonismo=1')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.archivo, data_only=True)
    filas = leer_indice(wb)
    corpus = leer_corpus(wb)

    cambios = {'AR': {}, 'BR': {}}
    for item in args.nivel:
        pais, resto = item.split(':', 1)
        ind, nivel = resto.split('=')
        cambios[pais.upper()][ind.upper()] = nivel.strip().lower()
    for item in args.codigo:
        pieza, resto = item.split(':', 1)
        sub, val = resto.split('=')
        for num in (1, 2):
            if pieza.upper() in corpus[num]:
                corpus[num][pieza.upper()][sub.strip().lower()] = int(val)

    if cambios['AR'] or cambios['BR']:
        for pais in ('AR', 'BR'):
            for fila in filas:
                if fila['codigo'] in cambios[pais]:
                    fila[pais] = cambios[pais][fila['codigo']]
        print('Niveles sustituidos:', cambios)
        print()

    res = puntajes(filas)
    imprimir_cuadro4(filas, res)
    imprimir_cuadro5(filas, corpus)
    imprimir_corpus(corpus)


if __name__ == '__main__':
    main()
