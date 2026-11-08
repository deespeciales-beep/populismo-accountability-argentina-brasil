# Populismo y accountability horizontal

Un estudio comparado del ciclo progresista en Argentina y Brasil durante el gobierno de Néstor Kirchner y el primer gobierno de Lula da Silva (2003-2007).

Denisse Avaca de Eyto y Claudio Andrés Spinelli. Universidad Católica de La Plata, Facultad de Derecho y Ciencias Políticas, Licenciatura en Ciencia Política y Relaciones Internacionales. Trabajo final de la materia Sistemas Políticos Comparados, versión corregida después de la evaluación de la cátedra. La Plata, noviembre de 2026.

Este repositorio contiene el trabajo, el instrumento de relevamiento con todas las fichas, los documentos oficiales sobre los que se codificó y un programa que recalcula el índice. Cualquiera puede cambiar una decisión de codificación y ver qué pasa con el resultado.

## El problema

¿De qué modo incidió el estilo populista de ejercicio del poder ejecutivo sobre la accountability horizontal en los gobiernos nacionales de Argentina y Brasil entre 2003 y 2007?

La hipótesis sostenía que a mayor intensidad del estilo populista, menor accountability horizontal, y que la matriz de gobierno (mayoría propia o coalición) modula esa relación. La condición de refutación quedó fijada antes de medir: la hipótesis cae si los controles argentinos operaron pese a la mayoría propia, o si los brasileños fueron desactivados pese a la coalición.

## El diseño

Un índice de 100 puntos en dos mitades, con la matriz de gobierno fuera del índice como condición de lectura.

Populismo como estilo de ejercicio del poder, 50 puntos: A, estilo populista del discurso, 20 puntos en cuatro subdimensiones (apelación directa, antagonismo, conducción personal, deslegitimación de restricciones institucionales), codificadas sobre dos corpus; B, ejercicio del poder sobre los controles, 30 puntos en siete indicadores (leyes de delegación, decretos con fuerza de ley, salidas del máximo tribunal, reforma del órgano de gobierno judicial, vacancias en órganos de control, recortes a esos órganos, negativas de información).

Accountability horizontal, 50 puntos: C1 control legislativo (20), C2 control judicial (15), C3 auditoría externa (15), cada uno medido por oportunidades de control activadas y con efecto dentro del período.

Cada indicador se lleva a un nivel (nulo, bajo, medio, alto) por una regla fijada en el Cuadro 1 del trabajo, y el nivel se convierte en una fracción del peso (0, un tercio, dos tercios, uno). Los umbrales de lectura son: los controles operaron si C es 25 o más con al menos un efecto; dos valores son semejantes si difieren en 5 puntos o menos.

Corpus de discurso: primer corpus, cinco piezas por presidente (asunción y mensajes anuales ante el Congreso); segundo corpus, los discursos de los dos presidentes ante la Asamblea General de las Naciones Unidas, nueve piezas, incorporado después de la evaluación para controlar el género discursivo. Dieciocho piezas en total, todas con cita y página en el anexo.

## El resultado

Populismo: Argentina 41,5, Brasil 6,3 sobre 50. Accountability horizontal: Argentina 28,3, Brasil 30,0 sobre 50. Los controles argentinos operaron pese a la mayoría propia, y la hipótesis queda refutada en los términos en que fue formulada.

El resultado es sensible a la codificación de los efectos argentinos y el trabajo lo presenta con esa condición: si la remisión a la Fiscalía que el Congreso resolvió en 2005 no se cuenta como efecto, la accountability argentina baja a 21,7 y el resultado pasa a ser compatible con la hipótesis. El Cuadro 5 del trabajo exhibe esa y las demás decisiones discutibles, una por vez y en un escenario combinado. Lo que el relevamiento sí muestra con independencia de esas decisiones es por dónde pasó el control en cada país: por la auditoría externa y la comisión revisora de cuentas en Argentina; por las comisiones parlamentarias de investigación en Brasil.

## Qué hay en el repositorio

| Carpeta | Contenido |
| --- | --- |
| paper | El trabajo completo en PDF (25 páginas). |
| anexo | El anexo de instrumentos de relevamiento en PDF: fichas de corpus, episodios, oportunidades de control, matriz de gobierno e índice. |
| datos | El instrumento en Excel, con las fichas, las citas que justifican cada código y las decisiones de codificación fechadas. |
| corpus | La lista de las dieciocho piezas con fecha, fuente y enlace. |
| onu, discursos, normativa, congreso | Los documentos oficiales sobre los que se codificó: las actas de la Asamblea General de las Naciones Unidas, los discursos y mensajes presidenciales, las normas y el Diario de Sesiones. Son documentos públicos de los organismos que los emiten. |
| codigo | El programa que recalcula el índice desde el Excel. |
| referencias.md | La bibliografía completa con enlaces. Los libros y artículos con derechos de autor no se incluyen como archivo, sólo se enlazan. |

## Cómo recalcular

Requiere Python 3 y openpyxl.

```
pip install openpyxl
cd codigo
python calcular_indice.py
```

Imprime el Cuadro 4 (niveles y puntajes por indicador), el Cuadro 5 (sensibilidad) y la dimensión A por corpus. Para cambiar una decisión:

```
python calcular_indice.py --nivel AR:C1=bajo
python calcular_indice.py --nivel BR:B1=medio --nivel AR:B3=nulo
python calcular_indice.py --codigo BR-06:antagonismo=1
```

La primera orden quita el efecto a la remisión de 2005. La segunda cuenta las autorizaciones presupuestarias brasileñas y descuenta las salidas de la Corte argentina. La tercera codifica con 1 el antagonismo del discurso de Lula en la ONU de 2003, el caso de borde que el trabajo resuelve con 0.

Cualquier celda de nivel de la hoja Índice o de código de la hoja Corpus del Excel puede editarse directamente, y el programa toma lo que encuentre.

## Cómo citar

Avaca de Eyto, D. y Spinelli, C. A. (2026). Populismo y accountability horizontal. Un estudio comparado del ciclo progresista en Argentina y Brasil durante el gobierno de Néstor Kirchner y el primer gobierno de Lula da Silva (2003-2007). Trabajo final, Sistemas Políticos Comparados, Universidad Católica de La Plata.

## Licencia

El texto del trabajo, el anexo y los datos se publican bajo Creative Commons Atribución 4.0 Internacional (CC BY 4.0). El programa se publica bajo licencia MIT. Los documentos oficiales de las carpetas onu, discursos, normativa y congreso pertenecen a los organismos que los emiten y se reproducen como documentos públicos.
