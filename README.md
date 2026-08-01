# Mis Finanzas — Instructivo

Guía completa de esta carpeta. No hace falta saber nada de técnica para
seguirla — es un manual de "qué hacer" para el día a día.

---

## 1. Qué es cada cosa en esta carpeta

| Archivo / carpeta | Qué es |
|---|---|
| **Finanzas_Toto.xlsx** | El Excel real, con todos tus datos: gastos, ingresos, inversiones, metas. Todo lo demás se arma a partir de esto. |
| **Mi Dashboard de Finanzas.html** | El Dashboard visual (gráficos, balance, portafolio). Doble clic para abrirlo en el navegador. Es la versión de la compu. |
| **Ver mi Dashboard.bat** | Atajo para abrir el Dashboard sin buscar el archivo. |
| **Actualizar mi iPhone.bat** | El botón que más vas a usar. Lee el Excel, actualiza el Dashboard (compu **e** iPhone) y sube los cambios a Internet para que el celular los vea. |
| **Actualizar desde Bybit.bat** | Trae tu portafolio de cripto actualizado desde Bybit y lo carga al Excel + Dashboard. |
| **Cargar gastos de MercadoPago.bat** | Arrastrá el CSV que bajaste de MercadoPago y soltalo **arriba** de este archivo (no lo abras con doble clic). Carga los gastos nuevos al Excel y avisa si encuentra alguno repetido. |
| **Programas/** | Carpeta con los scripts técnicos que hacen funcionar todo esto. No hace falta entrar ahí nunca para el uso normal. |
| **docs/** | La copia del Dashboard que se publica a Internet (para el iPhone). Se actualiza sola cuando corrés "Actualizar mi iPhone.bat" — no la toques a mano. |
| **Version anterior (no editar)/** | Copia vieja del Dashboard, de antes del rediseño. Solo de recuerdo, no se usa. |

---

## 2. Uso del día a día

### Edité el Excel a mano (cargué un gasto, un ingreso, etc.)
1. Guardá el Excel y cerralo.
2. Doble clic en **"Actualizar mi iPhone.bat"**.
3. Listo — compu e iPhone quedan con los datos nuevos en un par de minutos.

### Bajé un CSV de MercadoPago
1. Arrastrá el archivo CSV y soltalo arriba de **"Cargar gastos de MercadoPago.bat"**.
2. Se abre una ventana negra que te cuenta qué cargó. Si hay una transferencia a una
   persona o comercio nuevo (que ninguna regla reconoce), **te va a preguntar ahí
   mismo** a qué categoría cargarla — escribila y Enter. Si apretás Enter sin escribir
   nada, la deja como "?" para que la revises después a mano en el Excel.
   La próxima vez que aparezca esa misma persona, ya se va a cargar sola (queda
   guardada la regla).
3. Si dice "posibles duplicados", andá al Excel (hoja **Gastos**) y fijate esas filas.
4. Ese mismo proceso ya deja todo actualizado — no hace falta correr nada más.

### Quiero actualizar mi portafolio de Bybit
1. Doble clic en **"Actualizar desde Bybit.bat"**.
2. Deja todo actualizado solo (Excel + Dashboard).

### Solo quiero mirar el Dashboard
Doble clic en **"Ver mi Dashboard.bat"** (o directamente en el archivo .html).

---

## 3. El Dashboard — qué tiene cada pestaña

- **Resumen**: balance del mes, gastos por categoría, ingresos vs. gastos.
- **Portafolio**: todo tu portafolio financiero (cripto, acciones, CEDEARs, bonos) con
  precios en vivo, y el gráfico **"Evolución del portafolio"** (ver punto 5).
- **Inversiones**: tus **metas de ahorro** (ej. juntar plata para un depto) y tu
  **mercadería para reventa** (ej. la comida para perro) — son cosas aparte del
  portafolio financiero.
- **Noticias**: noticias de cripto, mercados y Argentina, traducidas.

Arriba a la derecha hay tres botoncitos: actualizar precios, modo claro/oscuro, y
modo privacidad (oculta los montos con ••••, útil si alguien te mira la pantalla).

---

## 4. El iPhone

El Dashboard del celular es una "app" instalada desde el navegador (PWA), pide un
PIN para entrar (**1930**) — es solo para que no lo abra cualquiera que agarre el
celular, no es una seguridad real tipo banco.

**El celular NO se actualiza solo apenas editás el Excel.** Se actualiza recién
cuando corrés alguno de los `.bat` de esta carpeta (o cuando corre sola la tarea
de los viernes, ver punto 5) y eso sube los cambios a Internet. Puede tardar
1-2 minutos en aparecer en el teléfono después de eso.

Si el celular te muestra algo viejo después de esperar un rato: cerrá la app del
todo (deslizala hacia arriba en el selector de apps) y volvé a abrirla.

---

## 5. El gráfico "Evolución del portafolio"

Este gráfico muestra cómo cambió el valor de tu portafolio con el tiempo (podés
verlo por Día, Semana, Mes o Año). Para armarlo hace falta ir guardando un
registro del valor cada tanto — **eso pasa solo, no tenés que hacer nada.**

Hay una tarea programada en Windows (corre sola, no hace falta abrir nada) que
**todos los viernes a las 20:00** (con la compu prendida) calcula el valor de tu
portafolio con precios en vivo y lo guarda. Si la compu está apagada ese
momento, se pone al día apenas la prendés de nuevo.

Va a tardar unas semanas en juntar suficientes registros para mostrar una línea
completa — al principio vas a ver solo el último valor guardado, es normal.

---

## 6. Categorías y duplicados

- En la hoja **Gastos** del Excel hay una columna **"Posible duplicado"** (la
  última, con letra chica) que se pinta sola en rojo si dos filas tienen la
  misma fecha, moneda y monto — para que notes si algo se cargó dos veces por
  error (a mano o por el script de MercadoPago).
- La categoría **"Inversión"** es solo para plata que metés en cripto/acciones
  (Bull Market, compras de cripto). No mezcles ahí compras de mercadería para
  reventa — esas van aparte, en la sección de Inversiones Físicas del Excel.
- Si una transferencia queda con categoría **"?"**, es porque MercadoPago no supo
  a qué categoría ponerla sola — andá a esa fila y elegí una categoría a mano.

---

## 7. Si algo no anda

| Problema | Qué hacer |
|---|---|
| El Dashboard muestra datos viejos | Corré **"Actualizar mi iPhone.bat"**. Si ya lo corriste y sigue igual, hacé Ctrl+Shift+R en la pestaña del navegador (fuerza a que descarte lo guardado en caché). |
| El celular muestra datos viejos | Esperá 1-2 minutos después de correr el `.bat`, después cerrá la app del todo y volvé a abrirla. |
| Un script dice "Excel abierto, no pude leer" | Guardá y cerrá el Excel, y volvé a correr el `.bat`. |
| Un número no cierra / parece raro | Fijate la columna "Posible duplicado" en Gastos, y revisá que la categoría de esa fila sea la correcta. |
| No sé qué hace algo de esto | Preguntame — no toques nada de la carpeta **Programas** por las dudas. |

---

## 8. Qué NO tocar

- La carpeta **Programas/** (son los scripts que hacen funcionar todo — tocar algo
  ahí sin querer puede romper la actualización del Dashboard).
- La carpeta **docs/** (se genera sola).
- La carpeta **Version anterior (no editar)/**.
- Cualquier archivo que empiece con un punto (`.git`, `.gitignore`, `.vscode`, `.venv`)
  — son cosas técnicas internas.
