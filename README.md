# Monitor de precios

Si tienes un negocio y cada semana revisas a mano si los precios de una página subieron o bajaron, este programa lo hace por ti: guarda los precios, los compara con la ejecución anterior y te manda un correo solo cuando algo cambió.

## Qué hace

1. Lee una página web y extrae los precios (scraper).
2. Limpia los datos con pandas.
3. Guarda un CSV que sirve de referencia para la siguiente ejecución.
4. Compara contra el CSV anterior para detectar cambios.
5. Si hay cambios, genera un reporte Excel y lo envía por correo. Si no, solo imprime "no hubo cambios" en la terminal.

## Demostración

El sitio de práctica books.toscrape.com no cambia sus precios. Para forzar un cambio, modifiqué a mano el precio de 3 libros en el CSV de referencia. En la segunda ejecución el programa detectó la diferencia y generó `reporte2026-11-01 18-15-45.xlsx`.

## Cómo ejecutarlo

Probado en Windows con PowerShell y Python 3.14.

1. Clona el repositorio.
2. Crea el entorno virtual `.venv` e instala los paquetes: `requests`, `beautifulsoup4`, `pandas`, `numpy`, `openpyxl` y `python-dotenv`.
3. Crea un archivo `.env` en `src/` con tres variables: `EMAIL_SENDER`, `PASSWORD` y `EMAIL_RECIVER` (se escriben así en el código).
   Ese archivo guarda la contraseña del correo, por eso está en `.gitignore`: si se subiera a GitHub, cualquiera podría leerla.
4. Ejecuta `parser.py`.

## Automatización

Una tarea programada de Windows ejecuta el programa todos los domingos a las 18:00.

- Requiere la PC encendida y la sesión iniciada.
- Si la PC estaba apagada a esa hora, hay que activar en la configuración de la tarea "Ejecutar la tarea lo antes posible si se omitió un inicio programado". Con eso, la tarea corre al encender el equipo.
- Si no hubo cambios de precios, el programa corre y termina sin errores, pero no genera Excel ni envía correo.
- El envío de correo se probó lanzando la tarea manualmente desde el Programador de tareas, no en una ejecución automática del domingo.

## Limitación conocida

Si un libro aparece o desaparece entre ejecuciones, el programa lo reporta como un cambio de precio (de 0 a X, o de X a 0), aunque ningún precio haya cambiado. Pasa porque el `merge` outer deja `NaN` donde no hay coincidencia y `fillna(0)` lo convierte en 0.
### en otras palabras:
Como el precio no estaba registrado el fillna(0)  mostrará el libro nuevo como que subio de precio, ya que antes de ese nuevo precio lo compara con 0.