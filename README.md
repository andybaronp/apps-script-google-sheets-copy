Este script copia el valor de una columna (origen) y lo pega en otra pestaña (destino.
-------------------------------------------------------
-------------------------------------------------------

Principalmente, se creó para asociar códigos a nombre en un Base de Datos. (Ver imagen).

## Modo de uso

- Cambiar los nombres de las pestañas, origen y destino
>const NONBRE_HOJA_PRINCIPAL = ""
const NONBRE_HOJA_DESTINO = ""

- Indicar el número de columna que se va a copiar en la pestaña NONBRE_HOJA_DESTINO
>// columna a copiar en BD codigo
const COLUMNA_ORIGEN = 4

- Indicar el número de columna que se va a copiar en la pestaña NONBRE_HOJA_PRINCIPAL
>// columna
const COLUMNA_DESTINO_RETORNO = 3

- Llamado inicial, cambiar la letra y número de la columna donde se va a pegar en NONBRE_HOJA_DESTINO, cambiar el nombre de la función a su gusto.

>// llamada inicial
function fb(){
ultiFilaCol("B",2)
}

- Menú para activar el script

// crea el menú lateral para activar

>function onOpen() {
var ui = SpreadsheetApp.getUi()
// cambiar el nombre del menú
>ui.createMenu("Generar código")
// nombre a mostrar y función a ejecutar
.addItem('FaceBook', 'fb')
.addToUi()
}




https://user-images.githubusercontent.com/46359791/163601909-eddf033d-b5b0-4c04-9ce4-53c11f4973e0.mp4

