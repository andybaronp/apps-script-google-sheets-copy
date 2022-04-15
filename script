
const NONBRE_HOJA_PRINCIPAL = "AGREGAR NOMBRE DE HOJA"
const NONBRE_HOJA_DESTINO = "AGREGAR NOMBRE DE HOJA"
// columna a copiar en BD codigo
const COLUMNA_ORIGEN = 4
// columna 
const COLUMNA_DESTINO_RETORNO = 3

// ---copiar nombre -----

function copiarnombre(number,numberCol) {

  let archivo = SpreadsheetApp.getActiveSpreadsheet() 
  let hojaOrigen = archivo.getActiveSheet()
  let nombreOrigen =hojaOrigen.getName()
  let hojaDestino = archivo.getSheetByName(NONBRE_HOJA_DESTINO)
  let celdaActiva =hojaOrigen.getActiveCell();
  let filaActiva = celdaActiva.getRow()

  //Define rango para copiar
  //el COLUMNA_ORIGEN es el numero de columna que copio
  let rangoOrigen = hojaOrigen.getRange(filaActiva, COLUMNA_ORIGEN,1,1);
  let rangoDestino = hojaDestino.getRange( number, numberCol,1,1)

  //Define rango para copiar de vuelta
  //el COLUMNA_DESTINO_RETORNO es el numero de columna que copio VER  getRange como funciona
  let rangoDestino2 = hojaOrigen.getRange(filaActiva, COLUMNA_DESTINO_RETORNO,1,1);
  let rangoOrigen2 = hojaDestino.getRange( number,rangoDestino.getColumn()-1 ,1,1)
   
  if(nombreOrigen === NONBRE_HOJA_PRINCIPAL ){
    rangoOrigen.copyTo(rangoDestino)
    rangoOrigen2.copyTo(rangoDestino2)
  }
     // activa las hojas para copiar
    rangoDestino.activate()
    rangoDestino2.activate()
  if(nombreOrigen === NONBRE_HOJA_PRINCIPAL){
    // Alerta con el nombre y código asociado
    Browser.msgBox("*Código generado* ",rangoDestino2.getValue() +" "+ rangoOrigen.getValue(),Browser.Buttons.OK )
  }
}  

 
// Calcula la última fila en blanco -  recibe la columna en letra y en numero
function ultiFilaCol(letraCol,numberCol){
  const columna = letraCol
  /////////////////
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NONBRE_HOJA_DESTINO)
  let maxfilasB = ss.getMaxRows()

  let valores = ss.getRange(columna+"1:"+columna+maxfilasB).getValues()
  
  for(; valores[maxfilasB-1]==""& maxfilasB > 0;maxfilasB--){
  }
  let res = parseInt(maxfilasB+1).toFixed(0)
   copiarnombre(res,numberCol)
}



// llamada inicial
function fb(){
  ultiFilaCol("B",2)
}
 
// llamada inicial
function referido(){
  ultiFilaCol("E",5)
}

// Crea el menú
function onOpen() {
  var ui = SpreadsheetApp.getUi()
    ui.createMenu("Generar código")
    .addItem('FaceBook', 'fb')
    .addItem('Referido', 'referido')

    .addToUi()
}
