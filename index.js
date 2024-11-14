const XlsxPopulate = require('xlsx-populate');
const path = require('path');

const arrayDeMeses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

const armarObjetos = (alumnos, montos, instituciones, niveles, meses, años) => {
    const arrayDeDatos = [];

    for(let i = 0; i < alumnos.length; i++) {
        const datos = {
            alumno: alumnos[i],
            monto: parseInt(montos[i]),
            institucion: instituciones[i],
            nivel: niveles[i],
            mes: parseInt(meses[i]),
            año: años[i],
        }

        arrayDeDatos.push(datos);
    }

    return arrayDeDatos;
}


const analizarDatos = async () => {
    try {
        const workbook = await XlsxPopulate.fromFileAsync(path.resolve(__dirname, './Input/morosidad.xlsx'));

        const hoja = workbook.sheet(0);
        const datos = hoja.usedRange(); // Necesario para contar las filas
        const cantFilas = datos.endCell().rowNumber();  // Recuento de filas con datos

        // Datos en crudo
        const dniAlumnos = hoja.range(`E2:E${cantFilas}`).value();
        const montosTotales = hoja.range(`I2:I${cantFilas}`).value();
        const instituciones = hoja.range(`AA2:AA${cantFilas}`).value();
        const nivelesEducativos = hoja.range(`AB2:AB${cantFilas}`).value();
        const meses = hoja.range(`AG2:AG${cantFilas}`).value();
        const años = hoja.range(`AH2:AH${cantFilas}`).value();
        // --------------------------------------------------------------

        const todosLosDatos = armarObjetos(dniAlumnos, montosTotales, instituciones, nivelesEducativos, meses, años);
        
        const resultado = todosLosDatos.reduce((acc, obj) => {
            // Crear una clave única basada en las propiedades que determinan la agrupación
            const clave = `${obj.institucion}-${obj.nivel}-${obj.mes}-${obj.año}`;
          
            // Si la clave ya existe en el acumulador, actualizamos la cantidad de alumnos y sumamos el monto
            if (acc[clave]) {
              acc[clave].alumno += 1;
              acc[clave].monto += obj.monto;
            } else {
              // Si no existe, creamos un nuevo objeto con la clave
              acc[clave] = {
                institucion: obj.institucion,
                nivel: obj.nivel,
                mes: obj.mes,
                año: obj.año,
                alumno: 1,
                monto: obj.monto
              };
            }
          
            return acc;
        }, {});
          
        // Convertir el objeto acumulado en un array de resultados
        const resultadoArray = Object.values(resultado);
        

        
        // -------------------  CREAR NUEVO EXCEL  -------------------

        const workbookNew = await XlsxPopulate.fromBlankAsync();

        const hojaNew = workbookNew.sheet(0);

        const cantAlumnosNew = resultadoArray.map(obj => obj.alumno);
        const institucionesNew = resultadoArray.map(obj => obj.institucion);
        const montosNew = resultadoArray.map(obj => obj.monto);
        const nivelesNew = resultadoArray.map(obj => obj.nivel);
        const mesesNew = resultadoArray.map(obj => obj.mes);
        const mesesStringNew = mesesNew.map(mes => arrayDeMeses[mes - 1]);
        const añosNew = resultadoArray.map(obj => obj.año);

        hojaNew.column('C').style({ numberFormat: 'General'});
        hojaNew.column('D').style({ numberFormat: '_($* #,##0.00_);_($* (#,##0.00);_($* '-'??_);_(@_)'});

        hojaNew.range('A1:F1').value([['Unidad Educativa', 'Nivel Educativo', 'Cant. Alumnos', 'Monto', 'Mes', 'Año']]);
        hojaNew.range(`A2:A${institucionesNew.length}`).value(institucionesNew);
        hojaNew.range(`B2:B${nivelesNew.length}`).value(nivelesNew);
        hojaNew.range(`C2:C${cantAlumnosNew.length}`).value(cantAlumnosNew.map(cant => [cant]));
        hojaNew.range(`D2:D${montosNew.length}`).value(montosNew.map(monto => [monto]));
        hojaNew.range(`E2:E${mesesStringNew.length}`).value(mesesStringNew.map(mes => [mes]));
        hojaNew.range(`F2:F${añosNew.length}`).value(añosNew);

        await workbookNew.toFileAsync(path.relative(__dirname, './Output/datosMorosidad.xlsx'));
    } catch(error) {
        console.log(error);
    }
}

analizarDatos();