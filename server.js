
const excelToJson = require('convert-excel-to-json');
var ignoreCase = require('ignore-case');
var json2xls = require('json2xls');
const fs = require('fs');


const  eliminarDiacriticos = (texto)=>  {
    return texto.normalize('NFD').replace(/[\u0300-\u036f]/g,"");
}
const result = excelToJson({
    sourceFile: './excel/xl_employee_20210204_1221.xlsx'
});

const result2 = excelToJson({
    sourceFile: './excel/Usuarios_Con_Correo_Pegaduro.xlsx'
});


const empleados = []
let contador = 0;
result.employee.map(empleadoPegaduro => {
    const nombreCompleto = `${empleadoPegaduro.A}${empleadoPegaduro.B}${empleadoPegaduro.C}`.replaceAll(' ','')
    const nombreFiltroPegaduro = eliminarDiacriticos(nombreCompleto)
    let correoRene = ''
    result2.Hoja1.map((empleadoRene,index)=>{

        const nombreCompletoRene = `${empleadoRene.A}${empleadoRene.B}${empleadoRene.C}`.replaceAll(' ','')
        const nombreFiltroRene = eliminarDiacriticos(nombreCompletoRene)
     
         if(nombreFiltroPegaduro===nombreFiltroRene)
         {
            correoRene =empleadoRene.D?empleadoRene.D:'';
         }
        
    })

    
    const empleado = {
        NombreEmpleado:empleadoPegaduro.A?empleadoPegaduro.A:'',
        ApellidoPaterno:empleadoPegaduro.B?empleadoPegaduro.B:'',
        ApellidoMaterno:empleadoPegaduro.C?empleadoPegaduro.C:'',
        IdEmpleadoNomina:empleadoPegaduro.D?empleadoPegaduro.D:'',
        NombrePosicion:empleadoPegaduro.E?empleadoPegaduro.E:'',
        CorreoElectronico:empleadoPegaduro.F?empleadoPegaduro.F:'',
        Depto:empleadoPegaduro.G?empleadoPegaduro.G:'',
        Region:empleadoPegaduro.H?empleadoPegaduro.H:'',
        Nivel:empleadoPegaduro.I?empleadoPegaduro.I:'',
        CorreoRene:correoRene
    }
    empleados.push(empleado)

})






    var xls = json2xls(empleados);

    fs.writeFileSync('Empleados.xlsx', xls, 'binary');
