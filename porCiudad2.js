const XLSX = require("xlsx");
const now = new Date();
const connection = require("./connection");
const ObjectsToCsv = require('objects-to-csv')
const xl = require("excel4node")


let ciudadesYdestinos = [
  {
    ciudad:"Barranquilla",
    consecionario:"AUTOMOTORES FUJIYAMA S.A.S."
  },
  {
    ciudad:"Bogotá",
    consecionario:"AUTONIZA S.A"
  },
  {
    ciudad:"Bogotá",
    consecionario:"JORGE CORTES Y CIA SAS DISTRIBUIDORA DE VEHICULOS"
  },
  {
    ciudad:"Bogotá",
    consecionario:"KIA PLAZA S.A."
  },
  {
    ciudad:"Bogotá",
    consecionario:"MARKIA S.A."
  },
  {
    ciudad:"Bogotá",
    consecionario:"MASSY MOTORS BOGOTÁ S.A.S"
  },
  {
    ciudad:"Bogotá",
    consecionario:"METROKIA S.A."
  },
  {
    ciudad:"Bucaramanga",
    consecionario:"BRACHOAUTOS S.A.S"
  },
  {
    ciudad:"Bucaramanga",
    consecionario:"CENTRAL MOTOR S.A.S"
  },
  {
    ciudad:"CALI",
    consecionario:"ALMOTORES S.A."
  },
  {
    ciudad:"CALI",
    consecionario:"AUTO ORION S.A.S"
  },
  {
    ciudad:"Duitama",
    consecionario:"GAMAMOTORS DUITAMA S.A.S."
  },
  {
    ciudad:"Envigado",
    consecionario:"METROKIA S.A._ENVIGADO"
  },
  {
    ciudad:"Ibagué",
    consecionario:"SIDA SAS"
  },
  {
    ciudad:"MANIZALES",
    consecionario:"ARMOTOR S.A"
  },
  {
    ciudad:"Medellín",
    consecionario:"DISTRIKIA S.A.-MEDELLIN"
  },
  {
    ciudad:"Medellín",
    consecionario:"MUNDO KIA S.A"
  },
  {
    ciudad:"Montería",
    consecionario:"DISTRIKIA S.A.-MONTERIA"
  },
  {
    ciudad:"Pasto",
    consecionario:"MOTOR K SAS"
  },
  {
    ciudad:"POPAYÁN",
    consecionario:"ALKA MOTOR S.A.S"
  },
  {
    ciudad:"Santa Marta",
    consecionario:"AUTOMOTORES FUJIYAMA DEL MAGDALENA S.A.S"
  },
  {
    ciudad:"Santa Marta",
    consecionario:"METROKIA SANTA MARTA"
  },
  {
    ciudad:"TULUA",
    consecionario:"CENTRO MOTORS S.A."
  },
  {
    ciudad:"Tunja ",
    consecionario:"CARRAZOS S.A.S"
  },
  {
    ciudad:"Valledupar",
    consecionario:"AUTOESTE SAS"
  },
  {
    ciudad:"Villavicencio",
    consecionario:"METROKIA S.A._LLANO"
  },

]

let ChasisNumerados = []
let ChasisDestinos = []
let Comparados = []
let mensaje =""
let info = []

const filtarChasis = (ciudad,consecionario) => {
  Comparados = []
  try {
    const result = ChasisDestinos.filter(data => data.ciudad == ciudad && data.consecionario == consecionario)
  for (let index = 0; index < result.length; index++) {
    const element = result[index].chasis;
    const c = ChasisNumerados.find(data => data.chasis === element)
    if(c !== undefined){
      Comparados.push(c)
        }
  }

  if (result.length === Comparados.length) {
     mensaje = "< -------------------- ESTA COMPLETO ARCHIVO GENERADO"
     crearCSV(ciudad,consecionario,Comparados)
  }else{
    mensaje = "incompleto"
  }

  info.push(({informacion:`${ciudad} - ${consecionario} Son ${result.length} y estan enumeradas ${Comparados.length} - ${mensaje}`}))
  return Comparados,info  
} catch (error) {
    return error
  }
}

const leerexcel = async (ruta,idMotonave) => {
  try {
    const workbook = XLSX.readFile(ruta);
  const sheet_name_list = workbook.SheetNames;
  const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  console.log("CARGANDO ARCHIVO EXCEL ...");
  console.log("-----------------------------------------------------");
  for (let i = 0; i < xlData.length; i++) {
       await ChasisDestinos.push({
      item:i,
      chasis:xlData[i].CHASIS,
      ciudad: xlData[i].CIUDAD,
      consecionario: xlData[i].CONSECIONARIO
    })

    
  }  

  const sql =
  `SELECT numeracion_improntas.id_chasis,numeracion.id_chasis, numeracion_improntas.chasis, numeracion_improntas.numero_impronta,chasis.numeroBl FROM numeracion,numeracion_improntas,chasis WHERE numeracion_improntas.id_chasis = numeracion.id_chasis and numeracion_improntas.id_chasis = chasis.id  and numeracion.id_motonave = ${idMotonave}`;
  connection.query(sql, (error, results) => {
    if (error) {
      console.error(error.message);
      return;
    }

    for (let index = 0; index < results.length; index++) {            
      ChasisNumerados.push({
        item:index,
        chasis:results[index].chasis,
        impronta:results[index].numero_impronta,
        bl:results[index].numeroBl
      })
      
    }

    for (let index = 0; index < ciudadesYdestinos.length; index++) {
      const element = ciudadesYdestinos[index];
      filtarChasis(element.ciudad,element.consecionario)
      const result = ChasisDestinos.filter(data => data.ciudad == element.ciudad && data.consecionario == element.consecionario)
      for (let index = 0; index < result.length; index++) {
        const element = result[index].chasis;
        const c = ChasisNumerados.find(data => data.chasis === element)
        if(c !== undefined){
          Comparados.push(c)
            }
      }

      if (result.length === Comparados.length) {
        mensaje = "< -------------------- ESTA COMPLETO ARCHIVO GENERADO"
        crearCSV(ciudad,consecionario,Comparados)
     }else{
       mensaje = "incompleto"
     }
   
     info.push(({informacion:`${element.ciudad} - ${element.consecionario} Son ${result.length} y estan enumeradas ${Comparados.length} - ${mensaje}`}))
    
      console.log(`-----`);
    }

    connection.end((err)=> {
      if(err) throw err;
      console.log("Conexion cerrada");
    })



  })
  
  return {ChasisDestinos,ChasisNumerados,info}

  } catch (error) {
    return error
  }
}

module.exports = leerexcel