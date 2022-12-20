const XLSX = require("xlsx");
const now = new Date();
const connection = require("./connection");
const ObjectsToCsv = require('objects-to-csv')
const xl = require("excel4node")


let ciudadesYdestinos = [
   {
    ciudad:"BUENAVENTURA",
    consecionario:"ALKA MOTOR S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"ALMOTORES S.A."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"AUTO ORION S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"AUTOESTE SAS"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"AUTOMOTORES FUJIYAMA DEL MAGDALENA S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"AUTOMOTORES FUJIYAMA S.A.S."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"AUTONIZA S.A"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"BRACHOAUTOS S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"CARRAZOS S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"CENTRAL MOTOR S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"CENTRO MOTORS S.A."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"ARMOTOR S.A-PEREIRA"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"DISTRIKIA S.A.-MEDELLIN"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"DISTRIKIA S.A.-MONTERIA"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"GAMAMOTORS DUITAMA S.A.S."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"JORGE CORTES Y CIA SAS DISTRIBUIDORA DE VEHICULOS"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"KIA PLAZA S.A."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"MARKIA S.A."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"MASSY MOTORS BOGOTÁ S.A.S"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"METROKIA S.A."
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"MOTOR K SAS"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"MUNDO KIA S.A"
  },
  {
    ciudad:"BUENAVENTURA",
    consecionario:"SIDA SAS"
  },
 

]


let ChasisNumerados = []
let ChasisDestinos = []
let Comparados = []
let mensaje =""
const leerexcel = (ruta) => {
  const workbook = XLSX.readFile(ruta);
  const sheet_name_list = workbook.SheetNames;
  const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  console.log("CARGANDO ARCHIVO EXCEL ...");
  console.log("-----------------------------------------------------");
  for (let i = 0; i < xlData.length; i++) {
        ChasisDestinos.push({
      item:i,
      chasis:xlData[i].CHASIS,
      ciudad: xlData[i].CIUDAD,
      consecionario: xlData[i].CONSECIONARIO
    })
  }  
  
};

const TraerChasis = (idMotonave) => {
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
      console.log(`-----`);
    }
    
    connection.end((err)=> {
      if(err) throw err;
      console.log("Conexion cerrada");
    })
    return console.log(`Enumeradas en Total ${ChasisNumerados.length}`);
  });
};

const filtarChasis = (ciudad,consecionario) => {
  Comparados = []
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

  console.log(`${ciudad} - ${consecionario} Son ${result.length} y estan enumeradas ${Comparados.length} - ${mensaje}`);
}

const crearCSV = async (ciudad,consecionario,data) => {
  let wb = new xl.Workbook()

  let ws = wb.addWorksheet("datos")
  ws.cell(1,1).string("item")
  ws.cell(1,2).string("chasis")
  ws.cell(1,3).string("impronta")
  ws.cell(1,4).string("bl")
  ws.cell(1,5).string("ciudad")
  ws.cell(1,6).string("consecionario")
     let fila = 2    
  for (let index = 0; index < data.length; index++) {
    const element = data[index];    
     
    ws.cell(fila,1).number(index+1)      
    ws.cell(fila,2).string(element.chasis)
    ws.cell(fila,3).number(element.impronta)
    ws.cell(fila,4).string(element.bl)
    ws.cell(fila,5).string(ciudad)
    ws.cell(fila,6).string(consecionario)
    fila++
  }

  console.log("Excel generado");

  const pathExcel = `./excelPorCiudad/${ciudad}_${consecionario}.xlsx`
  
  wb.write(pathExcel, (err,stats) => {
    if(err){
      console.log(err);
    }
  })

}

leerexcel("./porciudad.xlsx");
TraerChasis(39)

//filtarChasis("Envigado","METROKIA S.A._ENVIGADO")