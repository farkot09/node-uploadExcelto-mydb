const XLSX = require("xlsx");
const now = new Date();
const connection = require("./connection");

connection.connect((error) => {
  if (error) {
    console.error(error.message);
    return;
  }
  console.log("Connected to the MySQL server.");
});

const leerexcel = (ruta) => {
  const workbook = XLSX.readFile(ruta);
  const sheet_name_list = workbook.SheetNames;
  const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  console.log("CARGANDO ARCHIVO EXCEL ...");
  console.log("-----------------------------------------------------");
  for (let i = 0; i < xlData.length; i++) {
    xlData[i].PLANILLA ? xlData[i].PLANILLA : (xlData[i].PLANILLA = "0");
    const sql =
      "INSERT INTO programacion (id_chasis, conductor, cedula, placa, chasis, planilla, fecha) VALUES (?, ?, ?, ?, ?, ?, ?)";
    const values = [
      0,
      xlData[i].CONDUCTOR,
      xlData[i].CEDULA,
      xlData[i].PLACA,
      xlData[i].CHASIS,
      xlData[i].PLANILLA,
      now.toDateString("yyyy-MM-dd"),
    ];
    connection.query(sql, values, (error, results) => {
      if (error) {
        console.error(error.message);
        return;
      }
      console.log("...");
    });
  }
};

const agregarIdChasis = () => {
  const sql =
    "SELECT chasis.id, programacion.id_chasis,programacion.chasis from chasis,programacion WHERE programacion.id_chasis = 0 AND chasis.chasis = programacion.chasis";
  connection.query(sql, (error, results) => {
    if (error) {
      console.error(error.message);
      return;
    }
    console.log("ACTUALIZANDO ID CHASIS...");
    console.log("xxxxxxxxxxxxxxcambiando el console log de actualizar chasisxxxxxxxxxxxxxxxxxxxxxxxxx");
    for (let i = 0; i < results.length; i++) {
      const sql = "UPDATE programacion SET id_chasis = ? WHERE chasis = ?";
      const values = [results[i].id, results[i].chasis];
      connection.query(sql, values, (error, results) => {
        if (error) {
          console.error(error.message);
          return;
        }
        console.log("...");
      });
    }
  });
};

const eliminarDespachados = () => {
  const sql =
    "DELETE FROM programacion where chasis IN (SELECT chasis from despachos)";
  connection.query(sql, (error, results) => {
    if (error) {
      console.error(error.message);
      return;
    }
    console.log("ELIMINANDO DESPACHADOS...");
    console.log("zzzzzzzzzzz-espacio de eliminar despachos");
  });
};

const actualizarPlanilla = () => {
  const sql =
    "SELECT programacion.planilla,despachos.planilla FROM programacion,despachos ORDER BY despachos.planilla DESC LIMIT 1";
  connection.query(sql, (error, results) => {
    if (error) {
      console.error(error.message);
      return;
    }
    console.log(results[0].planilla);
    dataPlanilla = results[0].planilla;
  });
  const sql2 =
    "SELECT * FROM programacion where planilla = 0 GROUP BY placa ORDER BY placa DESC";
  connection.query(sql2, (error, results) => {
    if (error) {
      console.error(error.message);
      return;
    }
    console.log("ACTUALIZANDO PLANILLA...");
    console.log("-----------------------------------------------------");
    for (let i = 0; i < results.length; i++) {
      const sql = "UPDATE programacion SET planilla = ? WHERE placa = ?";
      const values = [dataPlanilla + i, results[i].placa];
      connection.query(sql, values, (error, results) => {
        if (error) {
          console.error(error.message);
          return;
        }
        console.log("...");
      });
    }
  });
};

leerexcel("./DESPA2.xlsx");
agregarIdChasis();
actualizarPlanilla();
eliminarDespachados();
