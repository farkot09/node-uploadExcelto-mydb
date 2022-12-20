const express = require("express")
const ejs = require("ejs")

const leerexcel = require("./porCiudad") 

const app = express()


app.use(express.json());
app.set("view engine", "ejs")

app.get('/', async(req, res) => {
   const data = leerexcel("./porciudad.xlsx",38)
    res.render('index.ejs',{titulo:"Titulo Mardito"});
  });


app.listen(3000)

console.log('Server is running in Port', 3000);

