const express = require( "express");
const { Workbook } = require( 'exceljs');
const bodyParser = require( "body-parser");
const ejs = require( "ejs");
const pg = require("pg");

const db=new pg.Client({
  user:"postgres",
  host:"localhost",
  database:"sgp",
  password:"sarthak",
  port:"5432"
});

db.connect();

const app = express();
const port = 3000;
var vissible=0;
let answer=[];

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static("public"));

app.get("/", (req, res) => {
  res.render("index.ejs",{answer:answer,vissible:vissible});
});

app.post("/add",async (req,res)=>{
  let query=req.body.query;
  if (query.trim() !== '') {
      vissible=1;
      const result=await db.query(query);
      answer=result.rows;
      console.log(answer);
      res.redirect("/");
  } else {
      console.log('Please enter something in the input field.');
      res.redirect("/");
  }
});

app.post("/download",(req, res)=>{
  res.redirect("/download");
});

app.get('/download', async (req, res) => {
  if(answer.length != 0) {
    try {
        // Sample array of objects (You can replace this with your data)
        let i=0;
        // Create a new workbook
        const workbook = new Workbook();
        const worksheet = workbook.addWorksheet('Data');
        keys_headers=Object.keys(answer[0]);
        // Add headers
        const columns = keys_headers.map(key => ({ header: key.charAt(0).toUpperCase() + key.slice(1), key, width: 15 }));
        worksheet.columns = columns;
        // Add data
        worksheet.addRows(answer);

        // Generate Excel file and send as a response
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=export.xlsx');
        await workbook.xlsx.write(res);
        res.end();
        res.redirect("/");
    } catch (error) {
        res.status(500).send('Internal Server Error');
    }
  }else{
    res.redirect("/");
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
