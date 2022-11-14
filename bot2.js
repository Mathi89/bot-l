const Painel = require("./public/js/Painel.js");
const painel = new Painel();

// const puppeteer = require('puppeteer');
const puppeteer = require('puppeteer-extra');
const ExcelJS = require("exceljs");
const fs = require("fs");
const workbook = new ExcelJS.Workbook();
const express = require('express');
var bodyParser = require('body-parser')
const path = require('path');
const { request } = require('express');
const app = express();
const port = process.env.PORT || 7001
const  StealthPlugin  =  require ('puppeteer-extra-plugin-stealth') 
puppeteer.use(StealthPlugin())


app.use( bodyParser.json() );       // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true
})); 

app.engine('html', require('ejs').renderFile);
app.set('view engine', 'html');
app.use('/public', express.static(path.join(__dirname, 'public')));
app.set('views', path.join(__dirname, '/pages'));



app.get('/',(req,res)=>{
    
  if(req.query.busca == null){
      res.render('home',{});
  }else{
    const teste = req.query.busca;
    res.render('busca',{teste});
  }


});


app.get('/baixar',(req,res)=>{

    res.render('baixar',{});
  
});

app.get('/downloadlog',(req,res)=>{
  painel.downloadSheet("ConsultaId_logistics",res)
});

app.get('/downloadtms',(req,res)=>{
  painel.downloadSheet("ConsultaId_tms",res)
});

app.get('/downloadnfe',(req,res)=>{
  painel.downloadSheet("Consulta_de_nfe",res)
});


app.listen(port,()=>{
  console.log('server rodando!');
})




app.post('/baixar',(req,res) => {

painel.scraping(req,res);
  
})
  
  
  
  
  
  
  
  