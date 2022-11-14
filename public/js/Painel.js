//const { text } = require('cheerio/lib/api/manipulation');

const { request } = require('http');
const { syncBuiltinESMExports } = require('module');

class Painel{

    constructor(){
        

    }

    downloadSheet(req,res){
        var path = require('path');
        var mime = require('mime');
        const fs = require("fs");

      

        var file = "/Users/Reginaldo/Desktop/loss-master/"+req+".xlsx";

        var filename = path.basename(file);
        var mimetype = mime.lookup(file);

        res.setHeader('Content-disposition', 'attachment; filename=' + filename);
        res.setHeader('Content-type', mimetype);

        var filestream = fs.createReadStream(file);
        filestream.pipe(res);
       
    }




      scraping(req,res){
        // const puppeteer = require('puppeteer');
        const puppeteer = require('puppeteer-extra');
        const  StealthPlugin  =  require ( 'puppeteer-extra-plugin-stealth' ) 
puppeteer . use ( StealthPlugin ( ) )
        const ExcelJS = require("exceljs");
        const fs = require("fs");
        const workbook = new ExcelJS.Workbook();
        const reader = require('xlsx')
        const express = require('express');
        const pageLogin = 'https://envios.mercadolivre.com.br/logistics/management-packages/';
        const loginPagetms = 'https://www.mercadolivre.com/jms/mlb/lgz/msl/login/H4sIAAAAAAAEA22NzWoEIRCE36XB2zBDzGyyCCEkB19DepzWldVVtPePZd89mDC3HIuq76sHxOzDyfC9ECigW4nBBoYBSkR2uSYTVlCQIgzQAtMWlz7BiomYagP16CJP6ze5XLvKYWwEA-CZD8bFfAX19wUDhGboxlRPGM2Vlkug3m6Ez6DgwFyamiaX25ioWlxzDJdKo81pXOrUtdMm-fT54xcQr19CaiE1p38xIXVBe0RPTUg9v-ze5DzL9_1eSL0SY4jwHMBhY8MV7REU1zM9fwBvqWXPJgEAAA/user';
        const pesquisatms = 'https://tms.mercadolivre.com.br/packages/41562442788/detail';
        const nfepage = 'https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ%20gAVw2g=';
        const sitekey = 'e72d2f82-9594-4448-a875-47ded9a1898a';
        const siteurl = 'https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ%20gAVw2g=';

   
        
  var pathSheetLog = "../../ConsultaId_logistics.xlsx";
  var pathSheetLogRead = "./ConsultaId_logistics.xlsx";
  var pathSheetTms = "../../ConsultaId_tms.xlsx";
  var pathSheetNfe = "../../Consulta_de_nfe.xlsx";

if (fs.existsSync(pathSheetLog)) {
  // path exists
  fs.unlinkSync(pathSheetLog);
}

if (fs.existsSync(pathSheetTms)) {
  // path exists
  fs.unlinkSync(pathSheetTms);
}


if (fs.existsSync(pathSheetNfe)) {
  // path exists
  fs.unlinkSync(pathSheetNfe);
}


  var textid = req.body.textid;

  var logstatus = req.body.log;
  var tmsstatus = req.body.tms;
  var nfeverify = req.body.nfe;
  var convertdata = req.body.data;
  var verifyhu = req.body.hu;
  var print = req.body.print;
  if(req.body.print == "on"){
    var print = true;
  }else{
    var print = false;
  }
const test = auto();



async function auto(contagem = false) {




if(logstatus == 'on'){
  
  await ids(print,contagem,true,200);
  console.log("Consulta no logistics terminou.")

}


if(tmsstatus == 'on'){
  const browser2 = await puppeteer.launch({
    headless: false,
  });
  const page2 = await browser2.newPage();
  await page2.goto(loginPagetms);
  var logadotms;
  logadotms = await veriFyLogintms(page2);  

  if(logadotms != true){
    await logintms(page2); 
    await page2.waitForNavigation(); 
    await veriFyLogintms(page2); 

    } 


  await entrandoTM(page2,print);
  //await page.waitForNavigation(); 
  await browser2.close();
  console.log("Consulta no TMS terminou.")

}

if(verifyhu == 'on'){
  const browser2 = await puppeteer.launch({
    headless: true,
  });
  const page2 = await browser2.newPage();
  await page2.goto(loginPagetms);
  var logadotms;
  logadotms = await veriFyLogintms(page2);  

  if(logadotms != true){
    await logintms(page2); 
    await page2.waitForNavigation(); 
    await veriFyLogintms(page2); 

    } 


  await verifyhufunc(page2,print);
  //await page.waitForNavigation(); 
  await browser2.close();
  console.log("Consulta no TMS terminou.")

}

if(nfeverify == 'on'){

  const browser3 = await puppeteer.launch({
    headless: true,
  });
  const page3 = await browser3.newPage();
  // await page3.waitForNavigation(); 
  // const page3 = await browser3.newPage();
  await page3.goto(nfepage);
  
  await nfe(page3,nfepage,print);
  
  // let token = await teste();
  // console.log(token);
  await browser3.close();
  console.log("Consulta de NFE terminou.")

}


if(convertdata == 'on'){
  
  await convertdataf(convertdata,print);
  
  // let token = await teste();
  // console.log(token);

  console.log("Conversão de data pronta.")

}
  
  res.render('baixar',{pathSheetLog:pathSheetLog, pathSheetTms:pathSheetTms, pathSheetNfe,pathSheetNfe});

  
async function curl(options) {
  return new Promise((resolve, reject) => {
      request(options, (err, res, body) => {
          if(err)
              return reject(err);
          resolve(body);
      });
  });
}

async function sleep(sec) {
  return new Promise((resolve, reject) => {
      setTimeout(function() {
          resolve();
      }, sec * 1000);
  });
}




async function disableCssImg(page){
  await page.setRequestInterception(true);
  page.on('request', (req) => {
  if(req.resourceType() == 'stylesheet' || req.resourceType() == 'font' || req.resourceType() == 'image'){
    req.abort();
    }
    else {
    req.continue();
    }
  })

}




  // CONVERTENDO DATA
async function convertdataf(page3,print){
    const page = page3;

    const sheet = workbook.addWorksheet("Converter data");
    
    sheet.columns = [
      {header: 'data', key: 'data'},

    
    ];

        var idss = textid;


          const arr = idss.split(/\r\n/);
          console.log(arr);
      

  for (let i = 0; i < arr.length; i++) {
    const id = arr[i];
    const newdata = id.split("/");
    const datainicio = newdata[0];
    const datameio = newdata[1];
    const datafim = newdata[2];

    const definenewdata = datameio+"/"+datainicio+"/"+datafim;


      if(print == true){
        await page.setViewport({ width: 1280, height: 720 });
           await page.screenshot({path: 'fotoPlanilha/'+id+'.png'});
    }
        sheet.addRow({
          data: definenewdata,
         })


         console.log("Data antiga - "+id+" => Data nova - "+definenewdata);
  }


  sheet.workbook.xlsx.writeFile("Converter_data.xlsx");
  sheet.workbook.removeWorksheet("Converter data");
 
  console.log("Planilha DATA pronta para download.");

}





  // ACESSANDO NFE
  async function nfe(page3,pageurl,print){
    const page = page3;

    const sheet = workbook.addWorksheet("Consulta de nfe");
    
    sheet.columns = [
      {header: 'NFE', key: 'ID'},
      {header: 'R$ VALOR', key: 'nfe'},
      {header: 'DESCRIPTION', key: 'description'},
      //{header: 'Logistic status', key: 'logisticsStatus'},
    
    ];

        var idss = textid;


          const arr = idss.split(/\r\n/);
          console.log(arr);
      

  for (let i = 0; i < arr.length; i++) {

    // await page.waitForNavigation();
    await page.goto(pageurl, {timeout: 0});
    
    
      const id = arr[i];

if(id == "-"){
  const description = "SEM NFE";
  const valor = "0";
  console.log(id+' - '+valor+' - '+description);
  sheet.addRow({
    ID: id,
    nfe: valor,
    description: description

   })

}else{


  await Promise.race([
    page.waitForSelector("#ctl00_ContentPlaceHolder1_txtChaveAcessoResumo").catch(),
    page.waitForSelector(".h-captcha").catch(), 
    ]);

    await page.type('[name="ctl00$ContentPlaceHolder1$txtChaveAcessoResumo"]',id)
    await page.click('.h-captcha');
    await sleep(2);
    await page.click('[type="submit"]')
    // await page.waitForNavigation();
    // await page.waitForSelector("#ctl00_ContentPlaceHolder1_btnVoltar");
    
    // await Promise.race([
    //   page.waitForSelector("#ctl00_ContentPlaceHolder1_btnVoltar").catch(),
    //   ]);

    await Promise.race([
      page.waitForSelector("#NFe > fieldset:nth-child(1) > table > tbody > tr > td:nth-child(6) > span").catch(),
      page.waitForSelector("#conteudoDinamico > div:nth-child(3) > div.XSLTNFeResumida > table:nth-child(11) > tbody > tr:nth-child(2) > td:nth-child(6)").catch(), 
      page.waitForSelector("#ctl00_ContentPlaceHolder1_btnVoltar").catch(), 
      page.waitForSelector("#tab_3").catch()
      
    ]);

    const valor = await page.evaluate(() => {
  
      const valor = document.querySelector("#tab_3");
      
        if(valor){
         
          const valor = document.querySelector('#NFe > fieldset:nth-child(1) > table > tbody > tr > td:nth-child(6) > span').innerHTML;
          return valor;
          
        }else{
          const valor =  document.querySelector('#conteudoDinamico > div:nth-child(3) > div.XSLTNFeResumida > table:nth-child(11) > tbody > tr:nth-child(2) > td:nth-child(6)').innerHTML;
         
          return valor;
        }

      });

    const existeabadescri = await page.evaluate(() => {
  
      const text = document.querySelector("#tab_3");
      
      // const descricao = document.querySelector('#Prod > fieldset > div > table.toggle.box > tbody > tr > td.fixo-prod-serv-descricao > span')
        if(text){
          //  page.click('#tab_3');
          const resultado = true;
           
           return resultado;
          
        }else{
          const resultado = false;
          return resultado;
        }

      });

      if(existeabadescri == true)
      {
       await page.click('#tab_3');
      } 
      

      

    const description = await page.evaluate(() => {
  
      const text = document.querySelector("#Prod > fieldset > div > table.toggle.box > tbody > tr > td.fixo-prod-serv-descricao > span");
      
      // const descricao = document.querySelector('#Prod > fieldset > div > table.toggle.box > tbody > tr > td.fixo-prod-serv-descricao > span')
        if(text){
          //  page.click('#tab_3');
          const descricao = document.querySelector('#Prod > fieldset > div > table.toggle.box > tbody > tr > td.fixo-prod-serv-descricao > span').innerHTML
           
           return descricao;
          
        }else{
          const vazio = "SEM DESCRIÇÃO NA NFE";
          return vazio;
        }

      });
      console.log(i+' - '+id+' - '+valor+' - '+description)
      // console.log(description);

  //     const valor = "ok";   
      //
      //
      //
      //
      //
      //

      

      if(print == true){
        await page.setViewport({ width: 1280, height: 720 });
           await page.screenshot({path: 'fotoPlanilha/'+id+'.png'});
    }
        sheet.addRow({
          ID: id,
          nfe: valor,
          description: description
    
         })


        }
      
  }
  
 
  sheet.workbook.xlsx.writeFile("Consulta_de_nfe.xlsx");
 sheet.workbook.removeWorksheet("Consulta de nfe");

 console.log("Planilha NFE pronta para download.");


 
  

  
  }




  //VERIFICANDO HU
  async function verifyhufunc(page,print){

    const sheettms = workbook.addWorksheet("Consulta hu");
    
    sheettms.columns = [
      {header: 'romaneio', key: 'romaneio'},
      {header: 'quantidade de pacotes', key: 'qtd'},
      // {header: 'tms hitoric', key: 'tmshistoric'},
      // {header: 'nfe', key: 'nfe'},
      //{header: 'Logistic status', key: 'logisticsStatus'},

   
    
    ];

    var idss = textid;


    const arr2 = idss.split(/\r\n/);
    console.log(arr2);


for (let i = 0; i < arr2.length; i++) {
  

const id = arr2[i];
// console.log(id);
await page.goto(`${"https://tms.mercadolivre.com.br/outbounds/"+id+"/detail"}`, {timeout: 0});



 //PEGANDO STATUS DO TMS

 await Promise.race([
  page.waitForSelector(".ui-empty-state__title").catch(),
  page.waitForSelector("#root-app > div > div > div > h4").catch(),
  page.waitForSelector(".sidebar__item-title").catch(),
  page.waitForSelector("#outbounds-detail > div > div.outbounds-detail__sidebar.sidebar > div > div:nth-child(3) > h2").catch()
 // page.waitForSelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.package-detail.package-detail__info.collapsible-panel > div.collapsible-panel__content.collapsible-panel__content--show > div:nth-child(6) > label:nth-child(4) > div.andes-form-control__control > input").catch()
]);

const quantidade = await page.evaluate(() => {

  const text = document.querySelector("#outbounds-detail > div > div.outbounds-detail__sidebar.sidebar > div > div:nth-child(3) > h2");
  

  if(text){
    const text1 = document.querySelector("#outbounds-detail > div > div.outbounds-detail__sidebar.sidebar > div > div:nth-child(3) > h2").innerHTML;
    const text2 = text1.split(" ");
    return text2[0];
   }else{
    const vazio = "-";
    return vazio;
   }
 
 });


 sheettms.addRow({
  romaneio: id,
  qtd: quantidade,

 })
//  console.log(i+' - '+id+' - '+companytms+' - '+nfe);
//  await page.waitForTimeout(2000);

}
sheettms.workbook.xlsx.writeFile("ConsultaId_hu.xlsx");
sheettms.workbook.removeWorksheet("Consulta hu");

console.log("Planilha TMS pronta para download.");


  }




// ACESSANDO TMS
  async function entrandoTM(page,print){

    const sheettms = workbook.addWorksheet("Consuls de IDs no tms");
    
    sheettms.columns = [
      {header: 'ID', key: 'ID'},
      {header: 'tms status', key: 'tmsstatus'},
      {header: 'canalização', key: 'canalizacao'},
      {header: 'tms hitoric', key: 'tmshistoric'},
      {header: 'nfe', key: 'nfe'},
      
      //{header: 'Logistic status', key: 'logisticsStatus'},

   
    
    ];

    var idss = textid;


    const arr2 = idss.split(/\r\n/);
    // console.log(arr2);


for (let i = 0; i < arr2.length; i++) {
  

const id = arr2[i];
// console.log(id);
await page.goto(`${"https://tms.mercadolivre.com.br/packages/"+id+"/detail"}`, {timeout: 0});

//await page.waitForNavigation({ waitUntil: 'load' });




//PEGANDO HISTORICO DO TMS
await Promise.race([
  page.waitForSelector(".ui-empty-state__title").catch(),
  page.waitForSelector("#root-app > div > div > div > h4").catch(),
  page.waitForSelector(".sidebar__item-title").catch(),
]);
const companytms = await page.evaluate(() => {

  const text = document.querySelector("#package-detail > div > div.package-detail-sidebar.sidebar > div:nth-child(2) > ul > li:nth-child(1) > h3");
  

  if(text){
    const text1 = document.querySelector("#package-detail > div > div.package-detail-sidebar.sidebar > div:nth-child(2) > ul > li:nth-child(1) > h3").innerHTML;
    const text2 = text1.split("<!");
    return text2[0];
   }else{
    const vazio = "-";
    return vazio;
   }
 
 });



 const canalizacao = await page.evaluate(() => {

  const text = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.package-detail.package-detail__info.collapsible-panel > div.collapsible-panel__content.collapsible-panel__content--show > div:nth-child(4) > label:nth-child(5) > div.andes-form-control__control > input");
  

  if(text){
    const text1 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.package-detail.package-detail__info.collapsible-panel > div.collapsible-panel__content.collapsible-panel__content--show > div:nth-child(4) > label:nth-child(5) > div.andes-form-control__control > input").value;
    const text2 = text1.split("<!");
    return text2[0];
   }else{
    const vazio = "-";
    return vazio;
   }
 
 });

 //await page.screenshot({path: 'print/'+id+'.png'});








 //PEGANDO STATUS DO TMS

//  await Promise.race([
//   page.waitForSelector(".ui-empty-state__title").catch(),
//   page.waitForSelector("#root-app > div > div > div > h4").catch(),
//   page.waitForSelector(".sidebar__item-title").catch(),
//   page.waitForSelector(".ui-empty-state__title").catch(),
//   page.waitForSelector("#root-app > div > div > div > h4").catch(),
//   page.waitForSelector(".sidebar__item-title").catch()
//  // page.waitForSelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.package-detail.package-detail__info.collapsible-panel > div.collapsible-panel__content.collapsible-panel__content--show > div:nth-child(6) > label:nth-child(4) > div.andes-form-control__control > input").catch()
// ]);
const status = await page.evaluate(() => {

  const status = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.shipment-details > div.shipment-details__content > div:nth-child(2) > label:nth-child(1) > div.andes-form-control__control > span");
  
  const status2 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.shipment-details > div.shipment-details__content > div:nth-child(4) > label:nth-child(1) > div.andes-form-control__control > input");
  

  if(status){
    const text1 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.shipment-details > div.shipment-details__content > div:nth-child(2) > label:nth-child(1) > div.andes-form-control__control > input").value;
    //return text1;
    const text2 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.shipment-details > div.shipment-details__content > div:nth-child(4) > label:nth-child(1) > div.andes-form-control__control > input").value;

    if(text1 == "" || text1 == "Status"){

       return text2;
      
    }else{
  
      return text1;

    }


   }else if(status2){

    const text1 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.shipment-details > div.shipment-details__content > div:nth-child(2) > label:nth-child(1) > div.andes-form-control__control > span").innerHTML;
    //return text1;
    const text2 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.shipment-details > div.shipment-details__content > div:nth-child(4) > label:nth-child(1) > div.andes-form-control__control > input").value;

    if(text1 == "" || text1 == "Status"){

       return text2;
      
    }else{
  
      return text1;

    }

   }  else{

    const vazio = "-";
    return vazio;
   }
 
 });











 //PEGANDO NFE DO TMS

//  await Promise.race([
//   page.waitForSelector(".ui-empty-state__title").catch(),
//   page.waitForSelector("#root-app > div > div > div > h4").catch(),
//   page.waitForSelector(".sidebar__item-title").catch()
// ]);
const nfe = await page.evaluate(() => {

  const nfe = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.package-detail.package-detail__info.collapsible-panel > div.collapsible-panel__content.collapsible-panel__content--show > div:nth-child(6) > label:nth-child(4) > div.andes-form-control__control > input");
  

  if(nfe){
    const text1 = document.querySelector("#package-detail > div > div.layout__content > div.layout__container > div.andes-card.package-detail.package-detail__info.collapsible-panel > div.collapsible-panel__content.collapsible-panel__content--show > div:nth-child(6) > label:nth-child(4) > div.andes-form-control__control > input").value;
    return text1;
   }else{
    const vazio = "-";
    return vazio;
   }
 
 });


 if(print == true){
  await page.setViewport({ width: 1280, height: 720 });
     await page.screenshot({path: 'fotoPlanilha/'+id+'.png'});
}

 sheettms.addRow({
  ID: id,
  tmshistoric: companytms,
  canalizacao: canalizacao,
  tmsstatus: status,
  nfe: nfe

 })

 


 console.log(i+' - '+id+' - '+canalizacao+' - '+companytms+' - '+nfe);
//  await page.waitForTimeout(2000);

}
sheettms.workbook.xlsx.writeFile("ConsultaId_tms.xlsx");
sheettms.workbook.removeWorksheet("Consuls de IDs no tms");

console.log("Planilha TMS pronta para download.");


  }




// ACESSANDO LOGISTICS
  async function ids(print,icont = false,headlessstatus = false,maxporvez){


    var page = false;
    var browser = false;
    var funcionando = false;

    var iniciarnovo = false;
    var titleidlog = 'ID';
    var titlelogstatus = 'Logistic status';
    var titlelogstatus2 = 'Logistic status 2';
    var titlelogorigem = 'Origem';
    var titlelogrota = 'Rota';
    var titlelogparada = 'Parada';
    var titlelogulhistorico = 'Ultimo historico';
    var titlelogultinventario = 'Ultimo Inventario';

    var titleplanilha = 'Consuls de IDs';
   

      
      // console.log("teste")
    var workbook = new ExcelJS.Workbook()
    var sheet = workbook.addWorksheet(titleplanilha);
    sheet.columns = [
      {header: titleidlog, key: 'ID'},
      {header: titlelogstatus, key: 'logisticsStatus'},
      {header: titlelogstatus2, key: 'logisticsStatus2'},
      {header: titlelogorigem, key: 'origem'},
      {header: titlelogrota, key: 'rota'},
      {header: titlelogparada, key: 'parada'},
      {header: titlelogulhistorico, key: 'datahora'},
      {header: titlelogultinventario, key: 'inventario'},
    
    ];
        
        var idss = textid;


          const arr = idss.split(/\r\n/);

      var contagem = 0;
      if(icont === false){
       var i = 0;
      }else{
       var i = icont;
      }
  for (i; i < arr.length; i++) {
    // console.log(contagem)
    
if(funcionando == false || browser == false)
{

  funcionando = true;
     //INICIOANDO ACOES DO LOGISTICS
     var browser = await puppeteer.launch({
      headless: headlessstatus,
    });
  var logado;
  var page = await browser.newPage();

  await disableCssImg(page);
  // await page.setRequestInterception(true);
  // page.on('request', (req) => {
  //   if(req.resourceType() == 'stylesheet' || req.resourceType() == 'font' || req.resourceType() == 'image'){
  //   req.abort();
  //   }
  //   else {
  //   req.continue();
  //   }
  //   });




  await page.goto(pageLogin);
  
  
  logado = await veriFyLogin(page);     

  if(logado != true){
  await login(page); 
  await page.waitForNavigation(); 
  await veriFyLogin(page); 
   
  }
  

}
     
// await disableCssImg(page)
      const id = arr[i];
      
      await page.goto(`${"https://envios.mercadolivre.com.br/logistics/management-packages/package/"+id}`, {timeout: 0});
      // await page.waitForNavigation({ waitUntil: 'networkidle2' });
  
      await Promise.race([
        page.waitForSelector(".package-history-list__row:nth-child(1) .package-history-list__row__items").catch(),
        page.waitForSelector(".ui-empty-state__title").catch(),
        page.waitForSelector("#package-edit > div.package-history > ul > li:nth-child(1) > div:nth-child(1) > span").catch(),
        page.waitForSelector("#package-edit > div.package-edit-content > div:nth-child(3) > div:nth-child(3) > div > div:nth-child(2) > div.andes-form-control.andes-form-control--textfield.andes-form-control--default.andes-form-control--disabled.historical-inventory-package-input > label > div > input").catch(),
        page.waitForSelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(1) > label > div > input").catch(),
        page.waitForSelector(".package-history__title-text").catch()
     ]);
      const company = await page.evaluate(() => {
  
        const text = document.querySelector(".package-history-list__row:nth-child(1) .package-history-list__row__items");
        
  
        if(text){
          const text1 = document.querySelector(".package-history-list__row:nth-child(1) .package-history-list__row__items").innerText;
          // const text2 = text1.split("<!");
          const text2 = text1.split("\n");
          return text2[0];
         }else{
          const vazio = "-";
          return vazio;
         }

        
       
       
       });



      const status2 = await page.evaluate(() => {
  
        const text = document.querySelector(".package-history-list__row:nth-child(2) .package-history-list__row__items");
        
  
        if(text){
          const text1 = document.querySelector(".package-history-list__row:nth-child(2) .package-history-list__row__items").innerHTML;
          const text2 = text1.split("<!");
          return text2[0];
         }else{
          const vazio = "-";
          return vazio;
         }

        
       
       
       });



      const origem = await page.evaluate(() => {
  
        const text = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.package-status.andes-card--padding-default > div.package-status-content > div.package-status-input--origin > p");
        
  
        if(text){
          const text1 = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.package-status.andes-card--padding-default > div.package-status-content > div.package-status-input--origin > p").innerHTML;
          return text1;
         }else{
          const vazio = "-";
          return vazio;
         }

        
       
       
       });



  
      const datahoraold = await page.evaluate(() => {
  
        const text = document.querySelector("#package-edit > div.package-history > ul > li:nth-child(1) > div:nth-child(1) > span");
        
  
        if(text){
          const text1 = document.querySelector("#package-edit > div.package-history > ul > li:nth-child(1) > div:nth-child(1) > span").innerText;
          return text1;
         }else{
          const vazio = "-";
          return vazio;
         }

        
       
       
       });

       const datahora = datahoraold.replace(" |","");

      const inventario = await page.evaluate(() => {
  
        const text = document.querySelector("#package-edit > div.package-edit-content > div:nth-child(3) > div:nth-child(3) > div > div:nth-child(2) > div.andes-form-control.andes-form-control--textfield.andes-form-control--default.andes-form-control--disabled.historical-inventory-package-input > label > div > input");
        
  
        if(text){
          const text1 = document.querySelector("#package-edit > div.package-edit-content > div:nth-child(3) > div:nth-child(3) > div > div:nth-child(2) > div.andes-form-control.andes-form-control--textfield.andes-form-control--default.andes-form-control--disabled.historical-inventory-package-input > label > div > input").value;
          return text1;
         }else{
          const vazio = "-";
          return vazio;
         }

        
       
       
       });

      



      const rota = await page.evaluate(() => {
  
        const text = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(1) > label > div > input");
  
        if(text){
          const text1 = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(1) > label > div > input").value;
          return text1;
         }else{
          const vazio = "-";
          return vazio;
         }

        
       
       
       });




       const paradanumero = await page.evaluate(() => {
  
        const text = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(2) > label > div > input");
  
        if(text){
          const text1 = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(2) > label > div > input").value;
          return text1;
         }else{
          const vazio = "";
          return vazio;
         }

        
       
       
       });


       const paradaletra = await page.evaluate(() => {
  
        const text = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(3) > label > div > input");
  
        if(text){
          const text1 = document.querySelector("#package-edit > div.package-edit-content > div.andes-card.andes-card--flat.andes-card--default.collapse-card.route-details.andes-card--padding-default > div:nth-child(3) > div > div > div:nth-child(3) > label > div > input").value;
          return text1;
         }else{
          const vazio = "";
          return vazio;
         }

        
       
       
       });

       const parada = paradanumero+paradaletra;



if(print == true){
    await page.setViewport({ width: 1280, height: 720 });
       await page.screenshot({path: 'fotoPlanilha/'+id+'.png'});
}
      


  console.log(i+' : '+id+" -> "+company+"; Rota: "+rota+"; Parada: "+parada+"; "+datahora);
       sheet.addRow({
        ID: id,
        logisticsStatus: company,
        logisticsStatus2: status2,
        origem: origem,
        rota: rota,
        parada:parada,
        datahora:datahora,
        inventario:inventario
  
       })

       sheet.workbook.xlsx.writeFile("ConsultaId_logistics.xlsx");


      if(contagem > maxporvez){
      

  
        await browser.close();

      funcionando = false;
      contagem = 0;
        
        
      }else{
        contagem++;
      }
      // break;


  }
  

    sheet.workbook.removeWorksheet(titleplanilha);
    console.log("Planilha LOGISTICS pronta para download.");
 

  


 
  

  
  }
  


  // LOGIN DO TMS
  async function logintms(page){
   
    // LOGANDO
  //   DIGITANDO USUARIO
  await page.type('[name="user_id"]','ME.BR.REGINALDOJESUS')
      
  //   CLICANDO EM AVANCAR PARA DIGITAR A SENHA
    await page.click('[type="submit"]')
  
  //   ESPERANDO A PAGINA CARREGAR
  //  await page.waitForNavigation();
await page.waitForSelector('#password')
  
    //   DIGITANDO SENHA

    await page.type('[name="password"]','casa.1209')
  
  //   CLICANDO EM ENTRAR
    await page.click('[name="action"]')
        
  }
  
 async function apagarInput(page,element)
 {
    await page.focus(element);
    await page.keyboard.down('Control');
    await page.keyboard.press('A');
    await page.keyboard.up('Control');
    await page.keyboard.press('Backspace');
 }

  async function  apliquessp30svc(page){

    console.log("Um momento, estamos aplicando as configurações necessárias...")

    await sleep(1)
    await page.click('#kraken__app-nav-trigger > div > span.appNav_selector__selection')

    await sleep(3)

    // await Promise.race([
    //   page.waitForSelector("body > div > div > div > div.andes-modal__scroll > div.andes-modal__actions > button.andes-button.andes-button--medium.andes-button--quiet").catch(),
    // ]);

    await page.click('body > div > div > div > div.andes-modal__scroll > div.andes-modal__actions > button.andes-button.andes-button--medium.andes-button--quiet')

    await sleep(1)

    await page.click('#remote-module-container-kraken-frm-provider-attributessite-default > div > button')

    await sleep(1)


    await page.click('#andes-dropdown-site-list-option-MLB')

    await sleep(1)

    await page.click('#remote-module-container-kraken-frm-provider-attributesregion-default > div > form > div > div.search-area')

    await sleep(1)


    await page.click('#remote-module-container-kraken-frm-provider-attributesregion-default > div > form > div > div.andes-card.andes-card--flat.andes-card--default.andes-card--padding-default > div > ul > li:nth-child(2)')

    await sleep(1)


    await page.click('#remote-module-container-kraken-frm-provider-attributesservicecenter-default > div > form > div > div.search-area')

    await sleep(1)

    await page.click('#remote-module-container-kraken-frm-provider-attributesservicecenter-default > div > form > div > div.andes-card.andes-card--flat.andes-card--default.andes-card--padding-default > div > ul > li')

   

    await page.click('body > div > div > div > div.andes-modal__scroll > div.andes-modal__actions > button.andes-button.andes-button--medium.andes-button--loud')
  
  }
  
  // LOGIN DO LOGISTICS
  async function login(page){
   
    // LOGANDO
  //   DIGITANDO USUARIO
    await page.type('[name="user_id"]','ME.MLB.SSP30MATEUSLOSS')
      
  //   CLICANDO EM AVANCAR PARA DIGITAR A SENHA
    await page.click('[type="submit"]')
  
  //   ESPERANDO A PAGINA CARREGAR
  //  await page.waitForNavigation();
    await page.waitForSelector('#password')
  
    //   DIGITANDO SENHA

    await page.type('[name="password"]','zaxozevavi.34')
  
  //   CLICANDO EM ENTRAR
    await page.click('[name="action"]')
        
  }


  async function  veriFyLogintms(page, logado){
  
    await Promise.race([
      page.waitForSelector(".ui-header__user-nickname").catch(),
      page.waitForSelector("#help-text").catch(),
  ]);
    const login = await page.evaluate(() => {
    
       const text = !!document.querySelector(".ui-header__user-nickname");
    
       if(text){
        return text;
       }else{
        return null;
       }
       
     });
    
    logado;
          
          if(login){
             logado = true;
           }else{
             logado = false;
           }
           
           if(logado != true){
           
            console.log("Aguarde... Voce nao esta logado no TMS.")
            return logado;
          }else{
            console.log("Voce esta Logado.")
          }
    
           
  }



  
  
  
  
  async function  veriFyLogin(page, logado){
  
   
    await Promise.race([
      page.waitForSelector(".kraken-nav__username").catch(),
      page.waitForSelector("#help-text").catch(),
  ]);
  const login = await page.evaluate(() => {
  
     const text = !!document.querySelector(".kraken-nav__username");
  
     if(text){
      return text;
     }else{
      return null;
     }
     
   });
  
  logado;
        
        if(login){
           logado = true;
         }else{
           logado = false;
         }
         
         if(logado != true){
         
          console.log("Aguarde... Voce nao esta logado no LOGISTICS.")
          return logado;
        }else{
          console.log("Voce esta Logado.")
        }
  
         
  }

  }


    }

}

module.exports = Painel;