const pageLogin = 'https://envios.mercadolivre.com.br/logistics/management-packages/';
const puppeteer = require('puppeteer');
const ExcelJS = require("exceljs");
const fs = require("fs");
const workbook = new ExcelJS.Workbook();
const express = require('express');
var bodyParser = require('body-parser')
const path = require('path');
const app = express();



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



app.listen(7000,()=>{
  console.log('server rodando!');
})

app.post('/',(req,res) => {
  const path = "./ConsultaId.xlsx";
if (fs.existsSync(path)) {
  // path exists
  fs.unlinkSync(path);
}
  var textid = req.body.textid;

// async function auto() {
  (async () => {

    



    const sheet = workbook.addWorksheet("Consuls de IDs");
    
    sheet.columns = [
      {header: 'ID', key: 'ID'},
      {header: 'Logistic status', key: 'logisticsStatus'},
    
    ];
    
      const browser = await puppeteer.launch({
          headless: false,
        });
  var logado;
  const page = await browser.newPage();
  await page.goto(pageLogin);
  
  
 // logado = await veriFyLogin(page);  
logado = false;   
  
  
  if(logado != true){
  await login(page); 
  await page.waitForNavigation(); 
  //await veriFyLogin(page); 
   
  
  }     
  
  await ids(page);

  await browser.close();
  
  
    // res.redirect('/home')
  
  
  
  //   await browser.close();
  async function ids(page){

        var idss = textid;


          const arr = idss.split(/\r\n/);
          console.log(arr);
      
    
  for (let i = 0; i < arr.length; i++) {
   
      const id = arr[i];
      console.log(id);
      await page.goto(`${"https://envios.mercadolivre.com.br/logistics/management-packages/package/"+id}`);
      // await page.waitForNavigation({ waitUntil: 'networkidle2' });
  
      await Promise.race([
        page.waitForSelector(".package-history-list__row:nth-child(1) .package-history-list__row__items").catch(),
        page.waitForSelector(".ui-empty-state__title").catch(),
        page.waitForSelector(".package-history__title-text").catch()
    ]);



      //await page.waitForSelector(".package-history-list__row:nth-child(1) .package-history-list__row__items") || page.waitForSelector("h4.ui-empty-state__title");
      const company = await page.evaluate(() => {
  
        const text = document.querySelector(".package-history-list__row:nth-child(1) .package-history-list__row__items");
        
  
        if(text){
          const text1 = document.querySelector(".package-history-list__row:nth-child(1) .package-history-list__row__items").innerHTML;
          const text2 = text1.split("<!");
          return text2[0];
         }else{
          const vazio = "-";
          return vazio;
         }
       
       
       });

       sheet.addRow({
        ID: id,
        logisticsStatus: company
  
       })
  
       console.log(company);
      // await page.waitForTimeout(2000);
       
  }
  
 
  sheet.workbook.xlsx.writeFile("ConsultaId.xlsx");
  sheet.workbook.removeWorksheet("Consuls de IDs");

  console.log("Planilha pronta para download.");


 
  

  
  }
  
  
  
  
  async function login(page){
   
    // LOGANDO
  //   DIGITANDO USUARIO
  await page.type('[name="user_id"]','SSP30.BR.A.ROBERTO')
      
  //   CLICANDO EM AVANCAR PARA DIGITAR A SENHA
    await page.click('[type="submit"]')
  
  //   ESPERANDO A PAGINA CARREGAR
    await page.waitForNavigation();
  
    //   DIGITANDO SENHA
    await page.type('[name="password"]','xegefuweyo.74')
  
  //   CLICANDO EM ENTRAR
    await page.click('[name="action"]')
        
  }
  
  
  
  
  async function  veriFyLogin(page, logado){
  
   
  // const login = await page.$(".kraken-navvvv__username");
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
         
          console.log("Aguarde... Voce nao esta logado.")
          return logado;
        }else{
          console.log("Voce esta Logado.")
        }
  
         
  }

  }
  )();
  // await page.evaluate(async() => {
    // res.send('aguarde...');
    res.redirect('/baixar');
  // })
  // res.send(res.render('home',{}));


  
})
  
  
  
  
  
  
  
  