const pageLogin = 'https://envios.mercadolivre.com.br/logistics/management-packages/';
const puppeteer = require('puppeteer');
const ExcelJS = require("exceljs");
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Consul de IDs");




// btn.addEventListener('click', async function(){
//   sheet.columns = [
//     {header: 'ID', key: 'ID'},
//     {header: 'Logistic status', key: 'logisticsStatus'},
  
//   ];
  
//     const browser = await puppeteer.launch({
//         headless: false,
//       });
// var logado;
// const page = await browser.newPage();
// await page.goto(pageLogin);


// logado = await veriFyLogin(page);     


// if(logado != true){
// await login(page); 
// await page.waitForNavigation(); 
// await veriFyLogin(page);    

// }     

// await ids(page);



// })

async function auto() {
// (async () => {
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


logado = await veriFyLogin(page);     


if(logado != true){
await login(page); 
await page.waitForNavigation(); 
await veriFyLogin(page);    

}     

await ids(page);



//   await browser.close();
}
// )();



async function ids(page){

  const ids = await page.evaluate(() => {

    // const ids =  document.getElementById("consulta-id").value;
    ids = "41448826879,41446487145,41449820347,41451096506,41453457474,41450508402,41452528914,41453508645";
    const arr = ids.split(",");
    return arr;
    
  });
  
for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    await page.goto(`${"https://envios.mercadolivre.com.br/logistics/management-packages/package/"+id}`);
    // await page.waitForNavigation({ waitUntil: 'networkidle2' });

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
}

sheet.workbook.xlsx.writeFile("ConsultaId.xlsx");

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



