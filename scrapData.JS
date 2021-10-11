// node scrapData.js --myUrl="https://www.mygov.in/covid-19" --myFileName="coronaData.json" --myCountry="meraPyaraBharat" --mycsv="dataOfcoronaofIndia.csv"
let minimist=require("minimist");
let axios=require("axios");
let excel=require("excel4node");
let fs=require("fs");
let args=minimist(process.argv);
let jsdom=require("jsdom");
let promiseOfAxios=axios.get(args.myUrl);
promiseOfAxios.then(function(response){
    let htmlPage=response.data;
    let virtualConsole = new jsdom.VirtualConsole();
    let dom = new jsdom.JSDOM(htmlPage, { virtualConsole });
    let document=dom.window.document;
    let coronastateInfo=document.querySelectorAll("div #stateCount > div .views-row");
let arr=[];
for(let i=0;i<coronastateInfo.length;i++){
  let stateWiseInfo={
  };
   stateWiseInfo.statename=coronastateInfo[i].querySelector("span.st_name").textContent;
   stateWiseInfo.faltu=coronastateInfo[i].querySelector("span.st_number").textContent;
   let innerInfo=coronastateInfo[i].querySelectorAll("div .st_all_counts > div");
   stateWiseInfo.confirmed=innerInfo[0].querySelector("div .tick-confirmed >small").textContent;
   stateWiseInfo.active=innerInfo[1].querySelector("div .tick-active >small").textContent;
   stateWiseInfo.discharged=innerInfo[2].querySelector("div .tick-discharged >small").textContent;
   stateWiseInfo.deaths= innerInfo[3].querySelector("div .tick-death >small").textContent;
   stateWiseInfo.vaccinated=innerInfo[4].querySelector("div .tick-total-vaccine >small").textContent;
    arr.push(stateWiseInfo);
}
let myData=JSON.stringify(arr);
fs.writeFileSync(args.myFileName,myData,"utf-8");
let jfile=fs.readFileSync(args.myFileName,"utf-8");
let rra=JSON.parse(jfile);
let  wb = new excel.Workbook();
let ws = wb.addWorksheet(args.myCountry);
let style1 = wb.createStyle({
  font: {
    color: '#FFFFFF',
    size: 12,
  },
  fill: { 
    type: 'pattern',
    patternType: 'solid',
    fgColor: '#0000ff',
  }});
let style2 = wb.createStyle({
    font: {
      color: '#FFFFFF',
      size: 12,
    },
    fill: { 
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#f7347a',
 }});
let style3 = wb.createStyle({
      font: {
        color: '#FFFFFF',
        size: 12,
      },
      fill: { 
        type: 'pattern',
        patternType: 'solid',
        fgColor: '#b4cbfe',
  }});
let style4 = wb.createStyle({
        font: {
          color: '#FFFFFF',
          size: 12,
        },
        fill: { 
          type: 'pattern',
          patternType: 'solid',
          fgColor: '#5ac18e',
  }});
let style5 = wb.createStyle({
          font: {
            color: '#FFFFFF',
            size: 12,
          },
          fill: { 
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#FF7F50',
  }});
let style6 = wb.createStyle({
            font: {
              color: '#FFFFFF',
              size: 12,
            },
            fill: { 
              type: 'pattern',
              patternType: 'solid',
              fgColor: '#FFD700',
  }});
let style7 = wb.createStyle({
              font: {
                color: '#FFFFFF',
                size: 12,
              },
              fill: { 
                type: 'pattern',
                patternType: 'solid',
                fgColor: '#00FFFF',
}});

ws.cell(1,1).string("STATE").style(style1);
ws.cell(1,2).string("FALTU").style(style2);
ws.cell(1,3).string("CONFIRMED").style(style3);
ws.cell(1,4).string("ACTIVE").style(style4);
ws.cell(1,5).string("DISCHARGED").style(style5);
ws.cell(1,6).string("DEE").style(style6);
ws.cell(1,7).string("VACCINATED").style(style7);

for(let i=0;i<rra.length;i++){
  ws.cell(i+2,1).string(rra[i].statename);
  ws.cell(i+2,2).string(rra[i].faltu);
  ws.cell(i+2,3).string(rra[i].confirmed);
  ws.cell(i+2,4).string(rra[i].active);
  ws.cell(i+2,5).string(rra[i].discharged);
  ws.cell(i+2,6).string(rra[i].deaths);
  ws.cell(i+2,7).string(rra[i].vaccinated);
}
wb.write(args.mycsv);


  

  }

);
