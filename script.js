function enviaEmailComDados() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var qtd = 0;


  var data = [];
  var valorData = ss.getRange(4, 3).getValue();
  for (i=3;i<=18;i++) {

    var dataRange = ss.getRange(i, 3).getValue();
    if (dataRange == valorData) {
        data.push(dataRange);
        qtd++;

        }
  }


  var tipo = [];
  var tipoQtd = 0;
  for (i=2;i<= qtd;i++) {

    var tipoRange = ss.getRange(i, 1).getValue();
    tipo.push(tipoRange);
    tipoQtd++;
  }

  Logger.log(qtd)
  Logger.log(data)


  var troca = 'TROCA'
  var teste = 'TESTE'
  var both = 'TESTE E TROCA'

  var arrayTroca = (tipo.indexOf("TROCA") > -1);
  var arrayTeste = (tipo.indexOf("TESTE") > -1);

  Logger.log(arrayTroca)

  if (arrayTroca == true && arrayTeste == true ) {
    var resultTipo = both
  }

  else if (arrayTeste == false) {
    var resultTipo = troca
  }

  else {
    var resultTipo = teste
  }

  Logger.log(tipoQtd);


  var hora = []
  for(i=2;i<=qtd;i++) {
   var horaRange = ss.getRange(i, 5).getValue();
   hora.push(horaRange);
  }

  var pressao = [];
  var pressaoQtd = 0;
  for (i=3;i<=qtd;i++) {
    var pressaoRange = ss.getRange(i, 6).getValue();
    pressao.push(pressaoRange);
    pressaoQtd += pressaoRange;
  }

  var maxPress = Math.max.apply(null, pressao);
  var minPress = Math.min.apply(null, pressao);
  var mediaPress = (pressaoQtd / tipoQtd);

//  Logger.log(mediaPress.toFixed(2));

  var umidade = [];
  var umidadeQtd = 0;
  for (i=3;i<=qtd;i++) {
    var umidadeRange = ss.getRange(i, 7).getValue();
    umidade.push(umidadeRange);
    umidadeQtd += umidadeRange;
  }


  var maxUmi = Math.max.apply(null, umidade);
  var minUmi = Math.min.apply(null, umidade);
  var mediaUmi = (umidadeQtd / tipoQtd);

//  Logger.log(mediaUmi.toFixed(2));

  var temperatura = [];
  var tempQtd = 0;

  for (i=3;i<=qtd;i++) {
    var temperaturaRange = ss.getRange(i, 8).getValue();
    temperatura.push(temperaturaRange);
    tempQtd += temperaturaRange;
  }

  var maxTemp = Math.max.apply(null, temperatura);
  var minTemp = Math.min.apply(null, temperatura);
  var mediaTemp = (tempQtd / tipoQtd);

  Logger.log(maxTemp);
  Logger.log(minTemp);
  Logger.log(mediaTemp.toFixed(2))

 var message="<table border='1',cellpadding='10',cellspacing ='0', width ='900'>"
    +"<tr>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '125'>"+"OPERAÇ."+"</td>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '125'>"+"QUANT."+"</td>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '136'>"+"DATA"+"</td>"
    +"<td bgcolor = '#F6F14E', Align = 'center', width = '250'>"+"PRESSÃO"+"</td>"
    +"<td bgcolor = '#8EF64E', Align = 'center', width = '250'>"+"UMIDADE"+"</td>"
    +"<td bgcolor = '#F6B44E', Align = 'center', width = '240'>"+"TEMPERATURA"+"</td>"
    +"</tr>"
    +"</table>"


    +"<table border='1',cellpadding='10',cellspacing ='0', width = '900'>"
    +"<tr>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '85'>"+""+"</td>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '85'>"+""+"</td>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '85'>"+""+"</td>"
    +"<td bgcolor = '#F3FF33', Align = 'center', width = '50'>"+"MIN"+"</td>"
    +"<td bgcolor = '#F3FF33', Align = 'center', width = '50'>"+"MED"+"</td>"
    +"<td bgcolor = '#F3FF33', Align = 'center', width = '50'>"+"MÁX"+"</td>"
    +"<td bgcolor = '#07F22E', Align = 'center', width = '50'>"+"MIN"+"</td>"
    +"<td bgcolor = '#07F22E', Align = 'center', width = '50'>"+"MED"+"</td>"
    +"<td bgcolor = '#07F22E', Align = 'center', width = '50'>"+"MÁX"+"</td>"
    +"<td bgcolor = '#F28D27', Align = 'center', width = '50'>"+"MIN"+"</td>"
    +"<td bgcolor = '#F28D27', Align = 'center', width = '50'>"+"MED"+"</td>"
    +"<td bgcolor = '#F28D27', Align = 'center', width = '50'>"+"MÁX"+"</td>"
    +"</tr>"
    +"</table>"

    +"<table border='1',cellpadding='10',cellspacing ='0', width = '900'>"
    +"<tr>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '85', height = '100'>"+resultTipo+"</td>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '85', height = '100'>"+qtd+"</td>"
    +"<td bgcolor = '#bed3f4', Align = 'center', width = '85', height = '100'>"+valorData+"</td>"

   // +"<td bgcolor = '#bed3f4', Align = 'center', width = '85', height = '100'>"+data[(data.length - 2)]+"</td>"

    +"<td bgcolor = '#F3FF33', Align = 'center', width = '50'>"+minPress+"</td>"
    +"<td bgcolor = '#F3FF33', Align = 'center', width = '50'>"+mediaPress.toFixed(2)+"</td>"
    +"<td bgcolor = '#F3FF33', Align = 'center', width = '50'>"+maxPress+"</td>"
    +"<td bgcolor = '#07F22E', Align = 'center', width = '50'>"+minUmi+"</td>"
    +"<td bgcolor = '#07F22E', Align = 'center', width = '50'>"+mediaUmi.toFixed(2)+"</td>"
    +"<td bgcolor = '#07F22E', Align = 'center', width = '50'>"+maxUmi+"</td>"
    +"<td bgcolor = '#F28D27', Align = 'center', width = '50'>"+minTemp+"</td>"
    +"<td bgcolor = '#F28D27', Align = 'center', width = '50'>"+mediaTemp.toFixed(2)+"</td>"
    +"<td bgcolor = '#F28D27', Align = 'center', width = '50'>"+maxTemp+"</td>"
    +"</tr>"
    +"</table>"


  // Logger.log(data.length)



  var subject = 'TITULO DO EMAIL';
  var emailAddress = 'your@email.com';
  MailApp.sendEmail({
    to: "recipient@email.com",
    subject: subject,
    htmlBody: message});



}