///////////////////////////////////////////////
//----------Created by Nathan Serra----------//
//-------------------------------------------//
//Soma os valores de horário de uma planilha //
///////////////////////////////////////////////

function CALCHORAS() {
  // recupera a planilha
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  // Pega o nome de todas as planilhas e coloca numa lista
  var end = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  for (var i = 0; i < end; ++i) {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    // Variável com todos os nomes de planilha
    var List = [[sheets[i].getName()]];
  }
  
  // Variáveis da última planilha
  var sheet = sheets[end-1];
  var lastColumn = sheet.getLastColumn();
  var row = sheet.getLastRow()+1;
  var column = sheet.getLastColumn();
  
  // Variáveis auxiliares
  var somaTotal = 0;
  var somaBruta = [];
  var soma = 0;
  var time = "";
  
  //Variáveis modificaveis de acordo com a necessidade
  var n = 4; //é usado para pular a cada n linhas
  var m = 2; //Linhas do título
  var o = 2; //Linhas à esquerda sem valores
  
  
  // Soma os dados em todas as planilhas
  for(var x = 0;x<sheets.length;x++){
    //Planilha atual
    var sheetAux = sheets[x];    
    var lRow = sheetAux.getLastRow();
    var lColumn = sheetAux.getLastColumn();
    
    // Adiciona uma linha no vetor para esta planilha   
    somaBruta.push([]);
    
    // Somar os dados de uma linha, para cada linha até a última
    for(var i = (m+1); i < lRow+1; i++){ //Anda pela linha 
      //A n-esima linha contem texto, como o local onde foi trabalhado.
      if((i-m)%n!=0){
        for(var j = o+1; j < lColumn+1; j++){ //Anda pela coluna
          if(sheetAux.getRange(i, j).getValue() != "" ) //Se não tiver nada na célula, não chama a função
            soma+= somarHora(i,j,sheetAux); //Retorna um valor decimal
        }
        somaBruta[x].push(soma); //Adiciona no vetor para as somas mais na frente
        
        time = decimalToHour(soma); //transforma decimal em tempo        
        
        var cell = sheetAux.getRange(i, (lColumn+1)); //Coloca o valor na linha atual, mas na coluna seguinte à ultima
        cell.setValue(time);
        cell.setBackground("#3498db"); //Azul
        
        /* somente para desenvolvimento
        sheetAux.getRange(i, (lColumn+2)).setValue(somaBruta[x][i-2]);
        */
        time = "";
        soma = 0;
        
      } else { // Caso seja a n-ésima linha ele deve somar os valores anteriores
        
        for(var k = n+m; k>m+1; k--) //Soma os valores anteriores, em função de 'n' e 'm'
          soma+=somaBruta[x][i-k];
        
        somaBruta[x].push(soma); // Adiciona na n-ésima linha o valor da soma para cálculos futuros
        
        time = decimalToHour(soma); //Transforma decimal em tempo
        
        var cell = sheetAux.getRange(i, (lColumn+1)); //Coloca o valor na linha atual, mas na coluna seguinte à ultima
        cell.setValue(time);
        cell.setBackground("#34495e"); //Azul escuro
        cell.setFontColor("white");
        
        /* somente para desenvolvimento
        sheetAux.getRange(i, (lColumn+2)).setValue(somaBruta[x][i-2]);
        */
        time = "";
        soma = 0;        
      }
    }
  }
  
  somarTudo(somaBruta, sheets,n,time,m); //Soma os valores dos vetores
  
  
}

/**
* [Soma todos os n-ésimos valores de um vetor e coloca numa célula]
* @param  {[Array]} somaBruta [Vetor com todos os valores somados das planilhas]
* @param  {[Array]} sheets    [Vetor com todas as planilhas]
* @return {[undefined]}       [Não retorna nada.]
*/
function somarTudo(somaBruta,sheets,n,time,m){
  // Variáveis da última planilha
  var sheet = sheets[sheets.length-1];
  var lastColumn = sheet.getLastColumn();
  var somaTotal = 0;
  var row = sheet.getLastRow()+1;
  var column = sheet.getLastColumn();
  var linha = m+1;
  
  //Soma de a n-ésima linha de cada planilha
  for(var i = (n-1); i<somaBruta[0].length;i+=n){    
    for(var x = 0; x<somaBruta.length;x++)
      somaTotal+=somaBruta[x][i];
    
    time = decimalToHour(somaTotal); //Transforma decimal em tempo
    
    var cell = sheet.getRange((linha),(lastColumn+2)); //Coloca o valor na linha, mas na coluna seguinte à ultima
    cell.setValue(time);
    cell.setBackground("#34495e"); //Azul escuro
    cell.setFontColor("white");
    cell.setFontSize(18);
    cell.setHorizontalAlignment("center");
    var cellToCopy = sheet.getRange((i+m-2),1);
    var cellToPaste = sheet.getRange((linha),lastColumn+1);
    cellToCopy.copyTo(cellToPaste);
    cellToPaste.setBackground("#34495e"); //Azul escuro
    cellToPaste.setFontColor("white");
    cellToPaste.setFontSize(18);
    cellToPaste.setWrap(true);
    
    time = "";
    linha++;
    somaTotal = 0;
  }
}

/**
* [Converte e soma os horários de uma célula]
* @param  {[Number]} linha    [Número da linha da célula]
* @param  {[Number]} coluna   [Número da coluna da célula]
* @param  {[Object]} planilha [Planilha que contém a célula]
* @return {[Number]}          [Retorna o resultado em inteiro]
*/
function somarHora(linha,coluna,planilha){
  //Renomeando as variáveis
  var i = linha;
  var j = coluna;
  var sheet = planilha;
  
  //Pegando as posições da célula 
  var cell = sheet.getRange(i,j);
  //Pega o valor da célula
  var tString = cell.getValue();
  //Separa as duas partes da string(e.g. "14h00-20h00" -> "14h00", "20h00")
  var resArray = tString.split("-");
  //Verifica se o que tinha na célula seguia a formatação correta(e.g. "20h-21h":TRUE; "20h~21h":FALSE)
  if(resArray[1]== undefined){
    cell.setBackground("#e74c3c"); // vermelho
    return 0; //Pinta de vermelho a célula e retorna zero para não contar nos cálculos
  } else {
    for(var x=0;x<resArray.length;x++) //retira quaisquer espaços nas string(e.g. " 1 4 h " -> "14h")
      resArray[x] = resArray[x].replace(/ /g, "");
    
    var entArray = resArray[0].split("h"); //Retira o 'h' e separa as horas dos minutos
    var saiArray = resArray[1].split("h"); //(e.g. "14h30"-> "14", "30")
    
    var horaEn = parseInt(entArray[0]); //Passa para inteiro o valor das horas
    
    if(isNaN(horaEn)){
      var Aria = entArray[0].replace(/0/, "");
      horaEn = parseInt(Aria);
    }
    
    if(entArray[1]!==''){ //Verifica se foi escrito os minutos. 
      var minEn = parseInt(entArray[1]);
      minEn = minEn/60;
    } else {
      var minEn = 0;
    }
    
    var horaSa = parseInt(saiArray[0]);
    
    if(saiArray[1]!==''){ //Verifica se foi escrito os minutos.
      var minSa = parseInt(saiArray[1]);
      minSa = minSa/60;
    } else {
      var minSa = 0;
    }    
    
    var resultado = (horaSa+minSa)-(horaEn+minEn); //calcula a diferença de horário
    
    if(isNaN(resultado)){ //Se deu para fazer o cálculo então resultado é um número
      cell.setBackground("#e74c3c"); // vermelho
      return 0;
    } else {
      cell.setBackground("#2ecc71"); //verde
      if(resultado<0) //Se foi contado de um dia para o outro o resultado é negativo
        resultado+=24;//Mas basta somar 24 para ter o horário certo
      return resultado;
    }
  }
}

/**
* [Transforma um número decimal em hora]
* @param  {[Number]} decimal [Número que será transformado em horas]
* @return {[String]}         [Horas com toda a formatação(e.g. 13:30)]
*/
function decimalToHour(decimal){
  var minutos = decimal%1; //Retira o decimal, que no caso são os minutos
  decimal -= minutos;
  
  if(decimal<10)
    decimal = "0" + decimal; //Adiciona o zero na frente caso o valor seja menor que 10(e.g 4 -> 04)
  
  minutos *= 60;	
  minutos = Math.round(minutos);	//Transforma em minutos de 0 ~ 60, arredonda e coloca o zero na frente
  if(minutos<10)					//caso necessário
    minutos = "0" + minutos;
  return decimal + ":" + minutos;
}


/* FIM DO CÓDIGO
--------------------------------------------------------------------------------------------
********************************************************************************************
--------------------------------------------------------------------------------------------
FIM DO CÓDIGO */




// Funções criadas para teste
//--------------------------------------------------------------------------------------------

function converteNToL(numero) {
  var aux = parseInt(numero/26);
  var codigos = "";
  if(aux>26){
    return "coluna muito grande";
  }
  var ok = numero%26;
  var alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  if(aux>=1){
    codigos = alfabeto.charAt(aux-1);
  }
  codigos += alfabeto.charAt(ok-1);
  return codigos;
}

function converteLToN(letras) {
  var alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var codigos = [];
  for (var i in letras) {
    codigos.push(alfabeto.indexOf(letras[i].toUpperCase()) + 1);
  }
  return codigos;
}


