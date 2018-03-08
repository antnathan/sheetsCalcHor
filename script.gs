function somando(valores){
  var soma = 0;
  var somaTotal = 0;
  var somaVec = [];
  for(var i = 0; i<valores.length; i++ ){
    for( var j = 0; j<valores[i].length ;j++)
      soma+= somarHora(valores[i][j]);
    somaVec.push(soma);
    somaTotal+=soma;
    soma = 0;
  }
  var somaFormatada = decimalToHour(somaTotal);
  return somaTotal/24;
}


function somarHora(tString){
  
  //Separa as duas partes da string(e.g. "14h00-20h00" -> "14h00", "20h00")
  var resArray = tString.split("-");
  //Verifica se o que tinha na célula seguia a formatação correta(e.g. "20h-21h":TRUE; "20h~21h":FALSE)
  if(resArray[1]== undefined){
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
      return 0;
    } else {
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

function proximaLin(cel1, cel2){
  var aux = cel1.split("");
  var letra = aux[0];
  aux = cel2.split("");
  if(aux.length>1)
    var num = aux[1]+aux[2];
  else
    var num = aux[1];
  
  num = Number(num);
  num+=1;
            
  
  return letra+num;
}


function convertA1ToNumber(celA1Not){
  var aux = celA1Not.split("");
  var col = converteLToN(aux[0]);
  var lin = aux[1];
  
  var vec = [col[0],lin];
  return vec;
  
}

function convertNumberToA1(colNumNot, linNumNot){
  var letra = convertNToL(colNumNot);
  var num = linNumNot;
  
  return letra+num;
}

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

function converteLToN(letra) {
  var alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var codigos = [];
  codigos.push(alfabeto.indexOf(letra.toUpperCase()) + 1);
  return codigos;
}