  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var depositS = ss.getSheetByName("Deposit");
  var dataS = ss.getSheetByName("Data");
  var values   = ss.getSheetByName("Data").getDataRange().getValues();
  var formS         = ss.getSheetByName("Form");
  var installS      = ss.getSheetByName("Installment");
  var reportS       = ss.getSheetByName("Personal_Report");
  let dic = {
    A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9
  };

function ClearCell() {

  var rangesToClear = ["D7", "D9", "D11", "D13", "D15", "D17", "D19"];
  for (var i=0; i<rangesToClear.length; i++){
    formS.getRange(rangesToClear[i]).clearContent();
  }
}

//Input New Data

function SubmitData(){
  var newRow        = dataS.getLastRow()+1;
  var values        = [formS.getRange("D7").getValue(),
                        formS.getRange("D9").getValue(),
                        formS.getRange("D11").getValue(),
                        formS.getRange("D13").getValue(),
                        formS.getRange("D15").getValue(),
                        formS.getRange("D17").getValue(),
                        formS.getRange("D19").getValue(),];

for (var i=0; i<values.length; i++){
  if ( i == 4){
    dataS.getRange(newRow, i+1).setValue(values[i]);
  }
  else{
    dataS.getRange(newRow, i+1).setValue(values[i]);
  }
  }
  installS.getRange(newRow,1).setValue(values[0]);
  installS.getRange(installS.getLastRow(), 2).setValue(values[1]);
  installS.getRange(installS.getLastRow(), 3).setValue(values[4]);

  dataS.getRange(dataS.getLastRow(),8).setValue(0);
  dataS.getRange(dataS.getLastRow(),9).setValue(0);

  ClearCell();
}
//--------------------------------------------------------

//Count the number of installment given by an individual
function countInstallment(text){

var count_installment; 
var installments = ss.getSheetByName("Installment").getDataRange().getValues();

  for (var i =0; i < installments.length; i++){
        var row_installment = installments[i];
        if (row_installment[2] == text){
          var j = 3;
          var k = 1;
          for (k =1; k <= 200 ; k++){
            if (row_installment[j] > 1){
              j = j+4;
            }
            else {
              k = 201;
            }
          }
          count_installment = ((j-3)/4);
          }
          
    }
   return count_installment;
}
// Count Days
// function sudcount(a,b,bokeya_ashol,kistinum){
//   tarikh_1 = new Date();
//   tarikh_2 = 
//   difference = tarikh_1.getTime()-b.getTime
//   msInDay = 1000 * 3600 * 24;
//   days = difference/msInDay;
//   depositS.getRange("I10").setValue(days);
//   depositS.getRange("I11").setValue(days-(90*msInDay));
//   if (kistinum = 0 ){
//     let bokeya_shud = (days- (90*msInDay))*(bokeya_ashol/7300);
//     return bokeya_shud;
//   }

// }

// Search

var SERACH_COL_IDX = 4;
function Search(){

  var str      = depositS.getRange("D2").getValue();
  var values   = ss.getSheetByName("Data").getDataRange().getValues();
  var kisti = countInstallment(str);

  for (var i =0; i < values.length; i++){
    var p = (kisti*4)+2;
    var row = values[i];
    if (row[SERACH_COL_IDX] == str) {
      
      
      depositS.getRange("D4").setValue(row[1]);
      depositS.getRange("D5").setValue(row[3]);
      depositS.getRange("I4").setValue(row[0]);
      depositS.getRange("D7").setValue(row[7]);
      depositS.getRange("D9").setValue(row[8]);
      depositS.getRange("D13").setValue(row[9]);
      depositS.getRange("D17").setValue(row[13]);
      depositS.getRange("D21").setValue(new Date());
      depositS.getRange("D11").setValue(countInstallment(str));
      if (kisti == 0){
      depositS.getRange("D19").setValue(row[6]);
      }
      else {
        
        var installments = ss.getSheetByName("Installment").getDataRange().getValues();

       for (var i =0; i < installments.length; i++){
        var row_installment = installments[i];
        if (row_installment[2] == str){
        depositS.getRange("D19").setValue(row_installment[p]);
      }
      }
    }
      tarikh_1 = new Date();
      tarikh_2 = row[6];
      bb = row[9]
  difference = tarikh_1.getTime()-tarikh_2.getTime()
  msInDay = 1000 * 3600 * 24;
  days = difference/msInDay;
  if (kisti == 0 ){
    bokeya_shud = (days - 90)*(bb/7300);
    depositS.getRange("D15").setValue(bokeya_shud); 
    dataS.getRange(i+1,11).setValue(bokeya_shud);
  }
  else {
    tarikh_3 = row_installment[p];
    difference_2 = tarikh_1.getTime() - tarikh_3.getTime();
    days_2 = difference_2/msInDay ;
    prev_bokeyaSud = row[10];
    bokeya_shud = prev_bokeyaSud + (days_2 * (bb/7300));
    depositS.getRange("D15").setValue(bokeya_shud); 
    dataS.getRange(i,11).setValue(bokeya_shud);
  }     
    }

  }
}
function input_install(a,b,c,d){
  var str      = depositS.getRange("D2").getValue();
  var installments = ss.getSheetByName("Installment").getDataRange().getValues();

      for (var i =0; i < installments.length; i++){
        var row_installment = installments[i];
        if (row_installment[2] == str){
            var j = 4;
            var k = 3;
            for (var k = 3 ; k < 200 ; k++){
              
            if (row_installment[j] > 1 ){
              j = j + 4;
            }
            else {
          installS.getRange(i+1,j).setValue(a).setNumberFormat(d);
          installS.getRange(i+1,j+1).setValue(b).setNumberFormat(d);
          installS.getRange(i+1,j+2).setValue(c).setNumberFormat(d);
          installS.getRange(i+1,j+3).setValue(new Date());
          installS.getRange(3,j).setValue("Inastallment (Principal)");
          installS.getRange(3,j+1).setValue("Installment (Interest)");
          installS.getRange(3,j+2).setValue("Savings");
          installS.getRange(3,j+3).setValue("Date");
          count = ((j-4)/4)+1 ;
          installS.getRange(2,j).setValue(count);
          installS.getRange(2,j,1,4).merge();
          installS.getRange(1,4).setValue("Installment Details");
          installS.getRange(1,4,1,installments[3].length+1).merge();
          k = 201;
        }
        
        }
        depositS.getRange("I12").clearContent();
        depositS.getRange("I14").clearContent();
        depositS.getRange("I16").clearContent();
   }
}
}

// Make Installment;

function make_installment(){
  var str      = depositS.getRange("D2").getValue();
  given_capital = depositS.getRange("I12").getValue();
  var a = depositS.getRange("I12").getNumberFormat();
  given_int = depositS.getRange("I14").getValue();
  given_dep = depositS.getRange("I16").getValue();
  for (var i = 0 ; i < values.length ; i++){
    var row = values[i];
    if (row[4] == str) {
     prev_capital = row[7];
     prev_int = row[8];
     prev_dep = row[12];
      bokeya_sud = row[10];
      new_bokeyaSud = bokeya_sud - given_int;
     new_capital = given_capital + prev_capital;
     new_int = given_int + prev_int;
     new_dep = given_dep + prev_dep;
     
    dataS.getRange(i+1,8).setValue(new_capital);
    dataS.getRange(i+1,9).setValue(new_int);
    dataS.getRange(i+1,13).setValue(new_dep);
    dataS.getRange(i+1,11).setValue(new_bokeyaSud);
    
    input_install(given_capital,given_int,given_dep,a);
    Search();
   
    
    }

}
}


//Report

function report(){
  Search();
    var str      = depositS.getRange("D2").getValue();
  var values   = ss.getSheetByName("Data").getDataRange().getValues();

  for (var i =0; i < values.length; i++){
    var row = values[i];
    if (row[4] == str) {
    reportS.getRange("D12").setValue(row[1]);
    reportS.getRange("D14").setValue(row[2]);
    reportS.getRange("D16").setValue(row[3]);
    reportS.getRange("D18").setValue(row[4]);
    reportS.getRange("D23").setValue(row[5]);
    reportS.getRange("D25").setValue(row[6]);
    reportS.getRange("D29").setValue(row[7]);
    reportS.getRange("D31").setValue(row[8]);
    reportS.getRange("D35").setValue(row[9]);
    reportS.getRange("D37").setValue(row[10]);
    reportS.getRange("D41").setValue(row[11]);
    reportS.getRange("D43").setValue(row[12]);
    reportS.getRange("D45").setValue(row[13]);
    t1 = depositS.getRange("D19").getValue();
    reportS.getRange("D48").setValue(t1);
    reportS.getRange("D50").setValue(new Date());
    reportS.getRange("F9").setValue(row[0]);
    }
}
}







