var express = require('express');
var router = express.Router();
var mysql = require('mysql');
var fs = require('fs');
var csv = require('fast-csv');

var app = express();

 //var xlsx = require('node-xlsx');  
 
var Excel = require('exceljs');
//var xl = require('excel4node');




/////////////////////////////////////////////////////////////////////////

router.all('/', function(req, res, next) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,GET,OPTIONS,PUT,DELETE");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Accept");
  next();
 });

router.all('/write', function(req, res, next) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,GET,OPTIONS,PUT,DELETE");
   res.setHeader("Access-Control-Allow-Headers", "Content-Type,Accept,X-XSRF-Token");
  next();
 });


router.post('/', function(req, res){
});

router.post('/write', function(req, res){
        
    console.log(req.body.dob);
    console.log(req.body.gender);
    console.log(req.body.salary);
    var dateob = req.body.dob;
    var gender = req.body.gender;
    var sal = req.body.salary;
    
    //=======spliit dob year month day============//
    
    var dobyear = dateob.substring(0, 4);
    console.log('birth year =' + dobyear);
    var dobmonth = dateob.substring(5, 7);
    console.log('birth month =' +dobmonth);
    var dobdate = dateob.substring(8, 10);
    console.log('birth date =' +dobdate);
    
    //===========================================//
    
    //==================calculate age============//
    
    var now = new Date()	
	var age = now.getFullYear() - dobyear
    var curmonth = now.getMonth() + 1
	var mdif = now.getMonth() - dobmonth + 1 //0=jan	
	
	if(mdif < 0)
	{
		--age
	}
	else if(mdif == 0)
	{
		var ddif = now.getDate() - dobdate
		
		if(ddif < 0)
		{
			--age
		}
	}

    console.log(dobmonth);
    console.log(curmonth);
    var roundage = dobmonth - curmonth;
    console.log(roundage);
    if(roundage < 6)
    {
        age = age + 1;
    }
	console.log('age =' +age);
    
    //===========================================//
    
    
    
  var workbook = new Excel.Workbook();
    workbook.xlsx.readFile('MIT.xlsx')
    .then(function(worksheet) {
  
        var worksheet = workbook.getWorksheet('Calculation');
        var worksheet2 = workbook.getWorksheet('Rates');
    
    
        var row = worksheet.getRow(5);
        row.getCell(6).value = sal; // A5's value set to 5
        row.commit();
    
        var rowwW = worksheet.getCell('G28').value;
        //console.log(rowwW);
    
        //===========sum at risk======row 1============//
        var val1 = sal;
        var val2 = worksheet.getCell('C15').value;
        var val3 = worksheet.getCell('C16').value;
        var val4 = worksheet.getCell('C14').value;
        
        var val5 = Math.round(val1*12*val2*(1-Math.pow((1/(1+val3)),val4))/val3);
        console.log(val5);
        //F5*12*C15*(1-(1/(1+C16))^C14)/C16
        //=============================================//
        
        //========sum at risk=======row 2==============//
        var val6 = val1*2*12;
        console.log(val6);
        //F5*2*12
        //=============================================//
    
        //========sum at risk=======row 3==============//
        var val7 = worksheet.getCell('I3').value;
        console.log(val7);
        //I3
        //=============================================//
    
        //========sum at risk=======row 4==============//
        var val8 = worksheet.getCell('I4').value;
        console.log(val8);
        //I4
        //=============================================//
    
        //========sum at risk=======row 5==============//
        var val9 = worksheet.getCell('I5').value;
        console.log(val9);
        //I5
        //=============================================//
    
        //========sum at risk=======row 6==============//
        var val10 = worksheet.getCell('I6').value;
        console.log(val10);
        //I6
        //=============================================//
    
        //========sum at risk=======row 7==============//
        var val11 = worksheet.getCell('I7').value;
        console.log(val11);
        //I7
        //=============================================//
    
        //========sum at risk=======row 8==============//
        var val12 = val1*12*2*(0.1);
        console.log(val12);
        //F5*12*2*10%
        //I8
        //=============================================//
    
        //==================================================================================================//
    
        //===========factor======row 1============//
        var val13 = age // NOT WANTED
        var val14 = worksheet2.getColumn('B');
        //var gender = 'Male';
        var val16;
        val14.eachCell(function(cell, rowNumber) {
            if(cell.value == val13)
            {
                var val15 = worksheet2.getRow(rowNumber);
                if(gender == 'Male')
                {
                val16 = val15.getCell('C').value;
                console.log('male factor row 1 =' +val16);
                }
                else{
                val16 = val15.getCell('D').value;
                console.log(val16);
                }
            }
        });
        
        //VLOOKUP($C$31,Rates!$B$4:$D$58,IF(F4="Female",3,2),FALSE)
        //=============================================//
    
        //===========factor======row 2============//
        var val18 = age; // NOT WANTED
        var val19 = worksheet2.getColumn('B');
        //var gender = 'Male';
        var val21;
        val19.eachCell(function(cell, rowNumber) {
            if(cell.value == val18)
            {
                var val20 = worksheet2.getRow(rowNumber);
                if(gender == 'Male')
                {
                val21 = val20.getCell('C').value;
                console.log(val21);
                }
                else{
                val21 = val20.getCell('D').value;
                console.log(val21);
                }
            }
        });
        
        //VLOOKUP($C$31,Rates!$B$4:$D$58,IF(F5="Female",3,2),FALSE)
        //=============================================//
    
        //===========factor======row 3============//
        var val23 = worksheet2.getColumn('N');
        var val24 = worksheet.getCell('C11').value;
        var val26;
        val23.eachCell(function(cell, rowNumber) {
            if(cell.value == val24)
            {
                var val25 = worksheet2.getRow(rowNumber);
                val26 = val25.getCell('O').value;
                console.log(val26);
                
            }
        });
        
        //VLOOKUP($C$11,Rates!$N$5:$O$11,2,FALSE)
        //=============================================//
    
        //===========factor======row 4============//
        var val27 = age; // NOT WANTED
        var val28 = worksheet2.getColumn('J');
       // var gender = 'Male';
        var val30;
        val28.eachCell(function(cell, rowNumber) {
            if(cell.value == val27)
            {
                var val28 = worksheet2.getRow(rowNumber);
                if(gender == 'Male')
                {
                val30 = val28.getCell('K').value;
                console.log(val30);
                }
                else{
                val30 = val28.getCell('L').value;
                console.log(val30);
                }
            }
        });
        
        //VLOOKUP($C$31,Rates!$J$5:$L$53,IF(F5="Female",3,2),FALSE)
        //=============================================//
    
        //===========factor======row 5============//
        var val32 = worksheet2.getColumn('Q');
        var val33 = worksheet.getCell('C11').value;
        var val35;
        val32.eachCell(function(cell, rowNumber) {
            if(cell.value == val24)
            {
                var val34 = worksheet2.getRow(rowNumber);
                val35 = val34.getCell('R').value;
                console.log(val35);
                
            }
        });
        
        //VLOOKUP($C$11,Rates!$Q$5:$R$11,2,FALSE)
        //=============================================//
    
        //===========factor======row 6============//
        var val36 = age; // NOT WANTED
        var val37 = worksheet.getCell('C5').value;
        var val38 = worksheet2.getColumn('AI');
        //var gender = 'Male';
        var val40;
        if(val37 == '39 Illnesses'){
            
            val38.eachCell(function(cell, rowNumber) {
            if(cell.value == val36)
            {
                var val39 = worksheet2.getRow(rowNumber);
                if(gender == 'Male')
                {
                val40 = val39.getCell('AJ').value;
                console.log(val40);
                }
                else{
                val40 = val39.getCell('AK').value;
                console.log(val40);
                }
            }
        })
        }
        
        //VLOOKUP(Calculation!C31,IF(C5="39 Illnesses",Rates!$AI$6:$AK$54,Rates!$AN$6:$AP$53),IF(F4="Female",3,2),FALSE)
        //=============================================//
    
        //===========factor======row 7============//
        var val42 = age; // NOT WANTED
        var val43 = worksheet2.getColumn('T');
        //var gender = 'Male';
        var val45;
        val43.eachCell(function(cell, rowNumber) {
            if(cell.value == val42)
            {
                var val44 = worksheet2.getRow(rowNumber);
                if(gender == 'Male')
                {
                val45 = val44.getCell('U').value;
                console.log(val45);
                }
                else{
                val45 = val44.getCell('V').value;
                console.log(val45);
                }
            }
        })
       
        
        //VLOOKUP(C31,Rates!$T$5:$V$71,IF(F4="Female",3,2),FALSE)
        //=============================================//
        
        //===========factor======row 8============//
        var val47 = age; // NOT WANTED
        var val48 = worksheet2.getColumn('B');
        //var gender = 'Male';
        var val50;
        val48.eachCell(function(cell, rowNumber) {
            if(cell.value == val47)
            {
                var val49 = worksheet2.getRow(rowNumber);
                if(gender == 'Male')
                {
                val50 = val49.getCell('C').value;
                console.log(val50);
                }
                else{
                val50 = val49.getCell('D').value;
                console.log(val50);
                }
            }
        })
       
        
        //VLOOKUP($C$31,Rates!$B$4:$D$58,IF(F11="Female",3,2),FALSE)
        //=============================================//
        
        //=======================================================================================//
        
        //===========risk loading======row 1 2 3 4 5 6 7 8============//
        var val52 = worksheet2.getColumn('AB');
        var val53 = worksheet.getCell('C11').value;
        var val555;
        val52.eachCell(function(cell, rowNumber) {
            if(cell.value == val53)
            {
                var val54 = worksheet2.getRow(rowNumber);
                var val55 = ((val54.getCell('AC').value)*100).toFixed(2) + '%';
                val555 = (((val54.getCell('AC').value)*100).toFixed(2))/100;
                console.log(val55);
                
            }
        });
       
        
        //VLOOKUP($C$11,Rates!$AB$5:$AC$12,2,FALSE)+$C$12
        //=============================================//
    
        //=======================================================================================//
    
        //===========company loading======row 1 2 3 4 5 6 7 8============//
        var val56 = worksheet.getCell('C7').value;
        var val57 = worksheet.getCell('C8').value;
        var val58 = worksheet.getCell('C9').value;
        var val59 = ((val56+val57+val58)*100)+ '%';
        var val599 = (((val56+val57+val58)*100).toFixed(2))/100;
        
                console.log(val59);      
        
        //SUM($C$7:$C$9)
        //=============================================//
    
        //=====================================================================================//
    
        //===========premium======row 1============//
            
        var val60 = ((val5*val16*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val60);         
        
        //(C34*D34*(1+E34))/(1-F34)
        //=============================================//
    
        //===========premium======row 2============//
            
        var val61 = ((val6*val21*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val61);         
        
        //(C35*D35*(1+E35))/(1-F35)
        //=============================================//
    
        //===========premium======row 3============//
            
        var val62 = ((val7/1000*val26*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val62);         
        
        //(C36/1000*D36*(1+E36))/(1-F36)
        //=============================================//
    
        //===========premium======row 4============//
            
        var val63 = ((val8/1000*val30*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val63);         
        
        //(C37/1000*D37*(1+E37))/(1-F37)
        //=============================================//
    
        //===========premium======row 5============//
            
        var val64 = ((val9/1000*val35*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val64);         
        
        //(C38/1000*D38*(1+E38))/(1-F38)
        //=============================================//
    
        //===========premium======row 6============//
            
        var val65 = ((val10/1000*val40*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val65);         
        
        //(C39/1000*D39*(1+E39))/(1-F39)
        //=============================================//
    
        //===========premium======row 7============//
            
        var val66 = ((val11/100*val45*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val66);         
        
        //(C40/100*D40*(1+E40))/(1-F40)
        //=============================================//
    
        //===========premium======row 8============//
            
        var val67 = ((val12*val50*(1+Number(val555)))/(1-Number(val599))).toFixed(2);
        console.log(val67);         
        
        //(C41*D41*(1+E41))/(1-F41)
        //=============================================//
    
        //============================premium per year====================================================//
    
        var val68 = (Number(val60)+Number(val61)+Number(val62)+Number(val63)+Number(val64)+Number(val65)+Number(val66)+Number(val67)).toFixed(2);
        console.log("val = ",val68);
        
        
    
        
    
       // workbook.xlsx.writeFile('new.xlsx');
        
    
        /*var roww = worksheet.getCell('F5').value = 50000;
            console.log(roww);
        var lol = worksheet.getCell('F5').value;
            console.log(lol);
        var rowwW = worksheet.getCell('G28').value;
            console.log(rowwW);
            
             var row = worksheet.getRow(5);
        row.getCell(6).value = 50000; // A5's value set to 5
        row.getCell(7).value = 00000;
        row.commit();
        workbook.xlsx.writeFile('new.xlsx');
        
    
        worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
            if(rowNumber == 3){
            row.eachCell(function(cell, colNumber) {
                if(colNumber == 6){
                    console.log('dob =' + cell.value);
                }
  
                });
            }
            if(rowNumber == 4){
            row.eachCell(function(cell, colNumber) {
                if(colNumber == 6){
                    console.log('gender =' + cell.value);
                }
  
                });
            }
            if(rowNumber == 5){
            row.eachCell(function(cell, colNumber) {
                if(colNumber == 6){
                    console.log('salary =' + cell.value);
                }
  
                });
            }
            
        });
*/
        
    var Annuity_Benefit = sal;
        console.log(Annuity_Benefit);
        
    var Death_Benefit = sal*12*2;
        console.log(Death_Benefit);
        
    var val70 = worksheet.getCell('I3').value;
    var Accidental_Death_Benefit = (sal*12*2)+val70;
        console.log(Accidental_Death_Benefit);
        
    //var val71 = worksheet.getCell('I4').value;
    var val71 = sal*12;
    var Total_and_Permanent_Disability = val71;
        console.log(Total_and_Permanent_Disability);
        
    //var val722 = worksheet.getCell('I5').value;
    var val722 = sal*12;
    var Partial_and_Permanent_Disability = val722;
        console.log(Partial_and_Permanent_Disability);
        
    //var val72 = worksheet.getCell('I6').value;
    var val72 = (sal*12)*(0.5);
    
    var new_critical_ill;
        if(val72 > 3000000){
            new_critical_ill = 3000000;
        }else{
            new_critical_ill = val72;
        }
    var Critical_Illness = val72;
        console.log("ill = ",Critical_Illness);
        
    var val73 = worksheet.getCell('I7').value;
    var Hospitalization_Benefit = val73;
        console.log(Hospitalization_Benefit);
        
    var val74 = sal*12*2*(0.1);
    var new_funeral_ex;
    console.log("fe = ",val74);
        if(val74 > 500000){
            new_funeral_ex = 500000;
        }else{
            new_funeral_ex = val74;
        }
    var Funeral_Expenses = new_funeral_ex;
        console.log("fe = ",Funeral_Expenses);
        
    if(age<18 || age>65)
    {
        var errorrs = '888';
        console.log('small age');
        res.send(errorrs);
    }
    else
    {   
    
    if(val68 > 0)
    {
        console.log('val sent');
        //res.send(val68);
        res.send(JSON.stringify({ a: Annuity_Benefit,b: Death_Benefit,c: Accidental_Death_Benefit,d: Total_and_Permanent_Disability,e: Partial_and_Permanent_Disability,f: Critical_Illness,g: Hospitalization_Benefit,h: Funeral_Expenses,i: val68 }));
    }
    }
        
    });
 
    
    
    
    
});

router.get('/write', function(req, res){
  
});


module.exports = router;