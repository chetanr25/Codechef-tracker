var list = [];
// Please fill the ID of the spreadsheet below
var ssID = ""; // ID of spreadsheet
var phoneNo = '..........' // phone number  
var sheet = SpreadsheetApp.openById(ssID);
function myFunction() {
  var ss = sheet.getSheetByName("CodeChef"); //accessing spreadsheet
  var ssWW = sheet.getSheetByName("CodeChef Week Wise");
  var cols = ss.getLastColumn()-1;
  var nRows = ss.getLastRow()+1;
  var lst = loopFetch(cols,ss);
  mkList(lst,ss,ssWW,cols,nRows);
  var date = updateTotal(ss,nRows,cols,lst);
  updateWeek(ssWW,cols);
  mailService(cols,date);
}


// making a list of members with their data and mail id
function mkList(lst,ss,ssWW,cols,nRows){
  for(var l=0; l<cols; l++){
    list.push({'name': ss.getRange(1,2+l).getValue(), 
    'questions': parseInt(lst[l]-ss.getRange(nRows-1,2+l).getValue()),
    'mail':ss.getRange(2,2+l).getValue(),
    'prevTotal':parseInt(ss.getRange(nRows-1,2+l).getValue()),
    'prevQuestions':parseInt(ssWW.getRange(ssWW.getLastRow(),2+l).getValue()),
    'total':lst[l],
    });
  }
}

// sending mail to each member
function mailService(cols,date){
    var staticsSsPS = sheet.getSheetByName("Analysis"); //accessing spreadsheet
    var leaderboardSS = sheet.getSheetByName("Leaderboard"); //accessing spreadsheet
    var noWeeks = staticsSsPS.getRange(1,staticsSsPS.getLastColumn());//.getValue();
    noWeeks.setValue(noWeeks.getValue()+1);
    noWeeks=noWeeks.getValue();
  var firstRow = staticsSsPS.getRange(1,2,1,staticsSsPS.getLastColumn()-1).getValues()[0];
  // construction of leaderboard and variables for statistics data.
  let str = "LEADERBOARD\n";
    let zerosNames = '';
    let zero = 0;
    let ms = "";
    var abovelast = 0;
    var belowlast = 0;
    var aboveLastName = [];
    var belowLastName = [];
    var aboveavg = 0;
    var belowavg = 0;
    var aboveAvgName = [];
    var belowAvgName = [];
  list.sort(function(a, b) { return ((a.total < b.total) ? 0 : -1);});
  list.sort(function(a, b) { return ((a.questions < b.questions) ? 0:-1);});
  // leaderboard

  // Generate the leaderboard string
  for(var i = 0; i<list.length; i++){
    if(parseInt(list[i].questions)!=0)
    str+= (i+1)+". " + list[i].name + " " + list[i].questions + "\n";
    else {
        // Update the zeros in the statics sheet
      var targetCol = firstRow.indexOf(list[i].name)+2;
      staticsSsPS.getRange(2,targetCol).setValue(staticsSsPS.getRange(2,targetCol).getValue()+1);
      cols--;
      zerosNames+=list[i].name+'\n';
      zero++;
    }
  }
  // Generating clickable whatsapp link for the leaderboard
  var waURL = 'https://wa.me/91'+phoneNo+'?text='+encodeURIComponent(str);

  var richValue = SpreadsheetApp.newRichTextValue()
   .setText("Week "+parseInt(leaderboardSS.getLastRow()))
   .setLinkUrl(waURL)
   .build();
//    Logger.log(str);
// Add the leaderboard row to the spreadsheet

  leaderboardSS.getRange(leaderboardSS.getLastRow()+1,1).setRichTextValue(richValue);
  let avgList = parseInt(avg(list)); 
  for(var l=0; l<cols; l++){
    leaderboardSS.getRange(leaderboardSS.getLastRow(),l+2).setValue(list[l].name);
    var targetCol = firstRow.indexOf(list[l].name)+2;
    let ps = str;
    ps+="\nFinished in ";

  if(l==0){ 

      ps+= (l+1) +"st place!";
      staticsSsPS.getRange(3,targetCol).setValue(staticsSsPS.getRange(3,targetCol).getValue()+1);
      staticsSsPS.getRange(4,targetCol).setValue(staticsSsPS.getRange(4,targetCol).getValue()+1)
    }
    else if(l==1){
      ps+= (l+1) +"nd place!";
      staticsSsPS.getRange(4,targetCol).setValue(staticsSsPS.getRange(4,targetCol).getValue()+1)
      }
    else if(l==2){
      ps+= (l+1) +"rd place!";
      staticsSsPS.getRange(4,targetCol).setValue(staticsSsPS.getRange(4,targetCol).getValue()+1)
      }
    else if(l==cols-1){
      ps+= "in last place";
      staticsSsPS.getRange(5,targetCol).setValue(staticsSsPS.getRange(5,targetCol).getValue()+1)
      }
    else ps+= (l+1) +"th place"
    ps+="\nAverage problems solved by our squad =" + avgList ;
    ps+="\n\nPracticed "+list[l].questions + " questions this week";
    ps+="\nTotal number of solved questions on CodeChef = " + list[l].total;
    if(list[l].questions==0)
    continue;
    
    var zerosStatic = staticsSsPS.getRange(2,targetCol).getValue();
    // Rank
    var tempo = ((staticsSsPS.getRange(6,targetCol).getValue()*(noWeeks-zerosStatic-1))+(l+1))/(noWeeks-zerosStatic);
    staticsSsPS.getRange(6,targetCol).setValue(tempo.toFixed(2));
    if(list[l].questions > avgList){
        // Above average
      staticsSsPS.getRange(7,targetCol).setValue(parseInt(staticsSsPS.getRange(7,targetCol).getValue()+1));
      ps=ps.concat("\nCongrats, You have solved ",parseInt(list[l].questions-avgList)," more than the average!\n");
      aboveavg++;
      aboveAvgName.push({'name':list[l].name,'no':parseInt(list[l].questions-avgList)});
    }
    else {
        // Below average
      staticsSsPS.getRange(8,targetCol).setValue(parseInt(staticsSsPS.getRange(8,targetCol).getValue()+1))
      ps=ps.concat("\nYou have solved ",Math.abs(list[l].questions-avgList)," less than average!\n");
      belowavg++;
      belowAvgName.push({'name':list[l].name,'no':parseInt(list[l].questions-avgList)});
    }
    
    if(list[l].prevQuestions>list[l].questions && list[l].prevQuestions!=NaN){
        // Decline in number of questions practiced
      staticsSsPS.getRange(10,targetCol).setValue(parseInt(staticsSsPS.getRange(10,targetCol).getValue()+1));
      ps=ps.concat("\nDecline in number of question practiced by ",(list[l].prevQuestions-list[l].questions)," problems than last week.");
      belowLastName.push({'name':list[l].name,'no':parseInt(list[l].questions-list[l].prevQuestions)});
      belowlast++;
    }
    else if(list[l].prevQuestions<list[l].questions && list[l].prevQuestions!=NaN){
        // Improvement in number of questions practiced
      staticsSsPS.getRange(9,targetCol).setValue(parseInt(staticsSsPS.getRange(9,targetCol).getValue()+1));
      ps=ps.concat("\nImproved than last week by ",list[l].questions-list[l].prevQuestions," problems. Good job!");
      abovelast++;
      aboveLastName.push({'name':list[l].name,'no':parseInt(list[l].questions-list[l].prevQuestions)});
    }
    var sub = "Hi, "+list[l].name+" here's an update regarding codechef  - " + date;
    
    // sending personalized mail to each member after analysis of their data
    MailApp.sendEmail(list[l].mail,sub,ps); 
    // break;
  }

  // statistics 
 ms="Total mail's sent = " + cols;
  if(zero!=NaN)
  ms+="\nTotal members who didn't solve any questions this week = " + zero;
  if(avgList!=NaN)
  ms+="\n\nThis week's Average problems solved = "+avgList;
  if(aboveavg!=NaN)
  ms+="\nTotal members who are above average = "+aboveavg;
  if(belowavg!=NaN)
  ms+="\nTotal members who are below average = "+belowavg;
  if(abovelast!=NaN)
  ms+="\n\nTotal members who solved more than last week = "+abovelast;
  if(belowlast!=NaN)
  ms+="\nTotal members who solved less than last week = "+belowlast;
  if(zero!=NaN)
  ms+="\n\nMembers who didn't care to solve even a single problem :\n"+ zerosNames;
  if(aboveavg!=NaN)
  ms+="\nMembers who solved more than average :\n" + srtLst(aboveAvgName);
  if(belowavg!=NaN)
  ms+="\nMembers who solved less than average :\n" + srtLst(belowAvgName);
  if(abovelast!=NaN)
  ms+="\nMembers who solved more compare to last week :\n" + srtLst(aboveLastName);
  if(belowlast!=NaN)
  ms+="\nMembers who solved less compare to last week :\n" + srtLst(belowLastName);
  MailApp.sendEmail('chetan250204@gmail.com',"Analysis ",ms); 
  if(cols>2)
  MailApp.sendEmail(list[0].mail,"Congrats for finishing 1st, Here's a treat for you","Analysis of this week's codechef data", ms); 

  var staticsSS = sheet.getSheetByName("Analysis-weekWise"); //accessing spreadsheet
  var rw = staticsSS.getLastRow();

  // Generating clickable whatsapp link for the statistics
  waURL = 'https://wa.me/91'+phoneNo+'?text='+encodeURIComponent(ms);
  var richValue = SpreadsheetApp.newRichTextValue()
   .setText("Week "+rw)
   .setLinkUrl(waURL)
   .build();
  
   // Add the statistics row to the spreadsheet
  staticsSS.appendRow([rw,cols,zero,avgList,aboveavg,belowavg,abovelast,belowlast]);
  staticsSS.getRange(staticsSS.getLastRow(),1).setRichTextValue(richValue);
}

// sorting the list of members according to their number of questions
function srtLst(data1){
  data1.sort(function(a, b) { return ((Math.abs(a.no) < Math.abs(b.no)) ? -1 : 0);});
  var str = '';
  for(var x = 0; x<data1.length; x++){
    str+=data1[x].name +" ("+data1[x].no+")\n";
  }
  return str;
}


// updating the total sheet
function updateTotal(ss,nRows,cols,lst){
  var date = new Date();
  date = date.toDateString().slice(4);
  ss.getRange(nRows,1).setValue(date);
  for(var l=0; l<cols; l++){
    ss.getRange(nRows,2+l).setValue(lst[l]);
  }
  return date;
}

// updating the week wise sheet
function updateWeek(week,cols){
  var nRows = week.getLastRow()+1;
  week.getRange(nRows,1).setValue(week.getRange(nRows-1,1).getValue()+1)
  for(var l=0; l<cols; l++){
    week.getRange(nRows,2+l).setValue(list[l].questions);
  }
}

// fetching the data from the link
function fetchit(str){//Logger.log(str);
  var response = UrlFetchApp.fetch(str); // link, data scraped...
  var content = (response.getContentText()); // convert(parse) to HTML
  var $ = Cheerio.load(content); // init HTML to cheerio(some random API)
  var item = $('.content').eq(0).find("h5").eq(0).text(); //extracting no. of practiced content
  item = item.slice(14,item.length-1); // trimming unwanted data and extract only number from it
  return item;
}

function avg(array) {
    let sum = 0;
    let n = 0;
    for (let i = 0; i < array.length; i++) {
        if(array[i].questions!=0)
          n++;
        sum += array[i].questions;
    }
    if(n)
    return sum / n;
    else
    return 0;
}

// looping through the spreadsheet and fetching the data
function loopFetch(cols,ss){
  var lst=[];
  for(var l=0; l<cols; l++){
    lst.push(fetchit(ss.getRange(1,2+l).getRichTextValue().getLinkUrl()));
  }
  return lst;
}