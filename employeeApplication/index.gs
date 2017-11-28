//������
function doGet(e) {
  if (e!=undefined){
    var parameters = e.parameter;
    var idFrom = parameters.idFrom;
    var nameFrom = parameters.nameFrom;
    var firstSendTo = parameters.firstSendTo;
    var secondSendTo = parameters.secondSendTo;
    var thirdSendTo = parameters.thirdSendTo;
    var pass = parameters.pass;
    var count = parameters.count;
    var begindate = parameters.begindate;
    var enddate = parameters.enddate;
    var hours = parameters.hours;
    var reason = parameters.reason;
    var type = parameters.type;
    var qjr = parameters.qjr;
    var qjrname = parameters.qjrname;
        
    Logger.log("idFrom:"+idFrom)
    Logger.log("nameFrom:"+nameFrom)
    Logger.log("firstSendTo:"+firstSendTo)
    Logger.log("secondSendTo:"+secondSendTo)
    Logger.log("thirdSendTo:"+thirdSendTo)
    Logger.log("begindate:"+begindate)
    Logger.log("enddate:"+enddate)
    Logger.log("hours:"+hours)
    Logger.log("reason:"+reason)
    Logger.log("type:"+type)
    Logger.log("qjr:"+qjr)
    Logger.log("qjrname:"+qjrname)
    
    Logger.log("parameters:"+parameters)
    if (idFrom){
      var _pass = pass=='1'?"ͨ��":"�ܾ�"
      setAppStatus(idFrom, count, _pass, qjr, hours, type)
      sendEmail(idFrom,nameFrom,firstSendTo,secondSendTo,thirdSendTo,count, begindate,enddate,hours,reason,type,qjr,qjrname)      
      return HtmlService.createHtmlOutput("��"+_pass+"�����ˣ�");
    }     
  } 
  return HtmlService.createHtmlOutputFromFile('Index');
}

//�����ʼ�
//idFrom:������Id
//nameFrom������������
//firstSendTo:��һ����������
//secondSendTo���ڶ�����������
//thirdSendTo����������������
//count����ǰ�µڼ��η����ʼ�
function sendEmail(idFrom,nameFrom,firstSendTo,secondSendTo,thirdSendTo,count, begindate,enddate,hours,reason,type,qjr,qjrname) {
  Logger.log("sendEmail")
   var _count=parseInt(count)+1
   var _href = "href='https://script.google.com/macros/s/AKfycbyJMJkalBj5VpUqCSHH_g5ou1bdSO8Vo5SYiXUh3nQ/dev?"+
                  "idFrom="+idFrom+
                  "&nameFrom="+nameFrom+
                  "&firstSendTo="+firstSendTo+
                  "&secondSendTo="+secondSendTo+
                  "&thirdSendTo="+thirdSendTo+  
                  "&count="+_count+
                  "&begindate="+begindate+
                  "&enddate="+enddate+
                  "&hours="+hours+
                  "&reason="+reason+
                  "&type="+type+
                  "&qjr="+qjr+
                  "&qjrname="+qjrname
   var _approve = "&pass=1"
   var _reject  = "&pass=0"                   
   var htmlBody = "<span style='color:red'>�������</span><br>"+
                  "<span>������������</span>"+nameFrom+"<br>"+
                  "<span>�����������</span>"+qjrname+"<br>"+
                  "<span>��ʼʱ�䣺</span>"+begindate+"<br>"+
                  "<span>����ʱ�䣺</span>"+enddate+"<br>"+
                  "<span>��ʱ��</span>"+hours+"<br>"+
                  "<span>������ͣ�</span>"+type+"<br>"+
                  "<span>���ɣ�</span>"+reason+"<br>"+                    
                  "<a "+_href + _approve + "'>ͨ��</a>" + 
                  "<a "+_href + _reject +"' style='margin-left:50px'>�ܾ�</a>"
                 
   var sendTo = ''
   if (_count ==1){
     sendTo = firstSendTo
   } else if (_count ==2){
     sendTo = secondSendTo
   } else if (_count==3){
     sendTo = thirdSendTo
   } else {
     return 
   }
   MailApp.sendEmail({
     to: sendTo,
     subject: "���ԡ�"+nameFrom+"�����������",
     htmlBody: htmlBody
    
   });  
}


//��ȡ��ǰ��¼�˵���Ϣ
function getActiveUser(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var user_data = sheet.getSheets()[0].getDataRange().getValues();
  var login_user=Session.getActiveUser().getEmail();
  var r = {};
   for (var i =1; i<user_data.length;i++){
    if (user_data[i][0]==login_user){
      r.email = user_data[i][0]
      r.no = user_data[i][1]      
      r.name = user_data[i][2] 
      r.part = user_data[i][3] 
      r.job = user_data[i][4] 
      r.firstSpr = user_data[i][5] 
      r.secondSpr = user_data[i][6] 
      r.thirdSpr = user_data[i][7] 
      r.allHours = user_data[i][8] 
      r.syHours = user_data[i][9] 
      break;
    }
  }  
  return r;
}

//����ID��ȡ�û�����Ϣ
function getUserById(id){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var user_data = sheet.getSheets()[0].getDataRange().getValues();
  var r = {};
   for (var i =1; i<user_data.length;i++){
    if (user_data[i][0]==id){
      r.email = user_data[i][0]
      r.no = user_data[i][1]      
      r.name = user_data[i][2] 
      r.part = user_data[i][3] 
      r.job = user_data[i][4] 
      r.firstSpr = user_data[i][5] 
      r.secondSpr = user_data[i][6] 
      r.thirdSpr = user_data[i][7] 
      r.allHours = user_data[i][8] 
      r.syHours = user_data[i][9] 
      break;
    }
  }  
  return r;
}

//��ȡ�����û��б�
function getUserList(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var user_data = sheet.getSheets()[0].getDataRange().getValues();
  var r = [];
   for (var i =1; i<user_data.length;i++){
     r.push({
       email: user_data[i][0],     
       name: user_data[i][2] 
     });   
  }  
  return r;
}

//���ڸ�ʽ����yyyy-MM-dd hh:mm:ss
function dateFormat(value){
  if (value==''){return ''}
  function add0(v){
    return v<10?'0'+v:v
  }
  var time = value
  var y = time.getFullYear();
  var m = time.getMonth()+1;
  var d = time.getDate();
  var h = time.getHours();
  var mm = time.getMinutes();
  var s = time.getSeconds();
  return y+'-'+add0(m)+'-'+add0(d)+' '+h+':'+add0(mm)+'-'+add0(s)
}

//��ȡ��ǰ��¼�˵������¼
function logProductInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getSheets()[1].getDataRange().getValues();
  var user_date = sheet.getSheets()[0].getDataRange().getValues();
  var r = {
    data: [],
    maxId: 0
  };
  var login_user=Session.getActiveUser().getEmail();
  var user_part = "";
  //getEmail(login_user);
  for (var i =1; i<user_date.length;i++){
    if (user_date[i][0]==login_user){
      user_part = user_date[i][3]
      break;
    }
  }
  var id = 0
  for (var i = 1; i < data.length; i++) {
    var str = data[i][0].toString()
    //��ȡ���id��
    id = str.substring(3,str.length)
    if (data[i][1]==login_user){
      r.data.push({
        id: data[i][0],
        email: data[i][1],
        qjr: data[i][2],
        part:  data[i][3],
        sqdate: dateFormat(data[i][4]),        
        begindate: dateFormat(data[i][5]),
        enddate: dateFormat(data[i][6]),
        hours:data[i][7],
        reason:data[i][8],
        type:data[i][9],        
        other:data[i][10],
        status1:data[i][11],
        status2:data[i][12],
        status3:data[i][13],
        fileUrl:data[i][14]
      }); 
    } 
  }
  r.id = id?parseInt(id)+1:20180001
  return r
}

//�ύ�������
function addInfo(data){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getSheets()[1].appendRow([data.id,data.email,data.qjr,data.part,data.sqDate,data.beginDate,data.endDate,data.hours,data.reason,data.type,data.other?data.other:'']); 
  var s = getUserById(data.qjr)
  sendEmail(data.id,data.name,data.firstSpr,data.secondSpr,data.thirdSpr,0, data.beginDate,data.endDate,data.hours,data.reason,data.type,s.email,s.name)
}

//�޸Ĺ�ʱ
function updateGongShi(qjr,hours){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();     
  //�޸�Ա�������ʣ�๤ʱ
  var r = sheet.getSheets()[0].getDataRange().getValues();
  for (var i=1; i<r.length;i++){
     if (r[i][0]==qjr){
       var allHours = parseInt(r[i][8]);
       var syHours = parseInt(r[i][9]);
       var _hours = parseInt(hours);
       if (isNaN(allHours)){allHours=0}
       if (isNaN(syHours)){syHours=0}
       if (isNaN(_hours)){_hours=0}
       sheet.getSheets()[0].getRange("J"+(i+1)).setValue(syHours-hours);       
       break;
     }
  } 
}

//�༭�������
function updateInfo(item){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  var data = sheet.getDataRange().getValues();  
  for (var i=1; i<data.length;i++){
     if (data[i][0]==item.id){
       sheet.getRange("C"+(i+1)).setValue(item.sqDate);
       sheet.getRange("D"+(i+1)).setValue(item.beginDate);
       sheet.getRange("E"+(i+1)).setValue(item.endDate);
       sheet.getRange("F"+(i+1)).setValue(item.hours);
       sheet.getRange("G"+(i+1)).setValue(item.reason);
       sheet.getRange("H"+(i+1)).setValue(item.type);
       sheet.getRange("I"+(i+1)).setValue(item.other);
       sendEmail(item.id,item.name,item.firstSpr,item.secondSpr,item.thirdSpr,0, item.beginDate,item.endDate,item.hours,item.reason,item.type)
       break;
     }
  } 
}

//ɾ����ټ�¼
function deleteInfo(d){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getSheets()[1].getDataRange().getValues();  
  for (var i=1; i<data.length;i++){   
    if (data[i][0]==d.id){
      sheet.getSheets()[1].deleteRow(i+1);
      break;
    }
  }
  
}

//�޸�����״̬
function setAppStatus(id, count, pass, qjr, hours, type){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getSheets()[1].getDataRange().getValues();  
  for (var i=1; i<data.length;i++){   
    if (data[i][0]==id){
      Logger.log("status:"+data[i][0])
       Logger.log("iii:"+i)
      if (count==1){
        sheet.getSheets()[1].getRange("L"+(i+1)).setValue(pass); 
      } else if (count==2){
        sheet.getSheets()[1].getRange("M"+(i+1)).setValue(pass); 
      } else if (count == 3){
        sheet.getSheets()[1].getRange("N"+(i+1)).setValue(pass); 
        //�޸Ĺ�ʱ
        if (pass=='ͨ��' && type=='���'){
          updateGongShi(qjr, hours)
        }
      }      
      break;
    }
  }
}

//�޸��ļ�url
function setFileUrl(id, url){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data = sheet.getSheets()[1].getDataRange().getValues();  
  for (var i=1; i<data.length;i++){   
    if (data[i][0]==id){           
      sheet.getSheets()[1].getRange("O"+(i+1)).setValue(url);        
      break;
    }
  }
}

function uploadFiles(form) {
  
  try {   
    var dropbox = "images"; 
    //----------------------------------------------
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    var blob = form.myFile;   
    Logger.log(form)
    var file = folder.createFile(blob);    
    file.setDescription("�ϴ��ߣ� " + form.myName);
    Logger.log("�ɹ�")
    return file.getUrl()
    
  } catch (error) {
    
    return error.toString();
  }
  
}