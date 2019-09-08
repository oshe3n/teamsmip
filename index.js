const express = require('express');
const app = express();
const path = require('path');
const router = express.Router();
var bodyParser = require('body-parser');    
var urlencodedParser = bodyParser.urlencoded({ extended: false })  
var request = require('request');
var sp = require("@pnp/sp").sp;
var SPFetchClient = require("@pnp/nodejs").SPFetchClient;

app.use(bodyParser.urlencoded({ extended: false }))

sp.setup({
    sp: {   
        fetchClientFactory: () => {
            return new SPFetchClient("https://m365x628217.sharepoint.com/sites/TestTeamsMIP", "c4ab9843-65f4-41e9-8a49-c5e04881f0db", "LtjTLjFxgYFZnG6D1YFfGlMzMZzvwI/BuU4DODu1v+I=");
        },
    },
});

app.engine('html', require('ejs').renderFile);

// -----------------------------------------TABS-----------------------------------------
router.get('/',function(req,res){
  res.sendFile(path.join(__dirname+'/index.html'));
  //__dirname : It will resolve to your project folder.
  console.log("Rendering Configuration")
});

router.get('/a',function(req,res){        
    var mainarr=[]    
    sp.web.lists.getByTitle("LOB_INTERNAL").items.orderBy('Created',false).top(10).get().then((items) => {        
        items.reverse().forEach(element => {      
            var subarr=[];                 
            if(element.Status == 'Open') 
            {
                subarr.push(element.Title)
                subarr.push(element.CaseTitle)
                subarr.push(element.CaseType)
                subarr.push(element.AssignedTeam)
                subarr.push(element.Status)                  
                mainarr.push(subarr);                 
            }                          
        })
    }).then(function(){        
        res.render(path.join(__dirname+'/a.html'),{data: mainarr });    
    })  
    console.log("Rendering A/Open")
});

router.get('/b',function(req,res){
    var mainarr=[]    
    sp.web.lists.getByTitle("LOB_INTERNAL").items.orderBy('Created',false).top(10).get().then((items) => {        
        items.reverse().forEach(element => {      
            var subarr=[];                 
            if(element.Status == 'InProgress') 
            {
                subarr.push(element.Title)
                subarr.push(element.CaseTitle)
                subarr.push(element.CaseType)
                subarr.push(element.AssignedTeam)
                subarr.push(element.Status)                  
                mainarr.push(subarr);                 
            }                          
        })
    }).then(function(){        
        res.render(path.join(__dirname+'/b.html'),{data: mainarr });    
    })  
    console.log("Rendering B/Closed")
});
// -----------------------------------------TABS-----------------------------------------

// -----------------------------------CONNECTORS-----------------------------------------
router.get('/connector',function(req,res){
    res.sendFile(path.join(__dirname+'/connector.html'));    
    console.log("Setting up connector")
});

app.post('/connector_save',urlencodedParser,function(req,res){ 
    console.log(req.body.webhook)
    sp.web.lists.getByTitle("LOB_WEBHOOK").items.add({
       "Title": req.body.webhook
    })
          
    console.log("Saving up connector")
    res.send("Saved")
});

// -----------------------------------CONNECTORS-----------------------------------------

// -----------------------------------LOB CRUD-------------------------------------------
router.get('/lob',function(req,res){
           
    var mainarr=[]    
    sp.web.lists.getByTitle("LOB_INTERNAL").items.orderBy('Created',false).top(10).get().then((items) => {        
        items.reverse().forEach(element => {      
            var subarr=[];      
            subarr.push(element.Title)
            subarr.push(element.CaseTitle)
            subarr.push(element.CaseType)
            subarr.push(element.AssignedTeam)
            subarr.push(element.Status)              
            mainarr.push(subarr);                           
        })
    }).then(function(){        
        res.render(path.join(__dirname+'/lob.html'),{data: mainarr });    
    })  
    
    console.log("LOB")
});

// WEBHOOK FOR CONNECTOR
app.post('/process_post', urlencodedParser, function (req, res) {  
    
    var x = '{"text":"Request ID : ESC-10234-TGS23","sections":[{"activityTitle":"Escalation Management System","activityText":"New Request has been created | Status - OPEN","activityImage":"https://ms-vsts.gallerycdn.vsassets.io/extensions/ms-vsts/vss-services-teams/1.0.11/1533756051713/Microsoft.VisualStudio.Services.Icons.Branding","facts":[{"name":"Customer","value":"'+req.body.Customer+'"},{"name":"Case Title","value":"'+req.body.CaseTitle+'"},{"name":"Case Type","value":"'+req.body.CaseType+'"}]}]}';        

    var webhookUrl="";
    sp.web.lists.getByTitle("LOB_WEBHOOK").items.orderBy('Created',false).top(1).get().then((items) => {    
        webhookUrl = items[0].Title
    }).then(function(){
        request.post(
            webhookUrl,
            { json: JSON.parse(x) },
            function (error, response, body) {
                if (!error && response.statusCode == 200) {
                    console.log(body);
                }
            }
        );
    })
    
    sp.web.lists.getByTitle("LOB_INTERNAL").items.add({
        "Title": req.body.Customer,
        "CaseTitle" : req.body.CaseTitle,
        "CaseType" : req.body.CaseType,
        "AssignedTeam" : "NotAssigned",
        "Status" : "Open"
     }).then(function(){
        res.redirect('/lob');
     })

    
 })  
// -----------------------------------LOB CRUD-------------------------------------------

app.use('/', router);
app.listen(process.env.port || 3333);

console.log('Running at Port 3333');
