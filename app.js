var fs = require('fs');
var http = require('http');

var paths = {
    '/publish': function(req, res){
        req.addListener("data", function(chunk) {
            req.content += chunk;
        });
        
        req.addListener("end", function() {
            //parse req.content and do stuff with it
        });
    }
}

var server = http.createServer(function(req, res) {
//    req.setEncoding("utf8");
//    req.content = '';
    
    var body = 'not yet implemented';
        console.log(req);


    if (req.method === 'POST') {
        console.log(req);
        body = "Oh, so you want to do a post, eh?"
        // pipe the request data to the console
//        req.pipe(process.stdout);
        var w = fs.createWriteStream('output.dat');
        req.pipe(w);
        
//        req.on('data', function (chunk) {
            
        //}
    }    


    res.writeHead(200, {
        'Content-Length': body.length,
        'Content-Type': 'text/plain' });
    res.write(body);
    res.end();
 
}).listen(8002);


