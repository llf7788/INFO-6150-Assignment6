const http = require('http');
const fs = require('fs');

// create a node http server, we can process every request and response in it's callback function
http.createServer((req, res) => {
  // get request url
  const url = req.url;
  // if match root url, display the page of spreadsheet.html
  if (url ==='/'){
    // write http header, for html file Content-Type is 'text/html'
    res.writeHead(200, {'Content-Type': 'text/html'}); 
    // read spreadsheet.html from dist and write it to response
    const buffer = fs.readFileSync('spreadsheet.html');
    res.write(buffer);
    res.end(); 
  // if url is '/load_data' , return a json response containing csv data
  } else if (url ==='/load_data'){
    // the csv file to read
    const datafile = 'sheet_1583652909952.csv';
    // read file from disk
    const csvContent = fs.readFileSync(datafile).toString('utf-8');
    let csvData = [];
    // save csv data in array 'csvData'
    for (let line of csvContent.split('\n')) {
      const row = [];
      for (let v of line.split('\t')) {
        row.push(v);
      }
      csvData.push(row);
    }
    // response as json
    res.writeHead(200, {'Content-Type': 'application/json; charset=utf-8'});
    res.write(JSON.stringify(csvData));
    res.end(); 
  // for js files, as treat it as static assets, response the file as it is.
  } else if (url.endsWith('.js')){
    res.writeHead(200, {'Content-Type': 'text/javascript; charset=utf-8'}); // http header
    const filePath = url.slice(1);
    const buffer = fs.readFileSync(filePath);
    res.write(buffer);
    res.end(); 
  // for css files, as treat it as static assets, response the file as it is.
  } else if (url.endsWith('.css')){
    res.writeHead(200, {'Content-Type': 'text/css; charset=utf-8'}); // http header
    const filePath = url.slice(1);
    const buffer = fs.readFileSync(filePath);
    res.write(buffer);
    res.end(); 
  // if request url does not match any pattern above, display the '404 not found' page
  } else {
    res.write('<h1>404<h1>');
    res.end(); 
  }
 }).listen(3000, function(){
  console.log("server start at port 3000"); //the server object listens on port 3000
 });