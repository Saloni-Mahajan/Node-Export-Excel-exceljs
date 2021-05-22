const MongoClient = require('mongodb').MongoClient;
const url = "mongodb://localhost:27017/excel";
const excel = require('exceljs');

const homePage=(req,res)=>{
    res.render('index');
}

const exportExcel=(req,res)=>{

    // Create a connection to the MongoDB database
    MongoClient.connect(url, { useNewUrlParser: true }, function(err, db) {
      if (err) throw err;
      
      let dbo = db.db("excel");
      
      dbo.collection("customers").find({}).toArray(function(err, result) {
        if (err) throw err;
        console.log(result);	
        let workbook = new excel.Workbook(); //creating workbook
        let worksheet = workbook.addWorksheet('Customers'); //creating worksheet

        //  WorkSheet Header
        worksheet.columns = [
            { header: 'Id', key: '_id', width: 10 },
            { header: 'Name', key: 'name', width: 30 },
            { header: 'Address', key: 'address', width: 30},
            { header: 'Age', key: 'age', width: 10, outlineLevel: 1}
        ];
        
        // Add Array Rows
        worksheet.addRows(result);
        
        // Write to File
         workbook.xlsx.writeFile("customer.xlsx")
            .then(function() {
                console.log("file saved!");
              
            });
            db.close();
        
        
      });
    });
    res.send('Excel Export Successfully')
}

module.exports={exportExcel,homePage}