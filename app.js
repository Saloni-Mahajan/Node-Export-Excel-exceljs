const fs = require('fs');
const express=require('express');
const ejs=require('ejs');
const app=express();
const {exportExcel,homePage}=require('./controller/index');
app.set('view engine','ejs');


app.get('/',homePage)
app.post('/export',exportExcel)

app.listen(5000,()=>{
    console.log('port listening to 5000 ')});