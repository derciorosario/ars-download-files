const express = require('express');
const {join}=require('path')
const app = express();
const fs = require('fs');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const XlsxPopulate = require('xlsx-populate');


app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.set('port',process.env.PORT || 80)

const cors = require('cors');

const corsOptions = {
    origin: '*',
    methods: ['GET', 'POST'],
};
  

app.use(cors(corsOptions));

app.get('/', (req, res) => {
    res.send('Server is running...')
})

app.post('/',async (req, res) => {


    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet(req.body);

    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const filePath = `files/`
    const fileName = `ars-tabela-de-lancamento-${new Date().getFullYear() + "-" + new Date().getDay() + "-" + new Date().getMonth() +"-"+new Date().getHours()+"_"+new Date().getMinutes()+"_"+new Date().getSeconds()}.xlsx`
    xlsx.writeFile(workbook, filePath + fileName);

    const workbook2 = await XlsxPopulate.fromFileAsync(filePath + fileName);
    const sheet2 = workbook2.sheet('Sheet1');
   
    for (let columnIndex = 1; columnIndex <= sheet2.usedRange()._numColumns; columnIndex++) {
        const column = sheet2.column(columnIndex);
        column.width(15); 
    }

    sheet2.row(1).style({borderStyle:'thick'});
    sheet2.row(2).style({ bold: true });

    for (let i = 0; i < sheet2.usedRange()._numRows; i++) {
        if(i + 1 != 1)   sheet2.row(i + 1).style({wrapText: true ,horizontalAlignment:'right',borderStyle:'dashed'});
    }

    await workbook2.toFileAsync(join(__dirname,filePath + fileName));

    setInterval(()=>fs.unlink(join(__dirname,filePath + fileName), (error) => {}),40000)

    res.json(fileName)
     
});


  
app.use(express.static(join(__dirname,'/files')))

app.listen(app.get('port'));


