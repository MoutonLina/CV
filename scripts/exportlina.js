const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const http = require('http');
const config = require('../config');
const TEMPLATE = 'creativeLina';

const {
    interval
} = require('rxjs');
const {
    filter,
    first,
    mergeMap
} = require('rxjs/operators');

const fetchResponse = () => {
    return new Promise((res, rej) => {
        try {
            const req = http.request(`http://localhost:${config.dev.port}/#/`, response => res(response.statusCode));
            req.on('error', (err) => rej(err));
            req.end();
        } catch (err) {
            rej(err);
        }
    });
};

const waitForServerReachable = () => {
    return interval(1000).pipe(
        mergeMap(async () => {
            try {
                const statusCode = await fetchResponse();
                if (statusCode === 200) return true;
            } catch (err) {}
            return false;
        }),
        filter(ok => !!ok)
    );
};
/*
const timedOut = timeout => {
    return new Promise(res => {
        setTimeout(res, timeout);
    });
};
*/
const convert = async () => {
    await waitForServerReachable().pipe(
        first()
    ).toPromise();

    console.log('Connected to server ...');
    console.log('Exporting ...');
    try {
        const fullDirectoryPath = path.join(__dirname, '../pdf/');
        const files = getFiles();
        
        console.log("les fichiers totaux : "+files);
        console.log(files.length);
        
        pagination=5;
        nbefiles=files.length //5
        nbetableaux=Math.floor(nbefiles/pagination)+1 //3
        console.log(nbetableaux)

        var SuperTableau = new Array(nbetableaux);
        i=0;
        k=0;




        /*
        for(var i=0; i < nbetableaux-1;i++)
            {    
                SuperTableau[i] = new Array(pagination-1);
                console.log(i);
                console.log("début de pagination");
                console.log(i*(pagination+1));
                j=0;
                for (var k=(i*(pagination+1));k<=((i+1)*pagination);k++){
                 //   console.log("i "+i);
                 // console.log("i "+i);
                  //   console.log("j "+j);                    
                     console.log("k "+k);
                   //  console.log("files "+files[k]);
                     SuperTableau[i][j]=files[k];
                     console.log("integration"+SuperTableau[i][j]);
                    j++;
                    
                }
            }
        */    
       
        k=0;
        while (k < nbefiles)
        {    
            for(var i=0; i < nbetableaux-1;i++){
                SuperTableau[i] = new Array(pagination-1);
                    for (var j=0;j<=pagination;j++){
                            if(typeof files[k] !== 'undefined'){
                            SuperTableau[i][j]=files[k];
                            console.log(SuperTableau[i][j]);
                            }
                            k++;                    
                        }
                }
        }


    for (var i = 0; i < nbetableaux-1; i++)
        {
            SuperTableau[i].forEach(async (file) => {
            const browser = await puppeteer.launch({
                args: ['--no-sandbox']
            });
            console.log('Generating resume', file, 'in', fullDirectoryPath);

            const page = await browser.newPage();
            await page.goto(`http://localhost:${config.dev.port}/#/resume/${file}?template=${TEMPLATE}`, {
                //waitUntil: 'networkidle2',
                waitUntil: 'load',
                timeout: 0
            });
            
            if (
                !fs.existsSync(fullDirectoryPath)
            ) {
                fs.mkdirSync(fullDirectoryPath);
            }
            await page.pdf({
                path: `${fullDirectoryPath}${file}.pdf`,
                format: 'A4',
                displayHeaderFooter: true,
                headerTemplate:"&nbsp;",
                footerTemplate:"<div style='height:2px; width:100%; text-align: right;font-size: 5pt;  bottom: 0; font-family:Roboto'><b>© LINAGORA 2020</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color='#c00d2d'>***&nbsp;:&nbsp;Expert&nbsp;&nbsp;&nbsp;&nbsp;**&nbsp;:&nbsp;Confirmé&nbsp;&nbsp;&nbsp;&nbsp;*&nbsp;:&nbsp;Compétent&nbsp;&nbsp;&nbsp;&nbsp; </font>&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;page&nbsp;&nbsp;<span class='pageNumber'></span>/<span class='totalPages'></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div>",
                margin:{
                    top: "12px",
                    bottom: "28px"
                }
            });
            await browser.close();
        });

    }





    } catch (err) {
        throw new Error(err);
    }
    console.log('Finished exports.');
};

const getFiles = () => {
    const srcpath = path.join(__dirname, '../resume');
    return fs.readdirSync(srcpath)
    .filter(file => file.endsWith('.yml'))
    .map(file => file.split('.')[0]);
};

convert();