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
        files.forEach(async (file) => {
            const browser = await puppeteer.launch({
                args: ['--no-sandbox']
            });
            console.log('Generating resume', file, 'in', fullDirectoryPath);

            const page = await browser.newPage();
            await page.goto(`http://localhost:${config.dev.port}/#/resume/${file}?template=${TEMPLATE}`, {
                waitUntil: 'networkidle2'
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
                footerTemplate:"<div style='height:2px; width:100%; text-align: center;font-size: 5pt;  bottom: 0; font-family:Roboto'><b>© Linagora 2020</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ***:Expert&nbsp;&nbsp;**:Confirmé&nbsp;&nbsp;*:Connaissances</div>",
                margin:{
                    top: "12px",
                    bottom: "28px"
                }
            });
            await browser.close();
        });
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