<h1 align="center">
  <br>
  information extraction and CV generation with best-resume-ever 
  <br>
</h1>

<br>
<br>

## Prerequisite

1. It is required to have Node.js with version 12.16.3 or higher. To see what version of Node.js is installed on your machine type the following command in the terminal:

```
node -v
```

2. If you do not have installed Node.js in your machine then (ubuntu) :
```
- curl -sL https://deb.nodesource.com/setup_12.x | sudo -E bash -
- sudo apt-get install -y nodejs
```


3. It is required to have Python 2.7 or higher and :
```
- sudo apt install python-pip
- pip install ezodf
- pip install lxml
- pip install unidecode
- sudo apt-get install python-yaml
```




## How to use

1. Clone this repository.

```
git clone https://github.com/MoutonLina/CV.git
```

2. Go to the cloned directory.

3. Run `npm install`.

(ubuntu) : complete the puppetteer installation :

```
sudo apt install -y gconf-service libasound2 libatk1.0-0 libc6 libcairo2 libcups2 libdbus-1-3 libexpat1 libfontconfig1 libgcc1 libgconf-2-4 libgdk-pixbuf2.0-0 libglib2.0-0 libgtk-3-0 libnspr4 libpango-1.0-0 libpangocairo-1.0-0 libstdc++6 libx11-6 libx11-xcb1 libxcb1 libxcomposite1 libxcursor1 libxdamage1 libxext6 libxfixes3 libxi6 libxrandr2 libxrender1 libxss1 libxtst6 ca-certificates fonts-liberation libappindicator1 libnss3 lsb-release xdg-utils wget
```

4. Customize CV template in src/creativeLINA.vue.

5. put the CV forms (.ods) to the 'extract/Formulaires' folder. names of CVs forms have to be like "NAME_CV.ods"

6. put the contributors forms (.ods) to the 'extract/Fichescontrib' folder. names of a contributors forms have to be like "NAME_CONTRIBUTIONS.ods"

7. go to 'extract' folder and generate CV with 'python extractCsv2Yml2Pdf.py' and all Cvs ("NAME.pdf" or "Contrib_NAME.pdf")  will be generated in the /pdf folder

8. to generate the skills matrix : open skillsmatrix.ods and click on the "récupérer les infos" button on the Matrice Tab. The list of the CV will be generated and the skills matrix will be available on the "Talbeaux" tab. 



## Creating and Updating Templates

Please read the <a href="DEVELOPER.md">developer docs</a> on how to create or update templates.

<br>


## Credits

This project uses several open source packages:

- <a href="https://github.com/vuejs/vue" target="_blank">Vue</a>
- <a href="https://github.com/GoogleChrome/puppeteer" target="_blank">Puppeteer</a>
- <a href="https://github.com/less/less.js" target="_blank">LESS</a>
- <a href="https://github.com/salomonelli/best-resume-ever.git">best resume ever</a>

<br>


