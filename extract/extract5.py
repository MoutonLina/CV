# coding: utf-8
from ezodf import opendoc, Sheet
from unidecode import unidecode
import os
import os.path
import glob


#nettoyage du dossier cible des yml
for fold in os.listdir('../resume'):
    if fold.endswith('.yml'):
         os.remove('../resume/'+fold)

#traitement des formulaires de CV
odsfiles = []
for file in glob.glob("Formulaires/*.ods"):
    odsfiles.append(file)


for f in odsfiles:
    doc = opendoc(f)
    sheets=doc.sheets
    targetsheet='Survey'
    for sheet in sheets:
        if (sheet.name =='Formulaire'):
            targetsheet='Formulaire'
    sheet = doc.sheets[targetsheet]
    # for sheet in doc.sheets:
    import yaml


    islls=23
    sllschain=''
    while islls < 45 : 
        sllschain += "{name: "+ unidecode(unicode(sheet['C'+ str(islls)].value)) +" , level: " + unidecode(unicode(sheet['G' + str(islls)].value))+ "},"
        islls += 1

    isllsbis=49
    sllsbischain=''
    while isllsbis < 65 : 
        sllsbischain += "{name: "+ unidecode(unicode(sheet['C'+ str(isllsbis)].value)) +" , level: " + unidecode(unicode(sheet['G' + str(isllsbis)].value))+ "},"
        isllsbis += 1

    iproject=734
    try: 
        sheet['E734']
    except:
        iproject=286    
    #iproject=734
    #iproject=286
    projectchain=''
    while iproject >= 83 : 
        projectchain += "{enterprise: "+ unidecode(unicode(sheet['E'+ str(iproject + 1)].value)).replace(":"," ").replace("'"," ") +" , period: " + unidecode(unicode(sheet['E' + str(iproject)].value)).replace(":"," ").replace("["," ").replace("]"," ").replace("'"," ")+ ", project: '" + unidecode(unicode(sheet['E' + str(iproject + 2)].value)).replace(":"," ").replace("'"," ")+ "', function: " + unidecode(unicode(sheet['E' + str(iproject + 3)].value)).replace(":"," ").replace("'"," ")+ ", missions: '|" + unidecode(unicode(sheet['E' + str(iproject + 4)].value)).replace(":"," ").replace("'"," ")+ "', technos: '|" + unidecode(unicode(sheet['E' + str(iproject + 5)].value)).replace(":"," ").replace("'"," ")+ "'},"
        print unidecode(unicode(sheet['E' + str(iproject + 2)].value)).replace(":"," ").replace("["," ").replace("]"," ").replace("'"," ")
        iproject -= 7
    



    iskills=71
    skillschain=''
    while iskills < 80 : 
        skillschain += "{name: "+ unidecode(unicode(sheet['C'+ str(iskills)].value)).replace(":"," ") +" , level: " + unidecode(unicode(sheet['G' + str(iskills)].value))+ "},"
        iskills += 1
    


    anexptab = str(sheet['C18'].value).split(".")
    anexp=anexptab[0]
    #aboutchain= "| \n\t" + unidecode(sheet['H4'].value)
    #print unidecode(sheet['H4'].value)
    aboutchain= unidecode(sheet['H4'].value)
    #on va rechercher les elements dans la fiche contrib si elle existe 
    tab_contrib=["D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X"]
    trigram = unidecode(sheet['C11'].value)
    contriblist=''
    eventlist=''
    publicationlist=''
    for filecontrib in glob.glob("Fichescontrib/*.ods"):
        #print os.path.basename(filecontrib).split("_",0)
        fichier=os.path.basename(filecontrib)
        #print trigram
        #print fichier.split("_") 
        splittab=fichier.split("_")
        if(splittab[0] == trigram):
            doccontrib = opendoc(filecontrib)
            sheetcontrib = doccontrib.sheets['Sheet1']
            #print(unidecode(unicode(sheetcontrib['B3'].value)))

            for y in tab_contrib:
                #print y
                #print unidecode(unicode(sheetcontrib[y + "4"].value))
                contriblist += "{community: "+ unidecode(unicode(sheetcontrib[y + "4"].value)).replace(":"," ") +" , subject: " + unidecode(unicode(sheetcontrib[y + "5"].value)).replace(":"," ") +" , description: " + unidecode(unicode(sheetcontrib[y + "6"].value)).replace(":"," ") + "},"
                eventlist += "{event: "+ unidecode(unicode(sheetcontrib[y + "10"].value)).replace(":"," ") +" , type: " + unidecode(unicode(sheetcontrib[y + "11"].value)).replace(":"," ") +" , description: " + unidecode(unicode(sheetcontrib[y + "12"].value)).replace(":"," ") + "},"
                publicationlist += "{pub: "+ unidecode(unicode(sheetcontrib[y + "15"].value)).replace(":"," ") +" , type: " + unidecode(unicode(sheetcontrib[y + "16"].value)).replace(":"," ") +" , description: " + unidecode(unicode(sheetcontrib[y + "17"].value)).replace(":"," ") + "},"
    #on va rechercher les elements dans la fiche contrib si elle existe 

    dict_file ={
    'name' : 
        {'first':unidecode(sheet['C15'].value) , 
        'trigramme':unidecode(sheet['C11'].value), 
        'last':unidecode(sheet['C14'].value)}
    ,
    'anexp': anexp,

    'about' : aboutchain,

    'position': unidecode(sheet['F18'].value),

    'exellence' :
            [
            {'content': unidecode(unicode(sheet['M15'].value))},
            {'content': unidecode(unicode(sheet['M17'].value))},
            {'content': unidecode(unicode(sheet['M19'].value))},
            {'content': unidecode(unicode(sheet['M21'].value))},
            {'content': unidecode(unicode(sheet['M23'].value))}]
    ,

    'experience' :
    [
    {'company': unidecode(unicode(sheet['J30'].value)) , 'timeperiod':unidecode(unicode(sheet['J31'].value)) , 'description': unidecode(unicode(sheet['J32'].value)) },
    {'company': unidecode(unicode(sheet['J35'].value)) , 'timeperiod':unidecode(unicode(sheet['J36'].value)) , 'description': unidecode(unicode(sheet['J37'].value)) },
    {'company': unidecode(unicode(sheet['J40'].value)) , 'timeperiod':unidecode(unicode(sheet['J41'].value)) , 'description': unidecode(unicode(sheet['J42'].value)) },
    {'company': unidecode(unicode(sheet['J45'].value)) , 'timeperiod':unidecode(unicode(sheet['J46'].value)) , 'description': unidecode(unicode(sheet['J47'].value)) },
    {'company': unidecode(unicode(sheet['J50'].value)) , 'timeperiod':unidecode(unicode(sheet['J51'].value)) , 'description': unidecode(unicode(sheet['J52'].value)) },
    {'company': unidecode(unicode(sheet['J55'].value)) , 'timeperiod':unidecode(unicode(sheet['J56'].value)) , 'description': unidecode(unicode(sheet['J57'].value)) },
    {'company': unidecode(unicode(sheet['J60'].value)) , 'timeperiod':unidecode(unicode(sheet['J61'].value)) , 'description': unidecode(unicode(sheet['J62'].value)) },
    ]
    ,


    'education' :
            [
            {'contenu': unidecode(unicode(sheet['I18'].value))},
            {'contenu': unidecode(unicode(sheet['I20'].value))},
            {'contenu': unidecode(unicode(sheet['I22'].value))},
            ]
    ,

    'caracteristics' :
    [
    {'name': unidecode(unicode(sheet['N30'].value)) , 'level':unidecode(unicode(sheet['O30'].value))},
    {'name': unidecode(unicode(sheet['N31'].value)) , 'level':unidecode(unicode(sheet['O31'].value))},
    {'name': unidecode(unicode(sheet['N32'].value)) , 'level':unidecode(unicode(sheet['O32'].value))},
    {'name': unidecode(unicode(sheet['N33'].value)) , 'level':unidecode(unicode(sheet['O33'].value))},
    {'name': unidecode(unicode(sheet['N34'].value)) , 'level':unidecode(unicode(sheet['O34'].value))},
    {'name': unidecode(unicode(sheet['N35'].value)) , 'level':unidecode(unicode(sheet['O35'].value))},
    {'name': unidecode(unicode(sheet['N36'].value)) , 'level':unidecode(unicode(sheet['O36'].value))},
    {'name': unidecode(unicode(sheet['N37'].value)) , 'level':unidecode(unicode(sheet['O37'].value))},
    ]
    ,

    'languages' :
    [
    {'name': unidecode(unicode(sheet['N41'].value)) , 'level':unidecode(unicode(sheet['O41'].value))},
    {'name': unidecode(unicode(sheet['N42'].value)) , 'level':unidecode(unicode(sheet['O42'].value))},
    {'name': unidecode(unicode(sheet['N43'].value)) , 'level':unidecode(unicode(sheet['O43'].value))},
    {'name': unidecode(unicode(sheet['N44'].value)) , 'level':unidecode(unicode(sheet['O44'].value))},
    {'name': unidecode(unicode(sheet['N45'].value)) , 'level':unidecode(unicode(sheet['O45'].value))},
    {'name': unidecode(unicode(sheet['N46'].value)) , 'level':unidecode(unicode(sheet['O46'].value))},
    {'name': unidecode(unicode(sheet['N47'].value)) , 'level':unidecode(unicode(sheet['O47'].value))},
    {'name': unidecode(unicode(sheet['N48'].value)) , 'level':unidecode(unicode(sheet['O48'].value))},
    {'name': unidecode(unicode(sheet['N49'].value)) , 'level':unidecode(unicode(sheet['O49'].value))},
    {'name': unidecode(unicode(sheet['N50'].value)) , 'level':unidecode(unicode(sheet['O50'].value))},
    {'name': unidecode(unicode(sheet['N51'].value)) , 'level':unidecode(unicode(sheet['O51'].value))},
    {'name': unidecode(unicode(sheet['N52'].value)) , 'level':unidecode(unicode(sheet['O52'].value))},
    ]
    ,

    'slls' :
    [
    sllschain
    ]
    ,

    'sllsbis' :
    [
    sllsbischain
    ]
    ,

    'skills' :
    [
    skillschain
    ]
    ,

    'listcontribs' :
    [
    contriblist
    ]
    ,

    'listevents' :
    [
    eventlist
    ]
    ,

    'listpubs' :
    [
    publicationlist
    ]
    ,

    'projects' :
    [
    projectchain
    ]
    ,

    'contributions' : unidecode(unicode(sheet['L77'].value)).replace(":"," "),

    'conferences' : unidecode(unicode(sheet['I66'].value)).replace(":"," "),

    'fierte' : unidecode(unicode(sheet['L91'].value)).replace(":"," "),

    'loisirs' : unidecode(unicode(sheet['N55'].value)).replace(":"," "),






    'lang' : 'fr',
    }







    with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '_tmp.yml', 'w') as file:
        file.write("/* #*/ export const PERSON = ` \n")
        yaml.unicode_supplementary=False
        documents = yaml.dump(dict_file, file)
        file.write("\n`")
        file.close()

    with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '_tmp.yml','r') as fichier:
        #with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '.yml','w') as fichierFinale:
        with open('../resume/' + unidecode(sheet['C11'].value) + '.yml','w') as fichierFinale:
            lignes = fichier.readlines()                # On parcours les lignes du fichier source
            for ligne in lignes:
                lignetmp=ligne.replace("['","[")
                lignetmp=lignetmp.replace("project: ''","project: '")
                lignetmp=lignetmp.replace("'', function","', function")
                lignetmp=lignetmp.replace("missions: ''","missions: '")
                lignetmp=lignetmp.replace("technos: ''","technos: '")
                lignetmp=lignetmp.replace("''|","'|")
                lignetmp=lignetmp.replace("''},","'},")
                lignetmp=lignetmp.replace("'',","',")
                lignetmp=lignetmp.replace("\"]","]")
                lignetmp=lignetmp.replace("[\"","[")
                lignetmp=lignetmp.replace("\"|","|")
                lignetmp=lignetmp.replace("\"","")
                ligneFinale=lignetmp.replace("']","]")
                

                fichierFinale.write(ligneFinale)               # On écrit la nouvelle ligne dans le nouveau fichier
        fichierFinale.close()                     # Fermeture du fichier source
    fichier.close()
    os.remove('./yamlfiles/' + unidecode(sheet['C11'].value) + '_tmp.yml')
os.system("npm run dev")