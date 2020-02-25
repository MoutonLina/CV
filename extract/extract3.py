# coding: utf-8
from ezodf import opendoc, Sheet
from unidecode import unidecode
import os
import glob
odsfiles = []
for file in glob.glob("Formulaires/*.ods"):
    odsfiles.append(file)


for f in odsfiles:
    doc = opendoc(f)
    sheet = doc.sheets['Formulaire']
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


    try:
        godata=sheet['E734'].value
    except IndexError:
        iproject=286
    else:
        iproject=734
    #print iproject
    #iproject=286
    projectchain=''
    while iproject > 83 : 
        projectchain += "{enterprise: "+ unidecode(unicode(sheet['E'+ str(iproject + 1)].value)).replace(":"," ") +" , period: " + unidecode(unicode(sheet['E' + str(iproject)].value)).replace(":"," ").replace("["," ").replace("]"," ")+ ", project: '|" + unidecode(unicode(sheet['G' + str(iproject + 2)].value)).replace(":"," ")+ "', function: " + unidecode(unicode(sheet['E' + str(iproject + 3)].value)).replace(":"," ")+ ", missions: '|" + unidecode(unicode(sheet['E' + str(iproject + 4)].value)).replace(":"," ")+ "', technos: '|" + unidecode(unicode(sheet['E' + str(iproject + 5)].value)).replace(":"," ")+ "'},"
        iproject -= 7
    



    iskills=71
    skillschain=''
    while iskills < 80 : 
        skillschain += "{name: "+ unidecode(unicode(sheet['C'+ str(iskills)].value)).replace(":"," ") +" , level: " + unidecode(unicode(sheet['G' + str(iskills)].value))+ "},"
        iskills += 1
    


    anexptab = str(sheet['C18'].value).split(".")
    anexp=anexptab[0]


    aboutchain= unidecode(sheet['H4'].value)

    dict_file ={
    'name' : 
        {'first':unidecode(sheet['C15'].value) , 
        'trigramme':unidecode(sheet['C11'].value), 
        'last':unidecode(sheet['C14'].value)}
    ,
    'anexp': anexp,

    'about': aboutchain,

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
        documents = yaml.dump(dict_file, file)
        file.write("\n`")
        file.close()

    with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '_tmp.yml','r') as fichier:
        with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '.yml','w') as fichierFinale:
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
                ligneFinale=lignetmp.replace("']","]")
                fichierFinale.write(ligneFinale)               # On écrit la nouvelle ligne dans le nouveau fichier
        fichierFinale.close()                     # Fermeture du fichier source
    fichier.close()
    os.remove('./yamlfiles/' + unidecode(sheet['C11'].value) + '_tmp.yml')