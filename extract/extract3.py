# coding: utf-8
from ezodf import opendoc, Sheet
from unidecode import unidecode

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
        sllschain += "{'name': "+ unidecode(unicode(sheet['C'+ str(islls)].value)) +" , 'level':" + unidecode(unicode(sheet['G' + str(islls)].value))+ "},"
        islls += 1


    isllsbis=49
    sllsbischain=''
    while isllsbis < 65 : 
        sllsbischain += "{'name': "+ unidecode(unicode(sheet['C'+ str(isllsbis)].value)) +" , 'level':" + unidecode(unicode(sheet['G' + str(isllsbis)].value))+ "},"
        isllsbis += 1


    #iproject=734
    iproject=286
    projectchain=''
    while iproject > 83 : 
        projectchain += "{'enterprise': "+ unidecode(unicode(sheet['E'+ str(iproject - 1)].value)) +" , 'period':" + unidecode(unicode(sheet['E' + str(iproject)].value))+ ", 'project': |" + unidecode(unicode(sheet['G' + str(iproject - 2)].value))+ ", 'function':" + unidecode(unicode(sheet['E' + str(iproject - 3)].value))+ ", 'missions': |" + unidecode(unicode(sheet['E' + str(iproject - 4)].value))+ "}, 'technos': |" + unidecode(unicode(sheet['E' + str(iproject - 5)].value))+ "},"
        print(iproject)
        iproject -= 7


    dict_file ={
    'name' : 
        {'first':unidecode(sheet['C15'].value) , 
        'trigramme':unidecode(sheet['C11'].value), 
        'last':unidecode(sheet['C14'].value)}
    ,
    'anexp' : str(sheet['C18'].value),

    'about': '|' + unidecode(sheet['H4'].value),

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

    'skills' :
    [
    {'name': unidecode(unicode(sheet['C71'].value)) , 'level':unidecode(unicode(sheet['G71'].value))},
    {'name': unidecode(unicode(sheet['C72'].value)) , 'level':unidecode(unicode(sheet['G72'].value))},
    {'name': unidecode(unicode(sheet['C73'].value)) , 'level':unidecode(unicode(sheet['G73'].value))},
    {'name': unidecode(unicode(sheet['C74'].value)) , 'level':unidecode(unicode(sheet['G74'].value))},
    {'name': unidecode(unicode(sheet['C75'].value)) , 'level':unidecode(unicode(sheet['G75'].value))},
    {'name': unidecode(unicode(sheet['C76'].value)) , 'level':unidecode(unicode(sheet['G76'].value))},
    {'name': unidecode(unicode(sheet['C77'].value)) , 'level':unidecode(unicode(sheet['G77'].value))},
    {'name': unidecode(unicode(sheet['C78'].value)) , 'level':unidecode(unicode(sheet['G78'].value))},
    {'name': unidecode(unicode(sheet['C79'].value)) , 'level':unidecode(unicode(sheet['G79'].value))},
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


    'projects' :
    [
    projectchain
    ]
    ,

    'contributions' : '|' + unidecode(unicode(sheet['L77'].value)),

    'conferences' : '|' + unidecode(unicode(sheet['I66'].value)),

    'lang' : 'fr',
    }


    with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '.yml', 'w') as file:
        file.write("/* #*/ export const PERSON = ` \n")
        documents = yaml.dump(dict_file, file)
        file.write("\n`")