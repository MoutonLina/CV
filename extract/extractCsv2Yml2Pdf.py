# coding: utf-8
from ezodf import opendoc, Sheet
from unidecode import unidecode
import os
import os.path
import glob
import sys
import math
#tri_contributeurs=[]
#tri_contributeurs=["ADU","CHA","DCO","MBAI","NCH","ROU","VZU","FMA","IAB","DCA","PVI","JSB","SEL","RBO","MSE","CDE","AHE","DLA","CTT","GSH","SBA","YBE","ZSE"]
#tri_contributeurs=["ADU","CHA","DCO","MBAI","NCH"]
#tri_contributeurs=["IAB"]
#tri_contributeurs=["ROU"]
#tri_contributeurs=["MBAI"]
tri_contributeurs=[]

contrib_absents=[]
contrib_presents=[]
contrib_absents_finaux=[]

#fonction de nettoyage de caractères spéciaux
def nettoyage(champ):
      return champ.replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace(",","&#44;").replace(".","&#46;").replace(" #"," &#35;").replace("-","&#45;").replace("["," ").replace("]"," ").replace("*","&#42;").replace(">","&#62;").replace("-"," ").replace("\n","<br/>").replace("`o"," ")      
                   

#nettoyage du dossier cible des yml
for foldyml in os.listdir('../resume'):
    if foldyml.endswith('.yml'):
         os.remove('../resume/'+foldyml)

#nettoyage du dossier cible des pdf
for foldpdf in os.listdir('../pdf'):
    if foldpdf.endswith('.pdf'):
         os.remove('../pdf/'+foldpdf)

#traitement des formulaires de CV
odsfiles = []
fichier_a_traiter= []
for file in glob.glob("Formulaires/*.ods"):
    odsfiles.append(file)

#recuperation des fiches contributeurs
if(tri_contributeurs!=[]):
    for fc in odsfiles:  
        #print fc
        id_fichier=os.path.splitext(fc)[0]
        id_fichier.split("_")[0].split("/")[1]
        if id_fichier.split("_")[0].split("/")[1] in tri_contributeurs:
            fichier_a_traiter.append(fc)
else:
    fichier_a_traiter=odsfiles
    #print fichier_a_traiter

#pour eviter le timeout, on partitione les traitements
pagination=20
nbefiles=len(fichier_a_traiter)
nbetab= math.floor((nbefiles/pagination))+1

tab=[[0] * int(nbetab) for i in range(int(nbetab))]
for j in range(0, int(nbetab)):
    tab[j]=fichier_a_traiter[j*pagination:(j+1)*pagination]
    for f in tab[j]:
        doc = opendoc(f)
        sheets=doc.sheets
        targetsheet='Survey'
        for sheet in sheets:
            if (sheet.name =='Formulaire'):
                targetsheet='Formulaire'
        sheet = doc.sheets[targetsheet]
        # for sheet in doc.sheets:
        tab_contrib=["D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X"]
        #tab_contrib=["D","E"]
        #trigram = unidecode(sheet['C11'].value)

        id_fichier=os.path.splitext(f)[0]
        trigram=id_fichier.split("_")[0].split("/")[1]



        contriblist=''
        contrib_ct=0
        eventlist=''
        event_ct=0
        publicationlist=''
        pub_ct=0
        import yaml

        #note : les islls correspondent aux cellules qui contiennent les informations recupérées dans le calc pour generer le yml
        #ne pas modifier leurs valeurs sans modifier le calc correspondant
        islls=23
        sllschain=''
        while islls < 45 : 
            sllschain += "{name: "+ sheet['C'+ str(islls)].value.split('.')[1] +" , level: " + unidecode(unicode(sheet['G' + str(islls)].value))+ "},"
            islls += 1

        isllsbis=49
        sllsbischain=''
        while isllsbis < 65 : 
            sllsbischain += "{name: "+ unidecode(unicode(sheet['C'+ str(isllsbis)].value)) +" ,"
            sllsbischain += "level: " + unidecode(unicode(sheet['G' + str(isllsbis)].value))+ "},"
            isllsbis += 1

        iproject=734
        try: 
            sheet['E734']
        except:
            iproject=286    
        #note : les iproject correspondent aux cellules qui contiennent les informations recupérées dans le calc pour generer le yml
        #ne pas modifier leurs valeurs sans modifier le calc correspondant
        #iproject=734
        #iproject=286
        projectchain=''
        while iproject >= 83 : 
            projectchain += "{enterprise: "+ nettoyage(unidecode(unicode(sheet['E'+ str(iproject + 1)].value))) +"," 
            projectchain += "period: " + nettoyage(unidecode(unicode(sheet['E' + str(iproject)].value))) + ","
            projectchain += "project: " + nettoyage(unidecode(unicode(sheet['E' + str(iproject + 2)].value))) + ","
            projectchain += "function: " + nettoyage(unidecode(unicode(sheet['E' + str(iproject + 3)].value))) + ","
            projectchain += "missions: " + nettoyage(unidecode(unicode(sheet['E' + str(iproject + 4)].value))) + ","
            #projectchain += "missions: " + nettoyage(unidecode(unicode(sheet['E' + str(iproject + 4)].value)))+ ","
            projectchain += "technos: " + nettoyage(unidecode(unicode(sheet['E' + str(iproject + 5)].value))) + "},"
            
            #print unidecode(unicode(sheet['E' + str(iproject + 2)].value)).replace(":"," ").replace("["," ").replace("]"," ").replace("'"," ")
            iproject -= 7
        



        iskills=71
        skillschain=''
        while iskills < 80 : 
            skillschain += "{name: "+ unidecode(unicode(sheet['C'+ str(iskills)].value)).replace(":"," ") +" , level: " + unidecode(unicode(sheet['G' + str(iskills)].value))+ "},"
            iskills += 1
        


        #anexptab = str(sheet['C18'].value).split(".")
        try :
            #anexptab = str(sheet['C18'].value.split("."))
            anexptab=str(sheet['C18'].value).split(".")
            anexp=anexptab[0]
        except :
            print "pb : "
            print trigram
            print "fin pb"
            anexp=''

        
        
        aboutchain= sheet['H4'].value.replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br />").replace(",","&#44;").replace(".","&#46;")
        #aboutchain= unidecode(unicode(sheet['H4'].value)).replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br />").replace(",","&#44;").replace(".","&#46;")
        #aboutchain= unidecode(unicode(sheet['H4'].value))    
        #on va rechercher les elements dans la fiche contrib si elle existe 
        
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
                    try: 
                        #contriblist += "{community: "+ unidecode(unicode(sheetcontrib[y + "4"].value)).replace(":"," ") +" , subject: " + unidecode(unicode(sheetcontrib[y + "5"].value)).replace(":"," ") +" , description: " + unidecode(unicode(sheetcontrib[y + "6"].value)).replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br/>").replace(",","&#44;").replace(".","").replace(" #"," &#35;").replace("-","") + "},"
                        contriblist += "{community: "+ nettoyage(unidecode(unicode(sheetcontrib[y + "4"].value))) +" , subject: " + nettoyage(unidecode(unicode(sheetcontrib[y + "5"].value)))+" , description: " + nettoyage(unidecode(unicode(sheetcontrib[y + "6"].value))) + "},"
                        if(nettoyage(unidecode(unicode(sheetcontrib["D4"].value)))!='None'):
                            contrib_ct=1
                    except:
                        print "pb contrib "+trigram

                    try: 
                        eventlist += "{event: "+ nettoyage(unidecode(unicode(sheetcontrib[y + "10"].value))) +" , description: " + nettoyage(unidecode(unicode(sheetcontrib[y + "12"].value))) +" , type: " + nettoyage(unidecode(unicode(sheetcontrib[y + "11"].value))) + "},"
                        if(nettoyage(unidecode(unicode(sheetcontrib["D10"].value)))!='None'):
                            event_ct=1
                    except:
                        print "pb event "+trigram

                    try: 
                        publicationlist += "{pub: "+ nettoyage(unidecode(unicode(sheetcontrib[y + "15"].value)))+" , type: " + nettoyage(unidecode(unicode(sheetcontrib[y + "16"].value))) +" , description: " + nettoyage(unidecode(unicode(sheetcontrib[y + "17"].value))) + "},"
                        if(nettoyage(unidecode(unicode(sheetcontrib["D15"].value)))!='None'):
                            pub_ct=1
                    except:
                        print "pb publication "+trigram

        #on va rechercher les elements dans la fiche contrib si elle existe 
        try:
            firstname=unidecode(sheet['C15'].value)
        except:
            firstname=""

        try:
            lastname=unidecode(sheet['C14'].value)
        except:
            lastname=""

        try:
            positionvalue=unidecode(sheet['F18'].value).replace("\n"," ")
        except:
            positionvalue=""

        dict_file ={
        'name' : 
            {'first':firstname , 
            'trigramme':trigram, 
            'last':lastname}
        ,
        'anexp': anexp,

        'partner': unidecode(sheet['F12'].value) ,

        'about' : aboutchain,

        'position': positionvalue,

        'exellence' :
                [
                {'content': unidecode(unicode(sheet['M15'].value)).replace("\n"," ")},
                {'content': unidecode(unicode(sheet['M17'].value)).replace("\n"," ")},
                {'content': unidecode(unicode(sheet['M19'].value)).replace("\n"," ")},
                {'content': unidecode(unicode(sheet['M21'].value)).replace("\n"," ")},
                {'content': unidecode(unicode(sheet['M23'].value)).replace("\n"," ")}]
        ,
        'partner' : unidecode(sheet['F12'].value),

        'experience' :
        [
        {'company': unidecode(unicode(sheet['J30'].value)) , 'timeperiod':unidecode(unicode(sheet['J31'].value)) , 'description': unidecode(unicode(sheet['J32'].value)).replace("\n","")  },
        {'company': unidecode(unicode(sheet['J35'].value)) , 'timeperiod':unidecode(unicode(sheet['J36'].value)) , 'description': unidecode(unicode(sheet['J37'].value)).replace("\n","")  },
        {'company': unidecode(unicode(sheet['J40'].value)) , 'timeperiod':unidecode(unicode(sheet['J41'].value)) , 'description': unidecode(unicode(sheet['J42'].value)).replace("\n","")  },
        {'company': unidecode(unicode(sheet['J45'].value)) , 'timeperiod':unidecode(unicode(sheet['J46'].value)) , 'description': unidecode(unicode(sheet['J47'].value)).replace("\n","")  },
        {'company': unidecode(unicode(sheet['J50'].value)) , 'timeperiod':unidecode(unicode(sheet['J51'].value)) , 'description': unidecode(unicode(sheet['J52'].value)).replace("\n","")  },
        {'company': unidecode(unicode(sheet['J55'].value)) , 'timeperiod':unidecode(unicode(sheet['J56'].value)) , 'description': unidecode(unicode(sheet['J57'].value)).replace("\n","")  },
        {'company': unidecode(unicode(sheet['J60'].value)) , 'timeperiod':unidecode(unicode(sheet['J61'].value)) , 'description': unidecode(unicode(sheet['J62'].value)).replace("\n","") },
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
        'contrib_ct' : contrib_ct,
        'listevents' :
        [
        eventlist
        ]
        ,
        'event_ct' : event_ct,
        'listpubs' :
        [
        publicationlist
        ]
        ,
        'pub_ct' : pub_ct,
        'projects' :
        [
        projectchain
        ]
        ,

        'contributions' : unidecode(unicode(sheet['L77'].value)).replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br/>").replace(",","&#44;"),

        'conferences' : unidecode(unicode(sheet['I66'].value)).replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br/>").replace(",","&#44;"),

        'fierte' : unidecode(unicode(sheet['L91'].value)).replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br/>").replace(",","&#44;"),

        'loisirs' : unidecode(unicode(sheet['N55'].value)).replace(":","&#58;").replace("*","&#149;").replace("'","&#145;").replace("\n","<br/>").replace(",","&#44;"),






        'lang' : 'fr',
        }






        if ((trigram in tri_contributeurs) or (tri_contributeurs==[])):
            with open('./yamlfiles/' + unidecode(trigram) + '_tmp.yml', 'w') as file:
                file.write("/* #*/ export const PERSON = ` \n")
                yaml.unicode_supplementary=False
                documents = yaml.dump(dict_file, file)
                file.write("\n`")
                file.close()

            with open('./yamlfiles/' + unidecode(trigram) + '_tmp.yml','r') as fichier:
                #print unidecode(trigram)
                #with open('./yamlfiles/' + unidecode(sheet['C11'].value) + '.yml','w') as fichierFinale:
                if contrib_ct==1:
                    triplus="contrib_"
                else:
                    triplus=""
                
                with open('../resume/' + triplus + unidecode(trigram)  + '.yml','w') as fichierFinale:
                    lignes = fichier.readlines()                # On parcours les lignes du fichier source
                    for ligne in lignes:
                        lignetmp=ligne.replace("['","[")
                        lignetmp=lignetmp.replace("''None","'None")
                        lignetmp=lignetmp.replace("project: ''","project: '")
                        lignetmp=lignetmp.replace("'', function","', function")
                        lignetmp=lignetmp.replace("missions: ''","missions: '")
                        lignetmp=lignetmp.replace("technos: ''","technos: '")
                        #lignetmp=lignetmp.replace("''","'")
                        lignetmp=lignetmp.replace("''|","'|")
                        lignetmp=lignetmp.replace("''},","'},")
                        lignetmp=lignetmp.replace("'',","',")
                        lignetmp=lignetmp.replace("\"]","]")
                        lignetmp=lignetmp.replace("[\"","[")
                        lignetmp=lignetmp.replace("\"|","|")
                        lignetmp=lignetmp.replace("\"","")
                        lignetmp=lignetmp.replace("'","'")
                        lignetmp=lignetmp.replace("'{","")
                        lignetmp=lignetmp.replace("`o"," ") 
                        lignetmp=lignetmp.replace("&#46;0","")
                        lignetmp=lignetmp.replace("''Description","Description")
                        lignetmp=lignetmp.replace("!!python/unicode","")
                        lignetmp=lignetmp.replace("[ name","[ {name") 
                        lignetmp=lignetmp.replace("- {company: None, description: None, timeperiod: None}","")
                        ligneFinale=lignetmp.replace("']","]")
                        

                        fichierFinale.write(ligneFinale)               # On écrit la nouvelle ligne dans le nouveau fichier
                fichierFinale.close()                     # Fermeture du fichier source
            fichier.close()
            os.remove('./yamlfiles/' + unidecode(trigram) + '_tmp.yml')
            contrib_presents.append(trigram)
        else :
            #print unidecode(sheet['C11'].value) 
            contrib_absents.append(trigram)

for c in tri_contributeurs:
    if c not in contrib_presents:
        contrib_absents_finaux.append(c)
print contrib_absents_finaux

# le run export genere les PDF / le run dev permet de generer les yml et de les consulter en localhost
os.system("npm run export")
#os.system("npm run dev")
