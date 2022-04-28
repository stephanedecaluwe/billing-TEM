# -*- coding: utf-8 -*-

import selenium
import time
import datetime
import zipfile
from zipfile import ZipFile
import sys
from os import listdir
from os.path import isfile, join
import shutil
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException,TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import xlrd
from openpyxl import Workbook
import os

wait_court = 2
wait_long = 60

def xpathclick(browser,xpath):
   time.sleep(2)
   dot()
   browser.find_element_by_xpath(xpath).click()
   return
   
def dot():
   print('.',end='')
   return

def tiny_file_rename(newname, folder_of_download):
    filename = max([f for f in os.listdir(folder_of_download)], key=lambda xa :   os.path.getctime(os.path.join(folder_of_download,xa)))
    if '.part' in filename:
        time.sleep(1)
        os.rename(os.path.join(folder_of_download, filename), os.path.join(folder_of_download, newname))
    else:
        os.rename(os.path.join(folder_of_download, filename),os.path.join(folder_of_download,newname))
    return
   



path_lut=r"C:\\Users\sdecaluwe\Desktop\TEM_CONNECT local\data\LUT.xlsx"
liste_regions=[]

file_input=path_lut
feuille = xlrd.open_workbook(str(file_input))
onglet = feuille.sheet_by_name(u'matrice clients TEM')
for row in range(onglet.nrows):
   region = onglet.row_values(row)[4]
   if region != "region TEM" and region != "" and "*" not in region:
      liste_regions.append(region)
      liste_regions=list(set(liste_regions))
      

#print(liste_regions)




for region in liste_regions:
#for region in ['FR_SATIN']:
#liste de "petits clients" pour ajuster les sleep
#en clair et dans le desordre :for region in['FR_AVEM', 'FR_ITBS', 'FR_JDC', 'FR_CMCIC', 'FR_REUNION_TELECOM', 'FR_BNP', 'FR_AVT','FR_EXM', 'FR_LA POSTE', 'FR_BPN', 'FR_MONEY30', 'FR_SYNALCOM', 'FR_IPSF', 'FR_BPRI', 'FR_SATIN', 'FR_HMTELECOM','FR_CLIENTS_FRANCE_OPERES']:
   try:
      print(region)
      
      options = webdriver.ChromeOptions()
      options.add_argument('--https://estate-manager-eu.icloud.ingenico.com/emgui/')    #proxy-server=http://85.115.60.150:80

      #paramètres chrome forcer l'affichage plein écran, sinon, 'MODULE_REPORTS' n'est pas reconnu
      options.add_argument('--start-maximized')
      options.add_experimental_option("prefs", {"download.default_directory": r'C:\Users\Public\Downloads',"download.prompt_for_download": False,   #c:\Users\sdecaluwe\Desktop\TEM_CONNECT local\downloads
      "download.directory_upgrade": True,"safebrowsing.enabled": False})

      browser = webdriver.Chrome(r'C:\\Users\sdecaluwe\Desktop\TEM_CONNECT local\chromedriver\chromedriver.exe',chrome_options=options)
      browser.get('https://estate-manager-eu.icloud.ingenico.com/emgui/')

      #login, password, valider
      browser.find_element_by_xpath('//*[@id="username"]').send_keys('Facturation')
      browser.find_element_by_xpath('//*[@id="password"]').send_keys('Facturation8765?')
      browser.find_element_by_xpath('//html/body/div/div/form/div[1]/div/input').click()




      # menu global rerporting
      xpathclick(browser,'//*[@id="MODULE_REPORT_GLOBAL_REPORTING"]/a')
      xpathclick(browser,'//*[@id="availableReportsGroup"]/button')
      xpathclick(browser,'//*[@id="-1dd2ea73:16b4b81f1a1:112b:AC150004"]/label')       
      xpathclick(browser,'//*[@id="saved-search-4b1bdc82:17ef8fb28c7:7b5b:AC1A2E24"]')
      #xpathclick(browser,'//*[@id="saved-search--7ba16025:171125f6806:-1718:AC1A2E24"]/span')  #//*[@id="saved-search--b4a06dc:16e83084f28:1b1b:AC1A2E24"]
      #//*[@id="saved-search-4b1bdc82:17ef8fb28c7:7b5b:AC1A2E24"]

      #
      xpathclick(browser,'//*[@id="sponsors-select-div"]/div/div/button')
      #input region a charger : ex FR_SATIN
      browser.find_element_by_xpath('//*[@id="sponsors-select-div"]/div/div/ul/li[1]/div/input').send_keys(region)
      #cocher tout le monde
      time.sleep(5)
      xpathclick(browser,'//*[@id="sponsors-select-div"]/div/div/ul/li[@class="dropdown-tree-parent"]/span/input')
      #chercher
      xpathclick(browser,'//*[@id="apply-search"]')
      time.sleep(30)
      dot()
      # telecharger tous les enregistrements
      xpathclick(browser,'//*[@id="export-all-btn"]')
      time.sleep(10)
      suivi_download = browser.find_element_by_xpath('//*[@id="generated-reports-table"]/tbody/tr[1]/td[3]').text
      while suivi_download != "Terminé" and suivi_download !="Finished":
         dot()
         time.sleep(30)
         suivi_download = browser.find_element_by_xpath('//*[@id="generated-reports-table"]/tbody/tr[1]/td[3]').text
      #print(browser.find_element_by_xpath('//*[@id="generated-reports-table"]/tbody/tr[1]/td[3]').text)
      #time.sleep(wait_court)
      #print(browser.find_element_by_xpath('//*[@id="generated-reports-table"]/tbody/tr[1]/td[3]').text)
      #time.sleep(wait_court)
      #print(browser.find_element_by_xpath('//*[@id="generated-reports-table"]/tbody/tr[1]/td[3]').text)
      #time.sleep(wait_court)
      # pour attraper le menu contextuel sans mouseover : exporter le rapport
      browser.find_element_by_class_name('infobulle').click()
      browser.find_element_by_xpath('//*[@id="generated-reports-table"]/tbody/tr[1]/td[5]/div/span/a[1]').click()
      dot()
      time.sleep(wait_court)
      # supprimer le rapport apres l avoir telecharge : oui
      browser.find_element_by_xpath('//*[@id="confirmDeleteDownloadedReport"]/div[2]/div/div/form/button[2]').click()
      dot()
      time.sleep(wait_long*2)
      dot()
      browser.quit()
      tiny_file_rename(region[3:]+'.zip',r"C:\\Users\Public\Downloads")
      #dest =os.path.join(region[3:]+'.zip',r"\\Frprfil\data\FranceTeam\Marketing&Communication\14. Projets\TEM\facturation TEM\automatisation extraction facturation\downloads") autre maniere
      newname = r"C:\\Users\Public\Downloads" + "\\" + region[3:]+ ".zip"
      print(newname)
      #newname =os.path.join(r"\\Frprfil\data\FranceTeam\Marketing&Communication\14. Projets\TEM\facturation TEM\automatisation extraction facturation\downloads" + "\\" + region[3:]+ ".zip") autre maniere
      
      
      #dezipper renommer mettre dans input
      zip = zipfile.ZipFile(newname)   #zip = zipfile.ZipFile(newname) 
      
      
      #os.path.exists( dossier ou fichier)
      zip.extractall()
      #into script's directory !!
      print(r"C:\\Users\Public\Downloads" + "\\" + region[3:])
      shutil.copy(r"C:\\Users\Public\Downloads\Multi customer.xlsx" , "C:\\Users\Public\Downloads\Input")
      print(r"C:\\Users\Public\Downloads" + "\\" + region[3:]+ "\\")
      tiny_file_rename(region[3:]+'.xlsx',r"C:\\Users\Public\Input")
      shutil.copy(r"C:\\Users\Public\Input\\"+region[3:]+'.xlsx' , r"\\"+"\\coprfil.usr.ingenico.loc\Public\ProfessionalServices_France\TEM\facturation TEM\TEM_CONNECT\input")
      dot()
      print(region)
   except Exception as exc:
      print("exception de type ", exc.__class__)
      print("message", exc)
      print('import raté')
      browser.quit()
