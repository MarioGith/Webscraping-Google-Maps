from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import xlwt

## ------------- Création d'une interface graphique simple à base de print et d'input pour récupérer les informations -------------
print("\033[1m"+"Bienvenue sur le scrapeur Google Maps d'Itsemi"+"\033[0m")
print('')
print("Ce parseur à pour but de récupérer les noms, adresses et numéros de téléphone des entreprises d'un certains type dans une zone prédéfinis")
print("Pour cela, il est nécessaire de se rendre au préalable sur Google Maps, de rentrer une adresse et/ou ville afin de repérer la zone en comptant les zoom/dézoom effectué")
print('')
print("\x1B[3m"+"Attention, ce scrapeur fonctionne uniquement en journée lorsque les entreprises sont ouvertes, sinon le numéro de téléphone ne se récupère pas"+"\x1B[23m")
print('')
print("\x1B[3m"+"Merci de répondre aux questions suivantes avec beacoup de minutie : "+"\x1B[23m")

## ------------- Récupération de l'adresse -------------
adresse = input('Quelle adresse ? : ')
code_postal = input('Quel code postal ? : ')
ville = input('Quelle ville ? : ')

## ------------- Gestion du zoom et dézoom pour permettre la recherche par zone
print("Si vous souhaitez zommer ou dézomer, merci d'indiquer de combien de click en + ou -")
zoom = input('Nombre de click ? ' + "\x1B[3m"+"(vide si aucun zoom/dézoom) : "+"\x1B[23m")
if (zoom == ''):
	zoom = 0
else:
	zoom = int(zoom)

## ------------- Récupération de l'entreprise -------------
entreprise = input("Quelle type d'entreprise ? : ")
while(len(entreprise) == 0):
	entreprise = input("Quelle type d'entreprise ? : ")

## ------------- Récupération du nom du fichier -------------
filename = input("Nom de l'excel de sortie ? : ")
while(len(filename) == 0):
	filename = input("Nom de l'excel de sortie ? : ")

## ------------- Path du geckodriver -------------
gecko_path = input("Path du geckodriver ? : ")
while(len(gecko_path) == 0):
	gecko_path = input("Path du geckodriver ? : ")
print('\n')


## ------------- Initialisation de Selenium -------------
driver = webdriver.Firefox(executable_path=gecko_path)
driver.get("https://www.google.com/maps")
driver.maximize_window()

## ------------- Direction vers Google Maps -------------
driver.find_element_by_xpath("/html/body/div/c-wiz/div/div/div/div[2]/div[1]/div[4]/form/div[1]/div/button/span").click()

## ------------- Gestion de la barre de recherche pour l'adresse -------------
searchBar = driver.find_element_by_xpath('//*[@id="searchboxinput"]')
searchBar.clear()
searchBar.send_keys(adresse+" "+ville+" "+code_postal)
searchBar.send_keys(Keys.ENTER)
time.sleep(4)

## ------------- Gestion du zoom -------------
if(zoom>0):
	for i in range(zoom):
		driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[22]/div[1]/div[2]/div[7]/div/div[1]/button/div").click()
		searchBar.click()
		time.sleep(1)
if(zoom<0):
	for i in range(-zoom):
		driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[22]/div[1]/div[2]/div[7]/div/button/div").click()
		searchBar.click()		
		time.sleep(1)
indice_zoom_initial = driver.find_element_by_xpath('//*[@id="gAWHhb-scale-V67aGc"]').text

## ------------- Gestion de la barre de recherche pour le type d'entreprise -------------
searchBar.clear()
searchBar.send_keys(entreprise)
searchBar.send_keys(Keys.ENTER)
time.sleep(5)

## ------------- Gestion du plus proche -------------
clickable = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[2]/div[1]/div/button/span[1]')

## ------------- Récupération des données -------------

## ------------- Initialisation du tableau des centres -------------
centres = []

## ------------- Fonction de tri -------------
def tri_details (tableau):
	tab = []
	for detail in tableau:
		if('Aucun avis' not in detail.text):
			if('(' not in detail.text):
				if ("\n" in detail.text):
					tab.append(detail.text)
	return tab

## ------------- Gestion par pages -------------
while (driver.find_element_by_xpath('//*[@id="gAWHhb-scale-V67aGc"]').text == indice_zoom_initial): 
	## ------------- Gestion du scroll pour charger les résultats -------------
	pause_time = 2
	max_count = 4
	scrollable_div = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[9]/div[8]/div/div[1]/div/div/div[4]/div[1]")
	for i in range(max_count):
		driver.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', scrollable_div)
		time.sleep(pause_time)

	## ------------- Récupération des noms/détails des centres
	time.sleep(1)
	nom_centres = driver.find_elements_by_class_name("qBF1Pd-haAclf")
	details_centres = driver.find_elements_by_class_name("ZY2y6b-RWgCYc")
	time.sleep(2)
	details = tri_details(details_centres)

	## Ajout du centre 
	for i in range (len(nom_centres)):
		print('-------------')
		nom = nom_centres[i].text
		print(nom)
		try: 
			_adresse = details[i].split('· ')[1]
			adresse = _adresse.split("\n")[0]
		except:
			adresse = 'Bug de récupération'
			pass
		print(adresse)
		try:
			numero = details[i].split('· ')[2]
		except:
			numero = 'Pas de numéro'
			pass
		print(numero)
		centres.append([nom, adresse, numero])
		

	print('-------------')
	driver.find_element_by_xpath('//*[@id="ppdPk-Ej1Yeb-LgbsSe-tJiF1e"]/img').click()
	time.sleep(3)
	print("Page suivante")

print('Plus de centre dans le secteur défini')
print(' ')


## ------------- Initialisation de l'Excel -------------
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Import Hubspot')
sheet.write(0, 0, "Nom de l'entreprise")
sheet.write(0, 1, 'Numéro de téléphone')
sheet.write(0, 2, 'Adresse')
sheet.write(0, 3, 'Code Postal')
sheet.write(0, 4, "City")

for i in range(1, len(centres)):
	sheet.write(i, 0, centres[i][0]) # Nom de l'entreprise
	sheet.write(i, 1, centres[i][2]) # Numéro de téléphone
	sheet.write(i, 2, centres[i][1]) # Adresse
	sheet.write(i, 3, code_postal) # Code postal
	sheet.write(i, 4, ville) # Ville
	sheet.write(i, 5, entreprise) # Prescripteur

print("Merci d'avoir utilisé mon scrapeur \nÀ très bientôt sur mes outils !")
print("\x1B[3m"+"Itsemi"+"\x1B[23m")
workbook.save(filename+'.xls')