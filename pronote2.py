from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import os

load_dotenv()  # Charge les variables du .env
sheet_url = os.getenv("SHEET_URL")
chrome_user_data_dir = os.getenv("CHROME_USER_DATA_DIR")
pronote_url = os.getenv("PRONOTE_URL")
username = os.getenv("PRONOTE_USERNAME")
password = os.getenv("PRONOTE_PASSWORD")

# web driver configurations
options = Options()
options.add_argument(f"--user-data-dir={chrome_user_data_dir}")
options.add_argument(r"--profile-directory=Default")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")

service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)

# ================ OPEN sheets ================
driver.get(sheet_url)

# Trouver le lien <a> avec le texte "209"
a_209 = driver.find_element(By.XPATH, "//a[normalize-space(text())='209']")
a_209.click()  # si tu veux cliquer dessus

# ================ testing ================
average_td = driver.find_element(By.XPATH, "//td[@class='s1' and normalize-space(text())='AVERAGE']")

# Supposons que tu as déjà ton élément td dans la variable td_element
tr_element = average_td.find_element(By.XPATH, "./ancestor::tr[1]")
td_list = tr_element.find_elements(By.TAG_NAME, "td")
td_index = td_list.index(average_td)
print(f"Index du <td> dans le <tr> : {td_index}")

first_tr = driver.find_element(By.XPATH, "//tbody/tr[1]")

# Supposons que tu as déjà récupéré la balise <tr> dans une variable tr_element
td_elements = first_tr.find_elements(By.TAG_NAME, "td")
print(f"Nombre de <td> dans la <tr> : {len(td_elements)}")

# ================ ================




# Attendre que la page se charge
time.sleep(1)



# ================ OPEN PRONOTE ================
driver.get(pronote_url)

# fullscreen
driver.maximize_window()

# wait for elements to load
wait = WebDriverWait(driver, 3)


# refuse RGPD chrome popup
# try:
#     # Attendre for the popup
#     refuse_button = wait.until(
#         EC.element_to_be_clickable((By.ID, "W0wltc"))
#     )
#     refuse_button.click()
#     print("Bouton 'Refuser tout' cliqué.")
# except:
#     print("Aucun pop-up RGPD détecté.")

# login
input_element = wait.until(EC.element_to_be_clickable((By.ID, "username")))
input_element.send_keys(username)

input_element = wait.until(EC.element_to_be_clickable((By.ID, "password")))
input_element.send_keys(password + Keys.ENTER)

# wait for the page to load
wait.until(EC.presence_of_element_located((By.ID, "GInterface.Instances[0].Instances[1]_Combo3")))

# navigate to the "Notes" section
notes_button = wait.until(EC.element_to_be_clickable((By.ID, "GInterface.Instances[0].Instances[1]_Combo3")))
notes_button.click()

# wait for the page to load
wait.until(EC.presence_of_element_located((By.ID, "GInterface.Instances[2].Instances[0].Instances[0].bouton_Edit")))

# click on the dropdown to select the class
class_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "GInterface.Instances[2].Instances[0].Instances[0].bouton_Edit")))
class_dropdown.click()

# Récupère tous les <li> avec role="presentation"
li_elements = driver.find_elements(By.XPATH, "//li[@role='presentation']")
time.sleep(1)

for li in li_elements:
    try:
        # Cherche une <div> enfant contenant exactement "209"
        div = li.find_element(By.XPATH, ".//div[normalize-space(text())='209']")
        if div:
            li.click()
            break
    except:
        continue

#get the cell that represents the exam
# Attendre et cliquer sur la div dont aria-label contient "CECI est un ESSAI"
# et qui a les classes "devoir" et "nom-devoir"
exam_div = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//div[contains(@title, 'CECI est un ESSAI') and contains(@class, 'Titre') and contains(@class, 'liste_gridTitre_cel') and contains(@class, 'b-right')]"
    ))
)

column = driver.execute_script(
    "return window.getComputedStyle(arguments[0]).getPropertyValue('grid-column');",
    exam_div
)
print(f"grid-column: {column}")

# input("Appuie sur Entrée pour continuer...")

#================= GET THE INPUT FOR A STUDENT ==================

# get the div with the value "BRENIEUX Inès"
time.sleep(1)
student_div = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//div[contains(text(), 'BRENIAUX Inès') and contains(@class, 'liste_contenu_ligne')]"
    ))
)

# Remonter de 3 niveaux dans la hiérarchie DOM
parent3 = student_div.find_element(By.XPATH, "../../..")

row = parent3.get_attribute("aria-rowindex")

input_row = int(row) - 2
input_column = int(column) + 4
input_id = f"GInterface.Instances[2].Instances[1]_{input_column}_{input_row}_div"

# get the input element"
input_element = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, input_id))
)

# Clique sur la div pour lui donner le focus
input_element.click()

# Envoie la note et Entrée au navigateur
actions = ActionChains(driver)
actions.send_keys("15").send_keys(Keys.ENTER).perform()

time.sleep(30)
driver.quit()