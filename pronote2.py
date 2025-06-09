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

notes = {}

# ================ OPEN sheets ================
driver.get(sheet_url)

# cliquer sur le lien <a> avec le texte "209"
a_209 = driver.find_element(By.XPATH, "//a[normalize-space(text())='209']")
a_209.click()

tbody = driver.find_element(By.XPATH, "//div[@id='1309140650']//tbody")
student_count = len(tbody.find_elements(By.TAG_NAME, "tr")) -1 # -1 to exclude the header row

student_td_index = int(tbody.find_element(By.XPATH, ".//td[@class='s0' and normalize-space(text())='STUDENT']").get_attribute("cellIndex"))
avg_td_index = int(tbody.find_element(By.XPATH, ".//td[@class='s1' and normalize-space(text())='AVERAGE']").get_attribute("cellIndex"))

print(f"Index du <td> 'STUDENT' dans le <tr> : {student_td_index}")
print(f"Index du <td> 'AVERAGE' dans le <tr> : {avg_td_index}")

for i in range(1, student_count + 1):  # start from 1 to skip the header row
    student_td = tbody.find_elements(By.TAG_NAME, "tr")[i].find_elements(By.TAG_NAME, "td")[student_td_index-1]
    avg_td = tbody.find_elements(By.TAG_NAME, "tr")[i].find_elements(By.TAG_NAME, "td")[avg_td_index-1]

    student_name = student_td.text.strip()
    avg_value = avg_td.text.strip()

    if student_name and avg_value:
        if avg_value == "ABS":
            notes[student_name] = avg_value
        elif avg_value == "X":
            continue
        else:
            notes[student_name] = float(avg_value)

print("Notes des étudiants :")
for student, avg in notes.items():
    print(f"{student}: {avg}")

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

# ================ LOGIN ================

input_element = wait.until(EC.element_to_be_clickable((By.ID, "username")))
input_element.send_keys(username)

input_element = wait.until(EC.element_to_be_clickable((By.ID, "password")))
input_element.send_keys(password + Keys.ENTER)

# ================ NAVIGATE TO NOTES ================

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
time.sleep(1)

# ================ FILL NOTES ================

#get the index of the input column for the exam
def getInputColumnIndex(examName):
    #exam_div = driver.find_element(By.XPATH, f"//div[contains(@title, '{examName}') and contains(@class, 'Titre') and contains(@class, 'liste_gridTitre_cel') and contains(@class, 'b-right')]")
    exam_div = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//div[contains(@title, 'CECI est un ESSAI') and contains(@class, 'Titre') and contains(@class, 'liste_gridTitre_cel') and contains(@class, 'b-right')]"
    ))
)
    input("get aria-colindex")
    print(f"exam_div.get_attribute: {exam_div.get_attribute("aria-colindex")}")
    print(f"Exam div: {exam_div.text}")
    return int(1) + 1

# get the location to fill the note and fill it
def fillNote(studentName, columnIndex, note):
    student_cell = driver.find_element(By.XPATH, f"//div[contains(@class, 'liste_contenu_ligne') and contains(text(), '{studentName}')]")
    inputRow = int(student_cell.find_element(By.XPATH, "../../..").get_attribute("aria-rowindex")) -2

    input_id = f"GInterface.Instances[2].Instances[1]_{columnIndex}_{inputRow}_div"
    
    note_cell = driver.find_element(By.ID, input_id)
    note_cell.click()
    
    actions = ActionChains(driver)
    actions.send_keys(str(note)).send_keys(Keys.ENTER).perform()

def fillNotesForAllStudents(examName, notes):
    columnIndex = getInputColumnIndex(examName)
    print(f"Index de la colonne apres : {columnIndex}")
    input("Press Enter to continue...")

    
    for student, note in notes.items():
        input(f"Press Enter to fill the note for {student} in column {columnIndex} with note {note}...")

        if note == "X":
            continue
        elif note == "ABS":
            fillNote(student, columnIndex, "A")
        else:
            fillNote(student, columnIndex, note)

fillNotesForAllStudents("CECI est un ESSAI", notes)






#================= GET THE INPUT FOR A STUDENT ==================

# #get the cell that represents the exam
# exam_div = WebDriverWait(driver, 10).until(
#     EC.element_to_be_clickable((
#         By.XPATH,
#         "//div[contains(@title, 'CECI est un ESSAI') and contains(@class, 'Titre') and contains(@class, 'liste_gridTitre_cel') and contains(@class, 'b-right')]"
#     ))
# )

# column = driver.execute_script(
#     "return window.getComputedStyle(arguments[0]).getPropertyValue('grid-column');",
#     exam_div
# )
# print(f"grid-column: {column}")

# get the div with the value "BRENIEUX Inès"
# time.sleep(1)
# student_div = WebDriverWait(driver, 10).until(
#     EC.element_to_be_clickable((
#         By.XPATH,
#         "//div[contains(text(), 'BRENIAUX Inès') and contains(@class, 'liste_contenu_ligne')]"
#     ))
# )

# # Remonter de 3 niveaux dans la hiérarchie DOM
# parent3 = student_div.find_element(By.XPATH, "../../..")

# row = parent3.get_attribute("aria-rowindex")

# input_row = int(row) - 2
# input_column = int(column) + 4
# input_id = f"GInterface.Instances[2].Instances[1]_{input_column}_{input_row}_div"

# # get the input element"
# input_element = WebDriverWait(driver, 10).until(
#     EC.element_to_be_clickable((By.ID, input_id))
# )

# # Clique sur la div pour lui donner le focus
# input_element.click()

# # Envoie la note et Entrée au navigateur
# actions = ActionChains(driver)
# actions.send_keys("15").send_keys(Keys.ENTER).perform()

time.sleep(30)
driver.quit()