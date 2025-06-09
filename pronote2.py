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
driver.maximize_window()
wait = WebDriverWait(driver, 10)  # Increased timeout

# ================ LOGIN ================
try:
    input_element = wait.until(EC.element_to_be_clickable((By.ID, "username")))
    input_element.clear()
    input_element.send_keys(username)
    
    input_element = wait.until(EC.element_to_be_clickable((By.ID, "password")))
    input_element.clear()
    input_element.send_keys(password + Keys.ENTER)
except Exception as e:
    print(f"Login failed: {e}")
    driver.quit()
    exit()

# ================ NAVIGATE TO NOTES ================
try:
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
    time.sleep(2)
    
except Exception as e:
    print(f"Navigation to the notes page failed: {e}")
    driver.quit()
    exit()

# ================ FILL NOTES ================
def getInputColumnIndex(examName, max_retries=3):
    """Get the column index for a specific exam with retry logic"""
    for attempt in range(max_retries):
        try:
            exam_div = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    f"//div[contains(@title, '{examName}') and contains(@class, 'Titre') and contains(@class, 'liste_gridTitre_cel') and contains(@class, 'b-right')]"
                ))
            )
            
            column = driver.execute_script(
                "return window.getComputedStyle(arguments[0]).getPropertyValue('grid-column');",
                exam_div
            )
            return int(column) + 4
            
        except Exception as e:
            print(f"Attempt {attempt + 1} failed to get column index: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise Exception(f"Failed to get column index after {max_retries} attempts")

def fillNote(studentName, columnIndex, note, max_retries=3):
    """Fill a note for a specific student with retry logic"""
    for attempt in range(max_retries):
        try:
            if attempt > 0:
                print(f"  Retry {attempt + 1}/{max_retries}")
            
            # ÉTAPE 1: Délai minimal pour stabilisation
            time.sleep(0.5 if attempt == 0 else 1.0)
            
            # ÉTAPE 2: Re-chercher et cliquer sur l'étudiant (FRESH ELEMENT)
            student_cell = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.XPATH, 
                    f"//div[contains(@class, 'liste_contenu_ligne') and normalize-space(text())='{studentName}']"
                ))
            )
            
            # Scroll et clic immédiat
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", student_cell)
            time.sleep(0.2)
            student_cell.click()
            time.sleep(0.5)
            
            # ÉTAPE 3: Re-chercher l'élément pour calculer l'ID (FRESH ELEMENT)
            student_cell_for_row = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((
                    By.XPATH, 
                    f"//div[contains(@class, 'liste_contenu_ligne') and normalize-space(text())='{studentName}']"
                ))
            )
            
            parent_row = student_cell_for_row.find_element(By.XPATH, "../../..")
            inputRow = int(parent_row.get_attribute("aria-rowindex")) - 2
            input_id = f"GInterface.Instances[2].Instances[1]_{columnIndex}_{inputRow}_div"
            
            # ÉTAPE 4: Attendre et cliquer sur la cellule de note
            note_cell = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, input_id))
            )
            
            # S'assurer que la cellule est visible et cliquer
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", note_cell)
            time.sleep(0.2)
            note_cell.click()
            time.sleep(0.3)
            
            # ÉTAPE 5: Saisir la note
            actions = ActionChains(driver)
            
            # Sélectionner tout le contenu existant et le remplacer
            actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
            time.sleep(0.1)
            
            # Saisir la nouvelle note
            actions.send_keys(str(note)).perform()
            time.sleep(0.2)
            
            # Confirmer avec Entrée
            actions.send_keys(Keys.ENTER).perform()
            time.sleep(0.8)  # Attendre que la note soit sauvegardée
            
            return True
            
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"  ⚠ Error: {type(e).__name__}")
                time.sleep(1.5)
                
                # Tentative de "reset" rapide de l'état de l'interface
                try:
                    body = driver.find_element(By.TAG_NAME, "body")
                    body.click()
                    time.sleep(0.5)
                    
                    actions = ActionChains(driver)
                    actions.send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.3)
                    
                except Exception:
                    pass  # Ignore les erreurs de reset
            else:
                print(f"  ✗ Failed after {max_retries} attempts: {type(e).__name__}")
                return False

def fillNotesForAllStudents(examName, notes):
    """Fill notes for all students with comprehensive error handling"""
    try:
        columnIndex = getInputColumnIndex(examName)
        print(f"📝 Starting batch fill - Column: {columnIndex}")
        
        successful_fills = 0
        failed_fills = []
        
        # Attendre que l'interface soit stable
        print("⏳ Initializing...")
        time.sleep(1.5)
        
        for i, (student, note) in enumerate(notes.items()):
            print(f"\n[{i+1}/{len(notes)}] {student} → {note}")
            if i == 25:
                input("Press Enter to continue...")
            
            if note == "X":
                print("  ⏭ Skipped")
                continue
            
            # Convertir ABS en A
            actual_note = "A" if note == "ABS" else note
            success = fillNote(student, columnIndex, actual_note)
            
            if success:
                successful_fills += 1
                print("  ✅ Success")
            else:
                failed_fills.append(student)
                print("  ❌ Failed")
                
            # Pause courte entre les étudiants
            if i < len(notes) - 1:
                time.sleep(1.0)
        
        # Résumé final
        print(f"\n{'='*50}")
        print(f"📊 RESULTS: {successful_fills}/{len(notes)} successful")
        if failed_fills:
            print(f"❌ Failed: {', '.join(failed_fills)}")
        else:
            print("🎉 All notes filled successfully!")
        print(f"{'='*50}")
            
        return len(failed_fills) == 0
            
    except Exception as e:
        print(f"💥 Critical error: {type(e).__name__} - {str(e)}")
        return False

def resetInterface():
    """Reset the interface state quickly"""
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.click()
        
        actions = ActionChains(driver)
        actions.send_keys(Keys.ESCAPE).perform()
        
        time.sleep(0.5)
        return True
    except Exception:
        return False

# Execute the main function
try:
    fillNotesForAllStudents("CECI est un ESSAI", notes)
except Exception as e:
    print(f"Program crashed with error: {e}")
finally:
    # Optional: keep browser open for debugging
    # input("Press Enter to close browser...")
    # driver.quit()
    pass

time.sleep(30)
driver.quit()