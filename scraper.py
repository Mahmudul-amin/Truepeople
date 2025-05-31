import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# Set up Chrome driver
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=chrome_options)

# Read address list from Excel
df = pd.read_excel("input.xlsx")
addresses = df[['Address', 'ctz']].dropna()

# Results list
results = []

def safe_find(element, by, value):
    try:
        return element.find_element(by, value)
    except:
        return None
    
for index, row in addresses.iterrows():
    Address = row['Address']
    ctz = row['ctz']  # or whatever the actual column name is

    try:
        print(f"🔍 Searching: {Address} | {ctz}")
        driver.get("https://www.truepeoplesearch.com/")
        time.sleep(random.uniform(30, 40))

        

        # ... continue your code using `address` and `ctz` ...



        # Click "Address" tab
        driver.find_element(By.XPATH, '//*[@id="searchTypeAddress-d"]/span').click()
        time.sleep(random.uniform(10, 15))

        # Enter the address
        driver.find_element(By.XPATH, '//*[@id="id-d-addr"]').send_keys(Address)
        driver.find_element(By.XPATH, '//*[@id="id-d-loc-addr"]').send_keys(ctz)
        driver.find_element(By.XPATH, '//*[@id="btnSubmit-d-addr"]').click()
        time.sleep(random.uniform(5, 8))

        # Select first result
        results_cards = driver.find_elements(By.CLASS_NAME, "card-summary")
        if not results_cards:
            raise Exception("No results found.")

        first_result = results_cards[0]
        name = safe_find(first_result, By.CLASS_NAME, "h4").text.strip()

        all_text = first_result.text.split("\n")
        phone = next((line for line in all_text if "Phone" in line or "Cell" in line), "Not found")
        email = next((line for line in all_text if "Email" in line), "Not found")

        # Click into the detailed profile
        first_result.click()
        time.sleep(random.uniform(4, 6))

        # Look for relative info
        relative_name = "Not found"
        relative_phone = "Not found"
        relative_email = "Not found"

        try:
            relatives_section = driver.find_elements(By.XPATH, "//a[contains(@href, '/find/')]")
            if relatives_section:
                relative_name = relatives_section[0].text.strip()
                relatives_section[0].click()
                time.sleep(random.uniform(3, 5))

                # Get details from relative’s card
                relative_card = driver.find_elements(By.CLASS_NAME, "card-summary")[0]
                rel_text = relative_card.text.split("\n")

                relative_phone = next((line for line in rel_text if "Phone" in line or "Cell" in line), "Not found")
                relative_email = next((line for line in rel_text if "Email" in line), "Not found")
        except Exception as re:
            print("⚠️ No relative info found or failed to extract.")

        # Save the result
        results.append({
            "Address": Address,
            "Name": name,
            "Phone": phone,
            "Email": email,
            "Relative Name": relative_name,
            "Relative Phone": relative_phone,
            "Relative Email": relative_email
        })

    except Exception as e:
        print(f"❌ Error with {Address}: {e}")
        results.append({
            "Address": Address,
            "Name": "Not found",
            "Phone": "Not found",
            "Email": "Not found",
            "Relative Name": "Not found",
            "Relative Phone": "Not found",
            "Relative Email": "Not found"
        })

    # Delay between searches
    time.sleep(random.uniform(4, 8))

# Save results to Excel
pd.DataFrame(results).to_excel("output_results.xlsx", index=False)
print("✅ Done! Results saved to output_results.xlsx")

driver.quit()
