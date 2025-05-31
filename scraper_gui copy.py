import time
import random
import pandas as pd
import threading
from tkinter import Tk, Button, Text, filedialog, Scrollbar, END, RIGHT, Y, LEFT, BOTH
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# GUI Setup
def start_gui():
    def log(message):
        output_text.insert(END, message + "\n")
        output_text.see(END)
        root.update()

    def browse_file():
        file_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))
        log(f"📄 Selected: {file_path.get()}")

    def start_scraping():
        threading.Thread(target=run_scraper, daemon=True).start()

    def run_scraper():
        try:
            df = pd.read_excel(file_path.get())
            addresses = df[['Address', 'ctz']].dropna()
        except Exception as e:
            log(f"❌ Failed to read Excel: {e}")
            return

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=chrome_options)

        results = []

        def safe_find(element, by, value):
            try:
                return element.find_element(by, value)
            except:
                return None

        for index, row in addresses.iterrows():
            Address = row['Address']
            ctz = row['ctz']

            try:
                log(f"🔍 Searching: {Address} | {ctz}")
                driver.get("https://www.truepeoplesearch.com/")
                time.sleep(random.uniform(5, 8))

                # CAPTCHA pause
                while True:
                    if "attention required" in driver.title.lower() or "captcha" in driver.page_source.lower():
                        log("⚠️ CAPTCHA detected! Please solve manually in the browser.")
                        input("✅ Press ENTER in console after solving CAPTCHA...")
                        time.sleep(2)
                    else:
                        break

                driver.find_element(By.XPATH, '//*[@id="searchTypeAddress-d"]/span').click()
                time.sleep(random.uniform(10, 15))
                driver.find_element(By.XPATH, '//*[@id="id-d-addr"]').send_keys(Address)
                driver.find_element(By.XPATH, '//*[@id="id-d-loc-addr"]').send_keys(ctz)
                driver.find_element(By.XPATH, '//*[@id="btnSubmit-d-addr"]').click()
                time.sleep(random.uniform(5, 8))

                results_cards = driver.find_elements(By.CLASS_NAME, "card-summary")
                if not results_cards:
                    raise Exception("No results found.")

                first_result = results_cards[0]
                name = safe_find(first_result, By.CLASS_NAME, "h4").text.strip()
                all_text = first_result.text.split("\n")
                phone = next((line for line in all_text if "Phone" in line or "Cell" in line), "Not found")
                email = next((line for line in all_text if "Email" in line), "Not found")

                first_result.click()
                time.sleep(random.uniform(4, 6))

                relative_name = "Not found"
                relative_phone = "Not found"
                relative_email = "Not found"

                try:
                    relatives_section = driver.find_elements(By.XPATH, "//a[contains(@href, '/find/')]")
                    if relatives_section:
                        relative_name = relatives_section[0].text.strip()
                        relatives_section[0].click()
                        time.sleep(random.uniform(3, 5))
                        relative_card = driver.find_elements(By.CLASS_NAME, "card-summary")[0]
                        rel_text = relative_card.text.split("\n")
                        relative_phone = next((line for line in rel_text if "Phone" in line or "Cell" in line), "Not found")
                        relative_email = next((line for line in rel_text if "Email" in line), "Not found")
                except:
                    log("⚠️ No relative info found.")

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
                log(f"❌ Error with {Address}: {e}")
                results.append({
                    "Address": Address,
                    "Name": "Not found",
                    "Phone": "Not found",
                    "Email": "Not found",
                    "Relative Name": "Not found",
                    "Relative Phone": "Not found",
                    "Relative Email": "Not found"
                })

            time.sleep(random.uniform(4, 8))

        driver.quit()
        pd.DataFrame(results).to_excel("output_results.xlsx", index=False)
        log("✅ Scraping completed. Results saved to output_results.xlsx")

    # GUI Layout
    root = Tk()
    root.title("TruePeopleSearch Scraper GUI")
    root.geometry("700x500")

    file_path = tk.StringVar()

    Button(root, text="📂 Select Excel File", command=browse_file).pack(pady=10)
    Button(root, text="🚀 Start Scraping", command=start_scraping).pack(pady=5)

    scrollbar = Scrollbar(root)
    scrollbar.pack(side=RIGHT, fill=Y)

    output_text = Text(root, wrap="word", yscrollcommand=scrollbar.set)
    output_text.pack(expand=True, fill=BOTH, padx=10, pady=10)
    scrollbar.config(command=output_text.yview)

    root.mainloop()

# Launch GUI
if __name__ == "__main__":
    import tkinter as tk
    start_gui()
