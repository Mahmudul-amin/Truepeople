import time
import random
import pandas as pd
import threading
from tkinter import Tk, Label, Button, filedialog, Text, Scrollbar, END
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# GUI Class
class ScraperGUI:
    def __init__(self, master):
        self.master = master
        master.title("TruePeopleSearch Scraper")

        self.label = Label(master, text="Upload Excel File with Addresses")
        self.label.pack(pady=10)

        self.upload_button = Button(master, text="📁 Upload Excel File", command=self.upload_file)
        self.upload_button.pack()

        self.run_button = Button(master, text="▶️ Run Scraper", command=self.run_scraper, state="disabled")
        self.run_button.pack(pady=5)

        self.text_area = Text(master, height=20, width=80)
        self.text_area.pack(padx=10, pady=10)

        self.scrollbar = Scrollbar(master, command=self.text_area.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.text_area.config(yscrollcommand=self.scrollbar.set)

        self.file_path = None

    def log(self, message):
        self.text_area.insert(END, message + "\n")
        self.text_area.see(END)
        self.master.update()

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_path = file_path
            self.label.config(text=f"✅ File Loaded: {file_path}")
            self.run_button.config(state="normal")

    def run_scraper(self):
        thread = threading.Thread(target=self.scrape)
        thread.start()

    def scrape(self):
        self.log("🚀 Starting Scraper...")

        # Set up Chrome driver
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=chrome_options)

        # Load Excel
        df = pd.read_excel(self.file_path)
        addresses = df['Address'].dropna().tolist()

        results = []

        def safe_find(element, by, value):
            try:
                return element.find_element(by, value)
            except:
                return None

        for i, address in enumerate(addresses, start=1):
            try:
                self.log(f"[{i}/{len(addresses)}] 🔍 Searching: {address}")
                driver.get("https://www.truepeoplesearch.com/")
                time.sleep(random.uniform(8, 10))

                driver.find_element(By.CLASS_NAME, "search-type-link").click()
                time.sleep(random.uniform(10, 15))

                driver.find_element(By.ID, "inputAddress").send_keys(address)
                driver.find_element(By.XPATH, '//button[text()="Search"]').click()
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
                    self.log("⚠️ No relative info found or failed to extract.")

                results.append({
                    "Address": address,
                    "Name": name,
                    "Phone": phone,
                    "Email": email,
                    "Relative Name": relative_name,
                    "Relative Phone": relative_phone,
                    "Relative Email": relative_email
                })

            except Exception as e:
                self.log(f"❌ Error with {address}: {e}")
                results.append({
                    "Address": address,
                    "Name": "Not found",
                    "Phone": "Not found",
                    "Email": "Not found",
                    "Relative Name": "Not found",
                    "Relative Phone": "Not found",
                    "Relative Email": "Not found"
                })

            time.sleep(random.uniform(4, 8))

        pd.DataFrame(results).to_excel("output_results.xlsx", index=False)
        self.log("✅ Done! Results saved to output_results.xlsx")
        driver.quit()

# Run GUI
if __name__ == "__main__":
    root = Tk()
    app = ScraperGUI(root)
    root.mainloop()
