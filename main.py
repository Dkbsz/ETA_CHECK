import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Odczytanie danych z pliku .xlsx
file_path = "baz21.xlsx"
df = pd.read_excel(file_path)

# Tworzenie słownika z danymi potrzebna do utworzenia nowego excela z wynikami
data_dict = {}
for index, row in df.iterrows():
    carrier = row["CARRIER"]
    ctd = str(row["CTD"])
    if carrier not in data_dict:
        data_dict[carrier] = []
    data_dict[carrier].append((index, ctd))

# Inicjalizacja przeglądarki
driver = webdriver.Chrome()  # ver.125.0.6422.142
driver.maximize_window()  # full screen

# Słownik z adresami URL i XPath dla poszczególnych przewoźników
carrier_info = {
    "MSCU": {
        "url": "https://www.msc.com/en/track-a-shipment",
        "input_id": "trackingNumber",
        "result_xpath": '//*[@id="main"]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[3]/div/div/div[1]/div/div[4]/div/div/div/span[2]',
    },
    "MAEU": {
        "url": "https://www.maersk.com/tracking/",
        "input_id": "trackTrace",
        "result_xpath": '//*[@id="maersk-app"]/div/div/div[3]/div/div/dl/dd[1]',
    },
}

# Pusty słownik do przechowywania wyników
results = {}


# Funkcja do akceptacji cookies
def accept_cookies(driver, carrier):
    try:
        if carrier == "MSCU":
            accept_button = driver.find_element(By.ID, "onetrust-accept-btn-handler")
        elif carrier == "MAEU":
            accept_button = driver.find_element(
                By.XPATH, '//*[@id="coiPage-1"]/div[2]/div/button[2]'
            )
        else:
            return
        accept_button.click()
        time.sleep(2)  # Czekaj chwilę na przetworzenie akceptacji
    except Exception:
        print(f"Could not accept cookies for {carrier}")


# Funkcja do wpisywania wartości CTD na stronach przewoźników i spis wyników
def enter_ctds_and_get_results(driver, carrier, url, ctds, input_id, result_xpath):
    driver.get(url)
    accept_cookies(
        driver, carrier
    )  # < poprzednio użyta funkcja do akceptowania ciasteczek
    for index, ctd in ctds:
        try:
            print(f"Przetwarzanie CTD: {ctd} dla przewoźnika {carrier}")

            # Szukanie formularza do wpisania CTD number
            input_element = driver.find_element(By.ID, input_id)
            input_element.clear()
            print(f"Wpisywanie CTD: {ctd}")
            input_element.send_keys(ctd)
            input_element.send_keys(Keys.RETURN)

            # Czekaj na załadowanie wyników
            time.sleep(2 if carrier == "MSCU" else 5)  # Strona MAEU dłużej się ładuje.

            # Pobierz wynik
            result_element = driver.find_element(By.XPATH, result_xpath)
            result_text = result_element.text
            print(f"Otrzymany wynik: {result_text}")

            # Zapisz wynik w słowniku
            results[index] = result_text

        except Exception:
            print(f"Error for CTD {ctd} on {carrier}")


# Przetwarzanie danych dla każdego przewoźnika
for carrier, ctds in data_dict.items():
    if carrier in carrier_info:
        info = carrier_info[carrier]
        url = info["url"]
        input_id = info["input_id"]
        result_xpath = info["result_xpath"]

        if carrier == "MAEU":
            # Usuwanie prefiksu 'MAEU' z CTD number < inaczej strona nie wyświetli karty przesyłki
            modified_ctds = [
                (index, ctd[4:] if ctd.upper().startswith("MAEU") else ctd)
                for index, ctd in ctds
            ]
            enter_ctds_and_get_results(
                driver, carrier, url, modified_ctds, input_id, result_xpath
            )
        else:
            enter_ctds_and_get_results(
                driver, carrier, url, ctds, input_id, result_xpath
            )

# Zamykanie przeglądarki
driver.quit()

# Zapisywanie wyników do nowego excela
result_df = df.copy()
result_df["Result"] = result_df.index.map(results)

final_results = "final_results.xlsx"
result_df.to_excel(final_results, index=False)
print(f"Wyniki zostały zapisane do pliku {final_results}")
