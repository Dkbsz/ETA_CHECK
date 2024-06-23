import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time


class MSCU:
    url = "https://www.msc.com/en/track-a-shipment",
    input_id = "trackingNumber"

    def __init__(self, driver):
        self.driver = driver

    def accept_cookies(self):
        accept_button = self.driver.find_element(By.ID, "onetrust-accept-btn-handler")
        accept_button.click()

    def get_results(self, ctds):
        xpath = '//*[@id="main"]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[3]/div/div/div[1]/div/div[4]/div/div/div/span[2]'
        result = enter_ctds_and_get_results(
            self.driver, self, self.url, ctds, self.input_id, xpath
        )
        return result


class MAEU:
    url = "https://www.maersk.com/tracking/"
    input_id = "trackTrace"

    def __init__(self, driver):
        self.driver = driver

    def accept_cookies(self):
        accept_button = self.driver.find_element(By.XPATH, '//*[@id="coiPage-1"]/div[2]/div/button[2]')
        accept_button.click()

    def get_results(self, ctds):
        xpath = '//*[@id="maersk-app"]/div/div/div[3]/div/div/dl/dd[1]'
        modified_ctds = [
            (index, ctd[4:] if ctd.upper().startswith("MAEU") else ctd)
            for index, ctd in ctds
        ]
        result = enter_ctds_and_get_results(
            self.driver, self, self.url, modified_ctds, self.input_id, xpath
        )
        return result


# Funkcja do wpisywania wartości CTD na stronach przewoźników i spis wyników
def enter_ctds_and_get_results(driver, carrier, url, ctds, input_id, result_xpath) -> dict:
    result = {}
    driver.get(url)
    carrier.accept_cookies()
    for index, ctd in ctds:
        try:
            print(f"Przetwarzanie CTD: {ctd} dla przewoźnika {carrier}")

            # Szukanie formularza do wpisania CTD number
            input_element = driver.find_element(By.ID, input_id)
            input_element.clear()
            print(f"Wpisywanie CTD: {ctd}")
            input_element.send_keys(ctd)
            input_element.send_keys(Keys.RETURN)

            # Pobierz wynik
            result_element = driver.find_element(By.XPATH, result_xpath)
            result_text = result_element.text
            print(f"Otrzymany wynik: {result_text}")

            # Zapisz wynik w słowniku
            result[index] = result_text

        except Exception:
            print(f"Error for CTD {ctd} on {carrier}")
    return result


def main():
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

    # Inicjalizacja przeglądarki oraz rozpoczecie stopera.
    start_time = time.time()
    driver = webdriver.Chrome()  # ver.125.0.6422.142
    driver.maximize_window()  # full screen
    driver.implicitly_wait(5)

    # Pusty słownik do przechowywania wyników
    results = {}
    carriers = {
        "MAEU": MAEU,
        "MSCU": MSCU
    }
    # Przetwarzanie danych dla każdego przewoźnika
    for carrier, ctds in data_dict.items():

        carrier_class = carriers.get(carrier)
        if carrier_class is None:
            raise KeyError(f"{carrier=} is not supported")
        carrier_instance = carrier_class(driver)
        result = carrier_instance.get_results(ctds)
        results.update(result)

    # Zamykanie przeglądarki
    driver.quit()

    # Zapisywanie wyników do nowego excela
    result_df = df.copy()
    result_df["Result"] = result_df.index.map(results)

    print("--- %s seconds ---" % (time.time() - start_time)) #timer wykonania programu

    final_results = "final_results.xlsx"
    result_df.to_excel(final_results, index=False)
    print(f"Wyniki zostały zapisane do pliku {final_results}")


main()
