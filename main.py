import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox
import openpyxl
import re


def scrap_website(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')

            #Podaje całą stronę i w ifach poniższych wyszukuje w całej treści strony -
            #do poprawy żeby tylko z linka sprawdzał IF
            #print(soup.text)

            # Najpierw sprawdź, czy strona zawiera '...'
            if '3kropki.pl' in soup.text:
                element = soup.find(class_="k3_font_01 k3_pbc_price")
                if element:
                    return f"Cena: {element.text.strip()}"
                else:
                    return "Nie znaleziono ceny."

            if 'masteredukacja.pl' in soup.text:
                element = soup.find(itemprop="price")
                if element:
                    #return f"Cena: {element.text.strip()}"
                    return element.text.strip()
                else:
                    return "Nie znaleziono ceny."

            if 'ŚwiatProgramow.pl' in soup.text:
                element = soup.find(id="st_product_options-price-brutto")
                if element:
                    # Bezpośrednio przekształcamy zawartość tekstową elementu na string
                    #  element_text = element.text.strip()  # Usuwamy białe znaki na początku i na końcu
                    #cleaned_text = element_text.replace(" *", "")  # Usuwamy niechciane znaki
                    #  return f"Cena: {cleaned_text}"
                    return f"Cena: {element.text.strip()}"
                else:
                    return "Nie znaleziono ceny."

            if 'harpo.com' in soup.text:
                element = (soup.find_all('span', class_='woocommerce-Price-amount amount'))
                if len(element) >= 2:
                    # Wybieranie drugiego elementu z listy
                    second_price_element = element[1]
                    # Pobranie zawartości tekstowej elementu <bdi>
                    price_text = second_price_element.bdi.text.strip()
                    price_text = price_text.replace(".", "")
                    return f"Cena: {price_text}"
                else:
                    return "Nie znaleziono ceny."


            if 'skleporto' in soup.text:
                element = soup.find(itemprop="price")
                if element:
                    return f"Cena: {element.text.strip()}"
                else:
                    return "Nie znaleziono ceny."

            if 'rerek.pl' in soup.text:
                element = soup.find(class_="st_product_options-price-brutto")
                if element:
                    return f"Cena: {element.text.strip()}"
                else:
                    return "Nie znaleziono ceny."

            if 'arante.pl' in soup.text:
                element = soup.find(id="st_product_options-price-brutto")
                # print("Sprawdzam Arante")
                if element:
                    #return f"Cena: {element.text.strip()}"
                    return element.text.strip()
                else:
                    return "Nie znaleziono ceny."
        else:
            return "Nie udało się połączyć z stroną."
    except Exception as e:
        return "Wystąpił błąd: " + str(e)


def on_scrap():
    url = url_entry.get()
    result = scrap_website(url)
    if result is None:
        result = "Brak wyniku."
    result_text.insert(tk.INSERT, result + "\n")


def on_import_urls():
    filepath = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if not filepath:
        return

    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active
    result_text.delete(1.0, tk.END)

    for row in sheet.iter_rows(min_row=1, values_only=True):
        product_id, our_url, competitor_urls = row[0], row[1], row[2]
        print(f"\nPrzetwarzanie produktu {product_id}...")

        result = scrap_website(our_url)

        ourPrice = result
        ourPriceFloat = ourPrice.strip()  # Usuwamy białe znaki na początku i na końcu
        #print(ourPriceFloat)
        ourPriceFloat = ourPriceFloat.replace(" zł", "")  # Usuwamy niechciane znaki
        #print(ourPriceFloat)
        ourPriceFloat = ourPriceFloat.replace(",", ".")
        #print(ourPriceFloat)
        ourPriceFloat = ourPriceFloat.replace(" ", "")
        float(ourPriceFloat)
        print(ourPriceFloat)

        print(f"ID Produktu: {product_id}, URL: {our_url}, {result}")
        result_text.insert(tk.INSERT, f"ID Produktu: {product_id}, URL: {our_url}, {result}\n")

        if competitor_urls:
            competitor_urls_list = competitor_urls.split(';')
            for url in competitor_urls_list:
                if url:  # Dodatkowo sprawdzamy, czy URL nie jest pusty
                    result = scrap_website(url.strip())

                    concurentPrice = result
                    concurentPriceFloat = concurentPrice.strip()  #Usuwamy białe znaki na początku i na końcu
                    print(concurentPriceFloat)
                    concurentPriceFloat = concurentPriceFloat.replace("zł", "")  #Usuwamy niechciane znaki
                    concurentPriceFloat = concurentPriceFloat.replace("Cena: ", "")
                    print(concurentPriceFloat)
                    concurentPriceFloat = concurentPriceFloat.replace(",", ".")
                    print(concurentPriceFloat)
                    concurentPriceFloat = concurentPriceFloat.replace(" ", "")
                    concurentPriceFloat = re.sub(r'\s+', '', concurentPriceFloat)
                    # float(concurentPriceFloat)
                    print(concurentPriceFloat)
                    minPrice = ourPriceFloat
                    if concurentPriceFloat < ourPriceFloat:
                        minPrice = concurentPriceFloat
                        print(f"Minimalna cena dla {product_id}: {url.strip()}, {minPrice}")

                    if result is None:
                        result = "Brak wyniku"

                    print(f"Konkurencyjny URL dla {product_id}: {url.strip()}, {result}")
                    result_text.insert(tk.INSERT, f"Konkurencyjny URL dla {product_id}: {url.strip()}, {result}")
        else:
            print(f"Brak linków konkurencji dla produktu {product_id}.")


def on_export_results():
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if not filepath:
        return

    wb = openpyxl.Workbook()
    sheet = wb.active

    content = result_text.get("1.0", tk.END)
    lines = content.split("\n")

    for i, line in enumerate(lines):
        if line:  # Pomiń puste linie
            sheet.cell(row=i + 1, column=1, value=line)

    wb.save(filepath)
    messagebox.showinfo("Export Complete", "Wyniki zostały zapisane do pliku Excel.")


# Ustawienie głównego okna aplikacji
root = tk.Tk()
root.title("Scraper Cen IXION v2024.03.21")

tk.Label(root, text="URL:").pack(pady=5)
url_entry = tk.Entry(root, width=50)
url_entry.pack(pady=5)

scrap_button = tk.Button(root, text="Scrap Website", command=on_scrap)
scrap_button.pack(pady=5)

import_button = tk.Button(root, text="Import URLs from Excel", command=on_import_urls)
import_button.pack(pady=5)

export_button = tk.Button(root, text="Export Results to Excel", command=on_export_results)
export_button.pack(pady=5)

result_text = scrolledtext.ScrolledText(root, height=10, width=50)
result_text.pack(pady=10)

#root.config(bg="#26242f")

root.mainloop()
