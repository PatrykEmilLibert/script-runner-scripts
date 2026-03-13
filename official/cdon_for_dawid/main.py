import xml.etree.ElementTree as ET
import openpyxl
import sys
import re
import requests
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os

# --- Core XML Processing and Excel Generation Functions ---

def download_xml(url):
    """Downloads and parses an XML file from a URL."""
    try:
        response = requests.get(url, timeout=20) # Increased timeout for larger files
        response.raise_for_status()
        root = ET.fromstring(response.content)
        return root
    except requests.exceptions.RequestException as e:
        raise ConnectionError(f"Błąd podczas pobierania pliku z URL {url}: {e}")
    except ET.ParseError as e:
        raise ValueError(f"Błąd podczas parsowania XML z URL {url}: {e}")
    except Exception as e:
        raise Exception(f"Nieoczekiwany błąd podczas przetwarzania URL {url}: {e}")

def build_products_dict(root):
    """Builds a dictionary of products from the XML root for quick lookup."""
    products = {}
    if root is None:
        return products
    for offer in root.findall('o'):
        product_id = offer.get('id')
        if product_id:
            products[product_id] = offer
    return products

def get_attr(offer, attr_name, default=""):
    """Gets a specific attribute value from an offer's 'attrs' section."""
    attrs = offer.find('attrs')
    if attrs is not None:
        for a in attrs.findall('a'):
            if a.get('name') == attr_name:
                return (a.text or "").strip()
    return default

def get_category(offer):
    """Gets the category text from an offer."""
    cat = offer.find('cat')
    return (cat.text or "").strip() if cat is not None else ""

def get_name(offer):
    """Gets the product name from an offer."""
    name = offer.find('name')
    return (name.text or "").strip() if name is not None else ""

def get_desc(offer):
    """Gets the product description from an offer."""
    desc = offer.find('desc')
    return (desc.text or "").strip() if desc is not None else ""

def get_main_image(offer):
    """Gets the main image URL from an offer."""
    imgs = offer.find('imgs')
    if imgs is not None:
        main = imgs.find('main')
        if main is not None:
            return main.get('url', '')
    return ""

def get_extra_images(offer):
    """Gets up to 9 extra image URLs from an offer."""
    imgs = offer.find('imgs')
    urls = []
    if imgs is not None:
        all_i = imgs.findall('i')
        for i in all_i[:9]:  # max 9, as per original logic
            url = i.get('url', '')
            if url:
                urls.append(url)
    return ";".join(urls)

def short(text, length):
    """Truncates text to a specified length."""
    return text if len(text) <= length else text[:length]

def get_price(offer):
    """Gets the price from an offer."""
    return offer.get('price', "")

def get_stock(offer):
    """Gets the stock quantity from an offer."""
    return offer.get('stock', "")

# NEW FUNCTION: Get weight directly from the offer tag
def get_weight(offer):
    """Gets the weight attribute from the offer tag."""
    return offer.get('weight', "")

def strip_html_tags(text):
    """Removes HTML tags from a string."""
    if not isinstance(text, str):
        return text
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

def get_brand(offer):
    """Determines the brand from attributes or product ID."""
    producent = get_attr(offer, "Producent")
    if producent:
        return producent.strip()
    prod_id = offer.get('id', '')
    if '_' in prod_id:
        return prod_id.split('_')[0]
    return prod_id

def process_feeds(se_url, dk_url, fi_url, output_file_path, status_callback):
    """
    Main function to process the three XML feeds and generate the Excel file.
    Uses a callback to update the GUI status.
    """
    try:
        status_callback("Pobieranie i parsowanie feeda SE...")
        root_se = download_xml(se_url)
        status_callback("Pobieranie i parsowanie feeda DK...")
        root_dk = download_xml(dk_url)
        status_callback("Pobieranie i parsowanie feeda FI...")
        root_fi = download_xml(fi_url)

        status_callback("Budowanie słowników produktów...")
        se_offers = build_products_dict(root_se)
        dk_offers = build_products_dict(root_dk)
        fi_offers = build_products_dict(root_fi)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dane"

        # MODIFIED: Added "weight" to headers
        headers = [
            "sku", "weight", "brand", "gtin", "stock", "mainImage", "extraImages",
            "titleSe", "descriptionSe", "titleDk", "descriptionDk", "titleFi", "descriptionFi",
            "category",
            "originalPriceSe", "originalPriceDk", "originalPriceFi",
            "shippingCostSe", "shippingCostDk", "shippingCostFi",
            "deliveryTimeMinSe", "deliveryTimeMinDk", "deliveryTimeMinFi",
            "deliveryTimeMaxSe", "deliveryTimeMaxDk", "deliveryTimeMaxFi",
            "vatSe", "vatDk", "vatFi",
            "deliverySe", "deliveryDk", "deliveryFi"
        ]
        ws.append(headers)

        total = len(se_offers)
        count = 0
        status_callback(f"Przetwarzanie {total} produktów...")

        for pid, se_offer in se_offers.items():
            count += 1
            if count % 50 == 0: # Update status every 50 products
                status_callback(f"Przetwarzanie produktu {count}/{total}...")

            dk_offer = dk_offers.get(pid)
            fi_offer = fi_offers.get(pid)

            # Extract data from the primary (SE) offer
            sku = se_offer.get('id', '')
            # NEW: Get weight
            weight = get_weight(se_offer)
            brand = get_brand(se_offer)
            gtin = get_attr(se_offer, "EAN")
            stock = get_stock(se_offer)
            main_image = get_main_image(se_offer)
            extra_images = get_extra_images(se_offer)
            name = get_name(se_offer)
            desc = get_desc(se_offer)
            category = get_category(se_offer)
            originalPriceSe = get_price(se_offer)

            short_name = short(name, 135)
            
            # MODIFIED LOGIC: Strip HTML only if description is too long
            if len(desc) > 9500:
                processed_desc = strip_html_tags(desc)
            else:
                processed_desc = desc
            short_desc = short(processed_desc, 9500)

            # Get prices from other feeds, if available
            originalPriceDk = get_price(dk_offer) if dk_offer is not None else ""
            originalPriceFi = get_price(fi_offer) if fi_offer is not None else ""

            # MODIFIED: Added weight to the row
            row = [
                sku, weight, brand, gtin, stock, main_image, extra_images,
                short_name, short_desc, short_name, short_desc, short_name, short_desc,
                category,
                originalPriceSe, originalPriceDk, originalPriceFi,
                "0", "0", "0",
                4, 4, 4,
                6, 6, 6,
                25, 25, 25.5, # Corrected VAT for FI
                "HomeDelivery", "HomeDelivery", "HomeDelivery"
            ]
            ws.append(row)

        status_callback("Zapisywanie pliku Excel...")
        wb.save(output_file_path)
        messagebox.showinfo("Sukces", f"Plik '{os.path.basename(output_file_path)}' został pomyślnie zapisany.")

    except (ConnectionError, ValueError, Exception) as e:
        messagebox.showerror("Błąd", str(e))
    except Exception as e:
        messagebox.showerror("Nieoczekiwany błąd", f"Wystąpił nieoczekiwany błąd: {e}")
    finally:
        status_callback("") # Clear status message

# --- GUI Setup with CustomTkinter ---

class FeedProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Window Configuration ---
        self.title("Narzędzie do Przetwarzania Feedów XML")
        self.geometry("600x360") # Adjusted window size for new widget
        self.resizable(False, False)

        # --- Theme and Appearance ---
        ctk.set_appearance_mode("System")  # Can be "Dark", "Light"
        ctk.set_default_color_theme("blue") # "blue", "green", "dark-blue"

        # --- Main Frame ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        self.create_widgets()

    def create_widgets(self):
        """Creates and lays out all the widgets in the application."""
        # --- Prefix Input Field ---
        ctk.CTkLabel(self.main_frame, text="Przedrostek pliku XML:", font=("Helvetica", 12)).pack(anchor="w", padx=10, pady=(10,0))
        self.prefix_entry = ctk.CTkEntry(self.main_frame, width=500, placeholder_text="np. nazwa_dostawcy")
        self.prefix_entry.pack(padx=10, pady=(0,10), fill="x")

        # --- Output Path Field ---
        ctk.CTkLabel(self.main_frame, text="Folder docelowy:", font=("Helvetica", 12)).pack(anchor="w", padx=10, pady=(10,0))
        
        output_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        output_frame.pack(fill="x", padx=10, pady=(0,20))

        self.output_path_entry = ctk.CTkEntry(output_frame)
        self.output_path_entry.pack(side="left", fill="x", expand=True)

        # Set default path to the user's Desktop
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.output_path_entry.insert(0, desktop_path)

        browse_button = ctk.CTkButton(output_frame, text="Przeglądaj...", command=self.browse_output_directory, width=120)
        browse_button.pack(side="left", padx=(10, 0))

        # --- Action Button ---
        self.process_button = ctk.CTkButton(self.main_frame, text="Generuj plik Excel", command=self.run_processing, height=40, font=("Helvetica", 14, "bold"))
        self.process_button.pack(pady=10, padx=10, fill="x")

        # --- Status Label ---
        self.status_label = ctk.CTkLabel(self, text="", text_color="gray")
        self.status_label.pack(pady=(0, 10), padx=20)

    def browse_output_directory(self):
        """Opens a dialog to select an output directory."""
        directory = filedialog.askdirectory(title="Wybierz folder docelowy")
        if directory:
            self.output_path_entry.delete(0, "end")
            self.output_path_entry.insert(0, directory)

    def update_status(self, message):
        """Callback function to update the status label from the processing function."""
        self.status_label.configure(text=message)
        self.update_idletasks() # Refresh the GUI to show the new message

    def run_processing(self):
        """Handles the button click event for processing and saving the file."""
        prefix = self.prefix_entry.get().strip()
        output_dir = self.output_path_entry.get().strip()

        if not prefix:
            messagebox.showwarning("Brakujące dane", "Proszę wprowadzić przedrostek pliku.")
            return

        if not output_dir:
            messagebox.showwarning("Brakujące dane", "Proszę podać folder docelowy.")
            return
        
        if not os.path.isdir(output_dir):
            messagebox.showerror("Błąd", f"Podany folder docelowy nie istnieje:\n{output_dir}")
            return

        # Construct the full output path and check for overwrite
        filename = f"{prefix}_output_feeds.xlsx"
        output_file_path = os.path.join(output_dir, filename)

        if os.path.exists(output_file_path):
            if not messagebox.askyesno("Potwierdzenie", f"Plik '{filename}' już istnieje w wybranej lokalizacji.\n\nCzy chcesz go nadpisać?"):
                return  # User chose not to overwrite

        # Construct the full URLs from the prefix
        base_url = "https://sm-prods.com/feeds/"
        se_url = f"{base_url}{prefix}_cdon_se.xml"
        dk_url = f"{base_url}{prefix}_cdon_dk.xml"
        fi_url = f"{base_url}{prefix}_cdon_fi.xml"
        
        # Disable button during processing
        self.process_button.configure(state="disabled", text="Przetwarzanie...")
        self.update_idletasks() # Ensure GUI updates before long task
        
        # Run the main processing function
        process_feeds(se_url, dk_url, fi_url, output_file_path, self.update_status)

        # Re-enable button after processing is complete or an error occurs
        self.process_button.configure(state="normal", text="Generuj plik Excel")


if __name__ == "__main__":
    app = FeedProcessorApp()
    app.mainloop()
