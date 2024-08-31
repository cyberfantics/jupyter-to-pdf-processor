'''
Common Libraries Used In Both
'''
import os
import time
import shutil
import pyfiglet
from colorama import init, Fore, Back, Style
import platform


'''
Libraries Which I Used For Jupyter To PDF
'''
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


'''
Libraries Which I Used To Process PDF
'''
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor
from io import BytesIO

'''
Create Banner and Add Clear Function
'''
init()

def clear():
    """Clear the console screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def first_banner():
    """Display the initial banner."""
    time.sleep(2)
    clear()
    salfi = pyfiglet.Figlet(font="starwars")
    banner = salfi.renderText("Cyber Fantics")
    print(f'{Fore.YELLOW}{banner}{Style.RESET_ALL}')
    print(f"{Fore.CYAN}[+] {Fore.GREEN}Welcome to Cyber Fantics PDF Processor!{Style.RESET_ALL}")
    time.sleep(2)  # Pause to let the banner be read


'''
Converting Jupyter NoteBook Into PDF
'''
def extract_pdf_name(jupyter_file_path):
    """Extract the PDF file name from the Jupyter notebook path."""
    file_name = os.path.splitext(os.path.basename(jupyter_file_path))[0]
    pdf_file_name = f"{file_name}.pdf"
    return pdf_file_name

def get_download_folder():
    """Detect the user's download folder based on the operating system."""
    if platform.system() == 'Windows':
        return os.path.join(os.path.expanduser("~"), "Downloads")
    elif platform.system() in ['Darwin', 'Linux']:
        return os.path.join(os.path.expanduser("~"), "Downloads")
    else:
        raise EnvironmentError("Unsupported operating system for automatic download folder detection.")

def download_jupyter(jupyter_file_path, pdf_file_name, count, browser_options):
    """Download and convert Jupyter notebook to PDF."""
    os.makedirs('pdf', exist_ok=True)

    destination_folder = os.path.dirname(jupyter_file_path)
    first_banner()
    print(f'\t{Fore.RED}[+] {Fore.GREEN}Processing PDF {Fore.CYAN}{count}{Style.RESET_ALL}')
    
    options = Options()
    for option in browser_options:
        options.add_argument(option)
    
    # Set environment variable to suppress USB error messages
    os.environ['WDM_LOG_LEVEL'] = '0'
    driver = webdriver.Chrome(options=options)

    try:
        driver.get('https://www.convert.ploomber.io/')
        driver.maximize_window()
        
        print(f'\t{Fore.RED}[+] {Fore.GREEN}Opened Chrome Browser{Style.RESET_ALL}')

        file_input = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[1]/div[1]/div[1]/div[1]/div/input'))
        )
        
        print(f'\t{Fore.RED}[+] {Fore.GREEN}Uploading File For Conversion{Style.RESET_ALL}')
        file_input.send_keys(jupyter_file_path)

        convert_button = WebDriverWait(driver, 40).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div[2]/div[1]/div[1]/div[1]'))
        )

        print(f'\t{Fore.RED}[+] {Fore.GREEN}Clicking Convert Button To Create PDF {Fore.CYAN}{pdf_file_name}{Style.RESET_ALL}')
        convert_button.click()

        download_folder = get_download_folder()
        pdf_file_path = os.path.join(download_folder, pdf_file_name)
        timeout = time.time() + 60*5  # 5 minutes from now
        
        while not os.path.isfile(pdf_file_path):
            print(f'\t{Fore.RED}[+] {Fore.GREEN}Downloading PDF {Fore.CYAN}{count}{Style.RESET_ALL}')

            if time.time() > timeout:
                print(f"\n\t{Fore.MAGENTA} [-] {Fore.RED}Error: {Fore.BLUE}PDF file not found in the download folder.{Style.RESET_ALL}")
                break
            time.sleep(30)  # Check every 30 seconds

        if os.path.isfile(pdf_file_path):
            new_location = os.path.join(destination_folder, pdf_file_name)
            shutil.move(pdf_file_path, new_location)
            print(f"\n\t{Fore.GREEN}[-] {Fore.BLUE}Conversion successful.{Style.RESET_ALL}")

    finally:
        driver.quit()

def notebook_main(browser_options):
    """Main function for processing Jupyter notebooks."""
    folder_path = os.path.abspath('.')  # Use the current directory or specify another folder
    count = 1

    for file_name in os.listdir(folder_path):
        if file_name.endswith('.ipynb'):
            jupyter_file_path = os.path.join(folder_path, file_name)
            pdf_file_name = extract_pdf_name(jupyter_file_path)
            download_jupyter(jupyter_file_path, pdf_file_name, count, browser_options)
            count += 1

'''
Process PDF Files
'''
def create_overlay_pdf(page_number, pagesize, header_text, footer_text, border_color):
    """Create an overlay PDF with borders and page numbers."""
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=pagesize)
    
    border_margin = 30
    width, height = pagesize
    can.setStrokeColor(border_color)  # Set border color
    can.setLineWidth(3)
    can.roundRect(border_margin, border_margin, width - 2 * border_margin, height - 2 * border_margin, radius=10, stroke=1, fill=0)
    
    can.setFont("Helvetica", 8)
    can.setFillColor(border_color)  # Set text color
    can.drawString(width - 71, 15, f'Page {str(page_number)}')
    
    can.setFont("Helvetica", 8)
    can.setFillColor(border_color)  # Set text color
    can.drawString(50, 15, footer_text)
    can.drawString(width - 89, height - 18, header_text)  # Custom header
    
    can.save()
    packet.seek(0)
    
    new_pdf = PdfReader(packet)
    return new_pdf.pages[0]

def add_borders_and_numbers(input_pdf_path, output_pdf_path, header_text, footer_text, count, start_page_number=None, border_color=HexColor('#095d40')):
    """Add borders and page numbers to a PDF."""
    first_banner()
    print(f'\n\t{Fore.MAGENTA}[+] {Fore.CYAN}Processing PDF {Fore.CYAN}{count}...{Style.RESET_ALL}')
    pdf_writer = PdfWriter()
    pdf_reader = PdfReader(input_pdf_path)
    
    for page_number, page in enumerate(pdf_reader.pages, start=start_page_number or 1):
        if start_page_number is not None:
            overlay = create_overlay_pdf(page_number, letter, header_text, footer_text, border_color)
            page.merge_page(overlay)
        pdf_writer.add_page(page)
    
    with open(output_pdf_path, 'wb') as f:
        pdf_writer.write(f)
    print(f'\t{Fore.MAGENTA}[+] {Fore.GREEN}PDF with borders and page numbers saved{Style.RESET_ALL}')
    time.sleep(2)  # Add a delay of 2 seconds

def process_pdfs(downloads_folder, data_pdf_path, ownername, count, border_color):
    """Handle the processing of PDF files."""
    header_text = ownername
    footer_text = os.path.splitext(os.path.basename(data_pdf_path))[0]

    pdf_dir = os.path.join(downloads_folder, 'pdf')
    if not os.path.exists(pdf_dir):
        os.mkdir(pdf_dir)

    processed_data_pdf_path = os.path.join(pdf_dir, f"{footer_text}.pdf")

    add_borders_and_numbers(data_pdf_path, processed_data_pdf_path, header_text, footer_text, count, start_page_number=1, border_color=border_color)

    try:
        os.remove(data_pdf_path)
        print(f'\n\t{Fore.RED}Deleted UnProcessed PDF: {Fore.CYAN}{os.path.basename(data_pdf_path)}{Style.RESET_ALL}')
    except FileNotFoundError:
        print(f'\n\t{Fore.RED}File not found: {Fore.CYAN}{data_pdf_path}{Style.RESET_ALL}')

    time.sleep(5)

def pdf_main(border_color):
    """Main function for processing existing PDF files.""" 
    download_folder = os.path.abspath('.')  # Use the current directory or specify another folder
    count = 1
    
    for file_name in os.listdir(download_folder):
        if file_name.endswith('.pdf'):
            full_pdf_path = os.path.join(download_folder, file_name)
            first_banner()
            ownername = input(f"\t{Fore.GREEN}[-] {Fore.BLUE}Enter the owner's name for the PDF '{Fore.CYAN}{file_name}{Fore.GREEN}': {Style.RESET_ALL}")
            process_pdfs(download_folder, full_pdf_path, ownername, count, border_color)
            count += 1
            

'''
Main To Control Operations
'''
def main():
    """Main function to control the program operations."""
    first_banner()
    
    print(f"{Fore.CYAN}[+] {Fore.GREEN}Please choose an option:{Style.RESET_ALL}")
    print(f"\n\t{Fore.YELLOW}[1]{Fore.BLUE} Convert Jupyter Notebooks to PDFs{Style.RESET_ALL}")
    print(f"\t{Fore.YELLOW}[2]{Fore.BLUE} Process Existing PDF Files{Style.RESET_ALL}")
    print(f"\t{Fore.YELLOW}[3]{Fore.BLUE} Convert Notebooks and Process PDFs{Style.RESET_ALL}")
    
    choice = input(f"\n\t{Fore.GREEN}Enter your choice {Fore.CYAN}(1/2/3): {Style.RESET_ALL}")
    
    first_banner()
    border_color_input = input(f"\n\t{Fore.GREEN}[+] {Fore.CYAN}Enter border color in HEX (e.g., #095d40): {Style.RESET_ALL}")
    try:
        border_color = HexColor(border_color_input)
    except ValueError:
        print(f"\t{Fore.RED}[!] {Fore.YELLOW}Invalid HEX color format. Using default color.{Style.RESET_ALL}")
        border_color = HexColor('#095d40')

    if choice == '1':
        browser_options = []

        if input(f"\t{Fore.GREEN}[+] {Fore.CYAN}Do you want to run the browser in headless mode? (yes/no): {Style.RESET_ALL}").lower()[0] == 'y':
            browser_options.append('--headless')
        if input(f"\t{Fore.GREEN}[+] {Fore.CYAN}Do you want to suppress browser logs? (yes/no): {Style.RESET_ALL}").lower()[0] == 'y':
            browser_options.append('--log-level=3')
            browser_options.append('--disable-logging')

        print(f"\t{Fore.GREEN}[+] {Fore.CYAN}Converting Jupyter Notebooks to PDFs...{Style.RESET_ALL}")
        notebook_main(browser_options)
    elif choice == '2':
        pdf_main(border_color)
    elif choice == '3':

        browser_options = []
        if input(f"\t{Fore.GREEN}[+] {Fore.CYAN}Do you want to run the browser in headless mode? (yes/no): {Style.RESET_ALL}").lower()[0] == 'y':
            browser_options.append('--headless')
        if input(f"\t{Fore.GREEN}[+] {Fore.CYAN}Do you want to suppress browser logs? (yes/no): {Style.RESET_ALL}").lower()[0] == 'y':
            browser_options.append('--log-level=3')
            browser_options.append('--disable-logging')
        print(f"\t{Fore.GREEN}[+] {Fore.CYAN}Converting Jupyter Notebooks to PDFs...{Style.RESET_ALL}")

        notebook_main(browser_options)
        pdf_main(border_color)
    else:
        print(f"\t{Fore.RED}[-] {Fore.YELLOW}Invalid choice. Exiting...{Style.RESET_ALL}")

    print(f"\t{Fore.GREEN}[+] {Fore.CYAN}Operation completed!{Style.RESET_ALL}")

if __name__ == "__main__":
    main()
