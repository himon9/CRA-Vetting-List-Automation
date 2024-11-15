from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
import time


# File paths
input_file = "client_names.txt"  # Text file containing client names
output_file = "clients_data.xlsx"  # Excel file to store results

# Initialize Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Client Data"
sheet.append(
    [
        "Account Name",
        "InOV",
        "Industry",
        "TAM Account",
        "Owner",
        "Primary BDE",
        "Services Provided",
        "Won Pipeline",
        "Open Pipeline",
        "Last Communication",
    ]
)

# Read client names from text file
with open(input_file, "r") as file:
    client_names = [line.strip() for line in file.readlines()]


chrome_user_data_path = "C:/Users/us83263/AppData/Local/Google/Chrome/User Data"
profile_name = "Profile 1"  # Using Himon's profile folder name

# Set up Chrome WebDriver with Himon's profile
chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={chrome_user_data_path}")
chrome_options.add_argument(f"--profile-directory={profile_name}")
chrome_options.add_argument("--start-maximized")


driver = webdriver.Chrome(options=chrome_options)

# Ask user for the URL of the website


for client_name in client_names:
    url = "https://gtusoneview.crm.dynamics.com/"
    driver.get(url)
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        print("gtusoneview.crm.dynamics.com is fully loaded.\n\n")

        print(f"Processing client: {client_name}")

        search_box = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "GlobalSearchBox"))
        )
        # Enter the client name
        search_box.send_keys(client_name)
        search_box.send_keys(Keys.RETURN)

        # Wait for the search body to be fully loaded
        InOV = False
        try:
            button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//div[@role="row" and @row-index="0"]//div[@aria-colindex="2"]//button',
                    )
                )
            )
            InOV = True  # If client result is found, set present to True
            print(f"InOV: {InOV}")
            button.click()

            # In clinet info home page
            header_controls_list = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "headerControlsList_2"))
            )
            # print(header_controls_list)

            # Find all child elements with the specific attributes
            column_elements = header_controls_list.find_elements(
                By.XPATH,
                './/div[@role="presentation" and @data-preview_orientation="column"]',
            )

            # Initialize a dictionary to store extracted data
            data_dict = {}
            # Iterate and extract text
            for index, element in enumerate(column_elements):
                text = element.text.strip()  # Get text and remove extra spaces
                # Split the text by '\n' and take only the first part (the value)
                if "\n" in text:
                    nameValue = text.split("\n")[
                        0
                    ].strip()  # Extract the first part (name)
                    # Extract the second part (label or value)
                    label_value = text.split("\n")[1].strip()

                    # Store the data in the dictionary
                    data_dict[label_value] = nameValue

                print(f"{label_value}: {nameValue}")

                # Wait for the target element with data-id "gt_servicesprovided" to load
            container = driver.find_element(
                By.CSS_SELECTOR, '[data-id="gt_servicesprovided"]'
            )

            # Locate the element with the class "msos-viewmode-text" inside the container
            msos_text_element = container.find_element(
                By.CLASS_NAME, "msos-viewmode-text"
            )

            # Get the text content
            services_provided = msos_text_element.text
            print(f"Services provide by GT: {services_provided}")

            # Locate the input element by its data-id attribute
            input_element = driver.find_element(
                By.CSS_SELECTOR,
                '[data-id="gt_openpipeline.fieldControl-currency-text-input"]',
            )

            # GETTING THE VALUE OF OPEN AND WON PIPELINE
            # Get the value of the 'value' attribute
            input_value = input_element.get_attribute("value")
            open_pipeline = input_value
            print(f"Open Pipeline: {open_pipeline}")

            # Locate the input element by its data-id attribute
            input_element = driver.find_element(
                By.CSS_SELECTOR,
                '[data-id="gt_wonpipelinecfy.fieldControl-currency-text-input"]',
            )

            # Get the value of the 'value' attribute
            input_value = input_element.get_attribute("value")
            won_pipeline = input_value
            print(f"Won Pipeline: {won_pipeline}")

            # RECENT COMMENT DATE & TIME
            # Wait until the parent div with the specific data-lp-id is present
            parent_div = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        'div[data-lp-id="MscrmControls.TimelineWallControl.TimelineWall|notescontrol|account-record-0"]',
                    )
                )
            )

            # Wait for the child div with a title starting with "Sort date" to be present inside the parent div
            child_div = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, './/div[starts-with(@title, "Sort date")]')
                )
            )

            # Get the text content of the child div
            last_communication = child_div.text.replace("Sort date: ", "").strip()
            print("Last Communication Date & Time: ", last_communication)

            # For locating the INDUSTRY TYPE
            li_element = driver.find_element(
                By.CSS_SELECTOR, 'li[data-id="tablist-tab_21"]'
            )

            # Scroll to the element if necessary (optional)
            actions = ActionChains(driver)
            actions.move_to_element(li_element).perform()

            # Click the <li> element
            li_element.click()

            wait = WebDriverWait(driver, 10)
            div_element = wait.until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        'div[data-id="gt_industrygrouping.fieldControl-LookupResultsDropdown_gt_industrygrouping_selected_tag_text"]',
                    )
                )
            )

            # Get the text inside the <div>
            industry = div_element.text

            print("Industry type:", industry)

            # Append data to Excel
            sheet.append(
                [
                    client_name,
                    InOV, 
                    industry,
                    data_dict.get("TAM Status", ""),  # Get TAM Account
                    data_dict.get("Owner", ""),  # Get Owner
                    data_dict.get("Primary CRE", ""),  # Get Primary CRE
                    services_provided,
                    won_pipeline,
                    open_pipeline,
                    last_communication,
                ]
            )

        except TimeoutException:
            print("Client details not found in OV")

        # Now viewing the client info page

    except Exception as e:
        print(f"gtusoneview.crm.dynamics.com failed to load")

    # print("End")

# Save Excel file
workbook.save(output_file)
print(f"Data successfully exported to {output_file}")

# Quit driver
driver.quit()
