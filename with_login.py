import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
import time
from tqdm import tqdm
import pandas as pd
import re


options = uc.ChromeOptions()
options.add_argument('--headless')  # Run Chrome in headless mode
driver = uc.Chrome(options=options)

def login(email, password):
    # Open the login page
    url = "https://rla.org/site/login"
    driver.get(url)
    time.sleep(3)

    # Locate the email/username input field and enter the email
    email_field = driver.find_element(By.ID, "loginform-username")
    email_field.send_keys(email)

    # Locate the password input field and enter the password
    password_field = driver.find_element(By.ID, "loginform-password")
    password_field.send_keys(password)

    # Locate the login button and click it
    login_button = driver.find_element(By.XPATH, "//input[@type='submit' and @value='Log In']")
    login_button.click()
    print("logged in succesfully!")
    # Add a wait time to see the result of the login action
    time.sleep(5)


def write_excel(items, path):
    if len(items) == 0:
        return

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(items)

    # Write DataFrame to Excel
    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

# Function to get text between headings
def get_text_between_headings(container, start_heading, end_heading=None):
    texts = []
    start_element = container.find_element(By.XPATH, f".//h4[text()='{start_heading}']")
    siblings = start_element.find_elements(By.XPATH, './following-sibling::*')

    for sibling in siblings:
        if end_heading and sibling.tag_name == 'h4' and sibling.text.strip() == end_heading:
            break
        texts.append(sibling.text.strip())

    return " ".join(texts).strip()


def scrape_profile(name, link):
    driver.get(link)
    time.sleep(0.3)
    segments = driver.find_element(By.ID, "segments")
    # Find all elements with 'bis_skin_checked="1"'
    segment_elements = segments.find_elements(By.TAG_NAME, "div")

    returns=""
    repair=""
    resell=""
    recycle=""
    resources=""
    location=""
    website = ""
    # Initialize variables to store the results
    company_overview = ""
    products = ""
    certifications = ""
    specialties = ""
    where_work_is_performed = ""
    established=""
    employees=""
    locations=""
    service_areas=""
    company_type=""
    # Iterate through the elements and extract the required information
    for element in segment_elements:
        try:
            # Get the segment name from the <span class="segment"> element
            segment_name = element.find_element(By.CLASS_NAME, "segment").text.strip()
            ul_text = element.find_element(By.TAG_NAME, "ul").text.strip().replace("\n", ", ")

            if segment_name == "Returns":
                returns = ul_text
            elif segment_name == "Repair":
                repair = ul_text
            elif segment_name == "Resell":
                resell = ul_text
            elif segment_name == "Recycle":
                recycle = ul_text
            elif segment_name == "Resources":
                resources = ul_text
        except Exception as e:
            print("Problem on segments: ", e)
            continue
    # Print the results
    #print("Returns:", returns)
    #print("Repair:", repair)
    #print("Resell:", resell)
    #print("Recycle:", recycle)
    #print("Resources:", resources)
    try:
        profile_details = driver.find_element(By.ID, "profileDetails")
    except Exception as e:
        print("No profile details")
        profile_details = ""
    if profile_details:
        #find locations and website url
        try: 
            location = profile_details.find_element(By.TAG_NAME, "p").text.strip().replace("\n", ", ")
        except Exception as e:
            location = ""
        try: 
            website = profile_details.find_element(By.TAG_NAME, "a")
            website = website.get_attribute("href")
        except Exception as e:
            website = ""
        #print("Location:", location)
        #print("Website:", website)

        try:
            # Get the HTML content of the profile details
            profile_details_html = profile_details.get_attribute('innerHTML')

            # Use regex to extract the text between the headings
            company_overview_match = re.search(r'<h4>Company Overview<\/h4>(.*?)<h4 class="mt-4">Products<\/h4>', profile_details_html, re.DOTALL)

            # Extract the company overview text
            if company_overview_match:
                company_overview = company_overview_match.group(1).strip()
                # Remove any HTML tags that may be present within the extracted text
                company_overview = re.sub(r'<.*?>', '', company_overview).strip()
            else:
                company_overview = ""
        except Exception as e:
            print("Problem on extracting company_overview: ", e)
        try:
            # Extract the "Products" text
            products = get_text_between_headings(profile_details, "Products", "Certifications:")
            products = products.replace('\n', ', ')
        except Exception as e:
            print("Problem on extracting products: ", e)
        try:
            # Extract the "Certifications:" text
            certifications = get_text_between_headings(profile_details, "Certifications:", "Specialties:")
            certifications = certifications.replace('\n', ', ')
        except Exception as e:
            print("Problem on extracting certifications: ", e)
        try:
            # Extract the "Specialties:" text
            specialties = get_text_between_headings(profile_details, "Specialties:", "Where Work Is Performed:")
        except Exception as e:
            print("Problem on extracting specialties: ", e)
        try:
            # Extract the "Where Work Is Performed:" text
            where_work_is_performed = get_text_between_headings(profile_details, "Where Work Is Performed:")
        except Exception as e:
            print("Problem on extracting where_work_is_performed: ", e)

        # Print the results
        #print("Company Overview:", company_overview)
        #print("Products:", products)
        #print("Certifications:", certifications)
        #print("Specialties:", specialties)
        #print("Where Work Is Performed:", where_work_is_performed)

        # Locate the container with ID 'additionalInfo'
        try:
            additional_info = driver.find_element(By.ID, "additionalInfo")
        except Exception as e:
            additional_info = ""
        if additional_info:
            # Get the HTML content of the additional info section
            additional_info_html = additional_info.get_attribute('innerHTML')

            # Extract established year
            established_match = re.search(r'<h4>Established<\/h4>\s*<p>(.*?)<\/p>', additional_info_html)
            established = established_match.group(1).strip() if established_match else ""

            # Extract employees
            employees_match = re.search(r'<h4>Employees<\/h4>\s*<p>(.*?)<\/p>', additional_info_html)
            employees = employees_match.group(1).strip() if employees_match else ""

            # Extract locations
            locations_match = re.search(r'<h4>Locations<\/h4>\s*<p>(.*?)<\/p>', additional_info_html)
            locations = locations_match.group(1).strip() if locations_match else ""

            try:
                service_areas = additional_info.find_element(By.CSS_SELECTOR, "div.col-sm-2.serviceArea").text.strip().replace("Service Area(s)", "").replace('\n', ', ')
            except Exception as e:
                service_areas = ""

            # Extract company type
            company_type_match = re.search(r'<h4>Company Type<\/h4>\s*<p.*?>(.*?)<\/p>', additional_info_html)
            company_type = company_type_match.group(1).strip() if company_type_match else ""

        # Print the results
        #print("Established:", established)
        #print("Employees:", employees)
        #print("Locations:", locations)
        #print("Service Areas:", service_areas)
        #print("Company Type:", company_type)
        #print("-------------------------------------------------------------------")


        #Return a Dictionary
    return{
        'Company Name': name,
        'Location': location,
        'Website': website,
        'Returns': returns,
        'Repair': repair,
        'Resell': resell,
        'Recycle': recycle,
        'Resources': resources,
        'Company Overview': company_overview,
        'Products': products,
        'Certifications': certifications,
        'Specialties': specialties,
        'Where Work Is Performed': where_work_is_performed,
        'Established': established,
        'Employees': employees,
        'Locations': locations,
        'Service Areas': service_areas,
        'Company Type': company_type,
    }






def main_function(link_ad):
    print("Starting to scrape the companies...")
    print("")
    results = []
    if link_ad:
        driver.get(link_ad)
        time.sleep(1.2)
        # Define the JavaScript code to scroll the page
        scroll_script = """
            window.scrollBy(0, window.innerHeight);
        """

        # Scroll down the page in middle speed
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 2);")
        time.sleep(0.9)

        # Scroll to the bottom of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")

        time.sleep(1.4)
        # Find all the td elements with class 'coName'
        co_name_elements = driver.find_elements(By.CSS_SELECTOR, "td.coName")

        # List to hold the company names
        companies = []

        base_url = "https://rla.org"

        # Iterate through each element to get the company name and profile URL
        for element in co_name_elements:
            # Get the company name
            div_element = element.find_element(By.TAG_NAME, "div")
            company_name = div_element.text
            
            # Get the URL from the button's onclick attribute
            button_element = element.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.btn-sm.mt-1")
            onclick_value = button_element.get_attribute("onclick")
            profile_url = base_url + onclick_value.split("'")[1]  # Extract URL from onclick value
            profile_url = profile_url.replace("../..", "")
            # Add the company name and URL as a tuple to the list
            companies.append((company_name, profile_url))

        for company, url in tqdm(companies):
            data = scrape_profile(company, url)
            results.append(data)
            print(data)
            time.sleep(0.6)
    print("Scraping companies finished succesfully!")
    print("")
    return results



def scrape_memeber_data(member):
    try:
        speaker_name_element = member.find_element(By.CSS_SELECTOR, "div.speaker-name a").text.strip()
    except Exception as e:
        speaker_name_element = ""
    try:
        com_role = member.find_element(By.CSS_SELECTOR, "div.committee-role").text.strip()
    except Exception as e:
        com_role = ""
    try:
        sp_role = member.find_element(By.CSS_SELECTOR, "div.speaker-role").text.strip()
    except Exception as e:
        sp_role = ""
    try:
        sp_company = member.find_element(By.CSS_SELECTOR, "div.speaker-company").text.strip()
    except Exception as e:
        sp_company = ""
    
    return{
        'Speaker Name': speaker_name_element,
        'Committee Role': com_role,
        'Speaker Role': sp_role,
        'Speaker Company': sp_company,
    }

def scrape_members(link_ad):
    print("")
    print("Starting the scraping of members...")
    print("")
    results2 = []
    if link_ad:
        driver.get(link_ad)
        time.sleep(0.7)
        try:
            # Locate the element with the ID 'committees'
            committees_element = driver.find_element(By.ID, "committees")

            # Find all the <a> tags within the committees element
            a_tags = committees_element.find_elements(By.TAG_NAME, "a")
        except Exception as e:
            a_tags = ""

        # List to hold the committee links
        committee_links = []
        # Iterate through each <a> tag to get the href attribute
        for a in a_tags:
            link = a.get_attribute("href")
            committee_links.append(link)
        # Print the committee links
        for link in tqdm(committee_links):
            if "https://rla.org/committee/" in link:
                driver.get(link)
                time.sleep(1.2)
                # Define the JavaScript code to scroll the page
                scroll_script = """
                    window.scrollBy(0, window.innerHeight);
                """

                # Scroll down the page in middle speed
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 2);")
                time.sleep(0.9)

                # Scroll to the bottom of the page
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")

                time.sleep(1.4)
                #find leaders
                leaders = driver.find_elements(By.CSS_SELECTOR, "div.leader")
                for lead in leaders:
                    data2 = scrape_memeber_data(lead)
                    results2.append(data2)
                    print(data2)
                
                memebers = driver.find_elements(By.CSS_SELECTOR, "div.member")
                for mem in memebers:
                    data2 = scrape_memeber_data(mem)
                    results2.append(data2)
                    print(data2)
    print("")
    print("Scraping members finished succesfully!")
    return results2



email = "xxx"
password = "xxxxx"
login(email, password)

URL_COMPANIES = "https://rla.org/directory/company-tag/list"
output_csv_path_companies = "results_companies2.xlsx"

URL_MEMBERS = "https://rla.org/committee/splash"
output_csv_path_members = "results_members2.xlsx"


scraped_data = main_function(URL_COMPANIES)
scraped_data2 = scrape_members(URL_MEMBERS)

# Close the webdriver
driver.quit()

# Write the results to a CSV file
write_excel(scraped_data, output_csv_path_companies)
# Write the results to a CSV file
write_excel(scraped_data2, output_csv_path_members)
