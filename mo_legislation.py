import re

import pandas
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import ElementClickInterceptedException, TimeoutException, NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


class MOLegislation:
    def __init__(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.page_load_strategy = 'eager'
        self._driver = webdriver.Chrome(options=chrome_options)
        self._wait = WebDriverWait(self._driver, timeout=2)
        self._house_legislators = {}
        self._senate_legislators = {}
        self._senate_id_mapping = {}
        self._house_committees = {}
        self._senate_committees = []

    def _get_house_members(self):
        r = requests.get("https://documents.house.mo.gov/xml/261-MemberList.XML")
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml-xml")
        representatives = soup.find_all(name="RepresentativeXMLLink")
        for representative in representatives:
            self._get_representative(representative.string)

    def _get_representative(self, url):
        r = requests.get(url)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml-xml")
        district = soup.find(name="DistrictNum").string
        first_name = soup.find(name="FirstName")
        last_name = soup.find(name="LastName")
        if first_name and last_name:
            first_name = first_name.string
            last_name = last_name.string
            party = soup.find(name="Party").string
            address = soup.find(name="CapitolAddress").string
            address_matches = re.match(r'Room (\S+),', address)
            if address_matches:
                room = address_matches.group(1)
            else:
                room = 'None'
            home_town = soup.find(name="Hometown").string
            year_entered = soup.find(name="YearElected").string
            phone = soup.find(name="PhoneNumber").string
            email = soup.find(name="EmailAddress").string
            self._house_legislators[district] = {'first_name': first_name, 'last_name': last_name, 'district': district,
                'party': party, 'room': room, 'home_town': home_town, 'year_entered': year_entered, 'phone': phone,
                'email': email}

    def _get_house_committees(self):
        r = requests.get("https://documents.house.mo.gov/xml/261-CommitteeList.XML")
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml-xml")
        committees = soup.find_all(name="Committee")
        for committee in committees:
            committee_name = committee.find(name="Name").string
            self._house_committees[committee_name] = []
            members = committee.find_all(name="CommitteeMember")
            for member in members:
                chamber = member.find("Chamber").string.upper()
                if chamber == 'H':
                    district = member.find('MemberDistrict').string
                    leadership_position = member.find('PositionName').string
                    if leadership_position == 'Member':
                        leadership_position = ''
                    try:
                        legislator_info = self._house_legislators[district]
                        self._house_committees[committee_name].append(
                            {'first_name': legislator_info['first_name'], 'last_name': legislator_info['last_name'],
                                'district': district, 'leadership_position': leadership_position,
                                'party': legislator_info['party'], 'room': legislator_info['room'],
                                'home_town': legislator_info['home_town'],
                                'year_entered': legislator_info['year_entered'], 'phone': legislator_info['phone'],
                                'email': legislator_info['email']})
                    except KeyError:
                        print(f"No legislator found for district {district}")

    def create_house_roster(self):
        self._get_house_members()
        self._get_house_committees()
        house_roster = pandas.DataFrame.from_dict(self._house_legislators, orient='index')
        house_committees = pandas.DataFrame.from_dict(self._house_committees, orient='index')
        with pandas.ExcelWriter("Legislator_Roster_2026.xlsx") as writer:
            house_roster.to_excel(writer, sheet_name="MO House", index=False)
            house_committees.to_excel(writer, sheet_name="House Committees")

    def _get_senate_members(self):
        self._driver.get("https://www.senate.mo.gov/senators/index")
        links = self._driver.find_elements(By.CSS_SELECTOR, ".senators-grid > a")
        for link in links:
            self._get_senator(link)

    def _get_senator(self, link: WebElement):
        name_district = link.find_element(By.CLASS_NAME, 'card__footer').text.split("\n")
        name = name_district[0].replace('Senator ', '').split(' ')
        district = name_district[1].split(' ')[1:]
        district = ' '.join(district)
        self._senate_legislators[district] = {'first_name': ' '.join(name[:-1]), 'last_name': ' '.join(name[-1:]),
            'district': int(district), 'party': None, 'room': None, 'home_town': None, 'year_entered': None,
            'phone': None, 'email': None}
        match = re.search('id=(\\d)+', link.get_attribute('href'))
        if match:
            self._senate_id_mapping[match.group(1)] = district
        ActionChains(self._driver).move_to_element(link).perform()
        self._wait.until(EC.element_to_be_clickable(link))
        try:
            link.click()
        except ElementClickInterceptedException:
            self._driver.execute_script("arguments[0].scrollIntoView({block: 'start', behavior: 'instant'});", link)
            self._wait.until(EC.element_to_be_clickable(link))
            link.click()
        details = self._driver.find_elements(By.CLASS_NAME, "detail-grid__value")
        if len(details) >= 4:
            self._senate_legislators[district]['party'] = details[0].text.strip()
            self._senate_legislators[district]['year_entered'] = details[2].text.strip()
        contact = self._driver.find_element(By.CSS_SELECTOR, ".sidebar-order .card:first-child .card__body").text
        room_match = re.search('Rm\\. (\\d+)', contact)
        phone = re.search('\\d{3}-\\d{3}-\\d{4}', contact)
        if room_match:
            self._senate_legislators[district]['room'] = room_match.group(1)
        if phone:
            self._senate_legislators[district]['phone'] = phone.group(0)
        email_link = self._driver.find_element(By.CSS_SELECTOR, '[title="Email the Senator"]')
        self._wait.until(EC.element_to_be_clickable(email_link))
        email_link.click()
        try:
            self._wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '[title="Continue to the contact page"]')))
        except TimeoutException:
            email_link.click()
            self._wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, '[title="Continue to the contact page"]')))
        continue_btn = self._driver.find_element(By.CSS_SELECTOR, '[title="Continue to the contact page"]')
        self._wait.until(EC.element_to_be_clickable(continue_btn))
        continue_btn.click()
        try:
            self._wait.until(EC.visibility_of_element_located((By.ID, "Contact")))
        except TimeoutException:
            if self._driver.current_url == 'https://www.senate.mo.gov/SenateHome':
                self._driver.back()
                self._wait.until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, '[title="Continue to the contact page"]')))
                continue_btn = self._driver.find_element(By.CSS_SELECTOR, '[title="Continue to the contact page"]')
                self._wait.until(EC.element_to_be_clickable(continue_btn))
                continue_btn.click()
        self._senate_legislators[district]['email'] = self._driver.current_url
        self._driver.back()
        self._wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '[title="Continue to the contact page"]')))
        self._driver.back()
        self._wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'detail-grid')))
        self._driver.back()
        self._wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "senators-grid")))

    def _get_senate_committees(self):
        self._driver.get("https://www.senate.mo.gov/committees/index")
        comm_list = self._driver.find_element(By.CLASS_NAME, "link-card-stack")
        links = [link.get_attribute("href") for link in comm_list.find_elements(By.TAG_NAME, "a")]
        links = [link for link in links if not link.lower().endswith('pdf')][:1]
        for link in links:
            self._get_senate_committees_by_type(link)

    def _get_senate_committees_by_type(self, url):
        self._driver.get(url)
        committee_btns = self._driver.find_elements(By.CSS_SELECTOR, '.content-grid .card.collapsible')
        committee_links = []
        for committee_btn in committee_btns:
            header = committee_btn.find_element(By.CLASS_NAME, 'card__header')
            try:
                header.click()
            except ElementClickInterceptedException:
                self._driver.execute_script("arguments[0].scrollIntoView({block: 'start', behavior: 'instant'});",
                                            header)
                self._wait.until(EC.element_to_be_clickable(header))
                header.click()
            self._driver.implicitly_wait(0.4)
            try:
                committee_link = committee_btn.find_element(By.LINK_TEXT, 'Committee Information')
            except NoSuchElementException:
                header.click()
                self._driver.implicitly_wait(0.4)
                committee_link = committee_btn.find_element(By.LINK_TEXT, 'Committee Information')
            self._wait.until(EC.element_to_be_clickable(committee_link))
            if committee_link and not committee_link.get_attribute("href").lower().endswith("pdf"):
                committee_links.append(committee_link.get_attribute('href'))
        for committee in committee_links:
            self._get_committee_page(committee)

    def _get_committee_page(self, url):
        self._driver.get(url)
        members = self._driver.find_elements(By.CSS_SELECTOR, ".senators-grid .card")
        committee_name = self._driver.find_element(By.CSS_SELECTOR, ".main-order > .card > .card__header").text
        for member in members:
            link = member.find_element(By.XPATH, '..')
            if link.tag_name.lower() == 'a':
                match = re.search('id=(\\d)+', link.get_attribute('href'))
                if match:
                    district = self._senate_id_mapping[match.group(1)]
                    senator_name = member.find_element(By.CLASS_NAME, "card__footer").text.split("\n")
                    try:
                        legislator_info = self._senate_legislators[district]
                        self._senate_committees.append({'committee': committee_name, 'district': str(district),
                            'first_name': legislator_info['first_name'], 'last_name': legislator_info['last_name'],
                            'leadership_position': senator_name[1] if len(senator_name) == 2 else '',
                            'party': legislator_info['party'], 'room': legislator_info['room'],
                            'home_town': legislator_info['home_town'], 'year_entered': legislator_info['year_entered'],
                            'phone': legislator_info['phone'], 'email': legislator_info['email']})
                    except KeyError:
                        print(f"No legislator found for district {district}")

    def create_senate_roster(self):
        self._get_senate_members()
        self._get_senate_committees()
        self._driver.close()
        senate_roster = pandas.DataFrame.from_dict(self._senate_legislators, orient='index')
        senate_committees = pandas.DataFrame(self._senate_committees)
        senate_committees = senate_committees.set_index(['committee', 'district']).sort_values(
            by=['committee', 'leadership_position', 'last_name'])
        with pandas.ExcelWriter("Legislator_Roster_2026.xlsx") as writer:
            senate_roster.to_excel(writer, sheet_name="MO Senate", index=False)
            senate_committees.to_excel(writer, sheet_name="Senate Committees")


    def create_joint_roster(self):
        self._get_house_members()
        self._get_house_committees()
        house_roster = pandas.DataFrame.from_dict(self._house_legislators, orient='index')
        house_committees = pandas.DataFrame.from_dict(self._house_committees, orient='index')
        self._get_senate_members()
        self._get_senate_committees()
        self._driver.close()
        senate_roster = pandas.DataFrame.from_dict(self._senate_legislators, orient='index')
        senate_committees = pandas.DataFrame(self._senate_committees)
        senate_committees = senate_committees.set_index(['committee', 'district']).sort_values(
            by=['committee', 'leadership_position', 'last_name'])
        with pandas.ExcelWriter("Legislator_Roster_2026.xlsx") as writer:
            house_roster.to_excel(writer, sheet_name="MO House", index=False)
            house_committees.to_excel(writer, sheet_name="House Committees")
            senate_roster.to_excel(writer, sheet_name="MO Senate", index=False)
            senate_committees.to_excel(writer, sheet_name="Senate Committees")