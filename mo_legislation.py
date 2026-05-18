import re

import pandas
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
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
        self._house_committees = {}
        self._senate_committees = []

    def _get_house_members(self):
        r = requests.get("https://documents.house.mo.gov/xml/254-MemberList.XML")
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
            self._house_legislators[district] = {
                'first_name': first_name,
                'last_name': last_name,
                'district': district,
                'party': party,
                'room': room,
                'home_town': home_town,
                'year_entered': year_entered,
                'phone': phone,
                'email': email
            }

    def _get_house_committees(self):
        r = requests.get("https://documents.house.mo.gov/xml/254-CommitteeList.XML")
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
                        self._house_committees[committee_name].append({
                            'first_name': legislator_info['first_name'],
                            'last_name': legislator_info['last_name'],
                            'district': district,
                            'leadership_position': leadership_position,
                            'party': legislator_info['party'],
                            'room': legislator_info['room'],
                            'home_town': legislator_info['home_town'],
                            'year_entered': legislator_info['year_entered'],
                            'phone': legislator_info['phone'],
                            'email': legislator_info['email']
                        })
                    except KeyError:
                        print(f"No legislator found for district {district}")

    def create_house_roster(self):
        self._get_house_members()
        self._get_house_committees()
        house_roster = pandas.DataFrame.from_dict(self._house_legislators, orient='index')
        house_committees = pandas.DataFrame.from_dict(self._house_committees, orient='index')
        with pandas.ExcelWriter("Legislator_Roster_2025.xlsx") as writer:
            house_roster.to_excel(writer, sheet_name="MO House", index=False)
            house_committees.to_excel(writer, sheet_name="House Committees")

    def _get_senate_members(self):
        self._driver.get("https://www.senate.mo.gov/Senators/Directory")
        rows = self._driver.find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 5:
                self._get_senator(cells)

    def _get_senator(self, cells: list[WebElement]):
        party_district = cells[1].text.split('-')
        party = party_district[0]
        district = party_district[1]
        room = cells[3].text
        phone = cells[4].text
        link = cells[0].find_element(By.TAG_NAME, "a")
        name = link.text.split(' ')
        self._senate_legislators = {
            'first_name': ' '.join(name[:-1]),
            'last_name': ' '.join(name[-1:]),
            'district': int(district),
            'party': party,
            'room': room,
            'home_town': None,
            'year_entered': None,
            'phone': phone,
            'email': None
        }
        link.click()
        year_elected = self._driver.find_element(By.XPATH, "//strong[contains(text(), 'First elected to the Senate')]")
        try:
            year_entered = re.match(r'First elected to the Senate:.*(\d{4})', year_elected.text).group(1)
            self._senate_legislators[district]['year_entered'] = year_entered
        except AttributeError:
            print(f"Missing year elected for {district}")
        email_link = self._driver.find_element(By.LINK_TEXT, 'Email the Senator')
        email_link.click()
        self._wait.until(EC.visibility_of_element_located((By.ID, "MainContent_btnContinue")))
        continue_btn = self._driver.find_element(By.ID, "MainContent_btnContinue")
        continue_btn.click()
        self._wait.until(EC.visibility_of_element_located((By.ID, "MainContent_txtName")))
        self._senate_legislators[district]['email'] = self._driver.current_url
        self._driver.back()
        self._wait.until(EC.visibility_of_element_located((By.ID, "MainContent_btnContinue")))
        self._driver.back()
        self._wait.until(EC.visibility_of_element_located((By.ID, "biography")))
        self._driver.back()
        self._wait.until(EC.visibility_of_element_located((By.TAG_NAME, "table")))

    def _get_senate_committees(self):
        self._driver.get("https://www.senate.mo.gov/committees/index")
        comm_list = self._driver.find_element(By.CSS_SELECTOR, "#main ul")
        links = [link.get_attribute("href") for link in comm_list.find_elements(By.TAG_NAME, "a")]
        links = [link for link in links if not link.lower().endswith('pdf')][:1]
        for link in links:
            self._get_senate_committees_by_type(link)

    def _get_senate_committees_by_type(self, url):
        self._driver.get(url)
        committees = [committee.get_attribute("href") for committee in self._driver.find_elements(By.CSS_SELECTOR, '#main p > a:first-child') if
                      not committee.get_attribute("href").lower().endswith("pdf")][:2]
        for committee in committees:
            self._get_committee_page(committee)

    def _get_committee_page(self, url):
        self._driver.get(url)
        members = self._driver.find_elements(By.CLASS_NAME, "panel-senator")
        committee_name = self._driver.find_element(By.CLASS_NAME, "entry-title").text
        for member in members:
            link = member.find_element(By.TAG_NAME, 'a')
            district = int(''.join(link.get_attribute("href").split('/')[-1:]))
            senator_name = member.find_element(By.CLASS_NAME, "senator-Text").text.split(' ')
            try:
                legislator_info = self._senate_legislators[district]
                self._senate_committees.append({
                    'committee': committee_name,
                    'district': str(district),
                    'first_name': legislator_info['first_name'],
                    'last_name': legislator_info['last_name'],
                    'leadership_position': senator_name[2] if len(senator_name) == 3 else '',
                    'party': legislator_info['party'],
                    'room': legislator_info['room'],
                    'home_town': legislator_info['home_town'],
                    'year_entered': legislator_info['year_entered'],
                    'phone': legislator_info['phone'],
                    'email': legislator_info['email']
                })
            except KeyError:
                print(f"No legislator found for district {district}")

    def _get_mock_senate(self):
        roster = pandas.read_csv("senate_roster.csv", index_col='district')
        self._senate_legislators = pandas.DataFrame.to_dict(roster, orient='index')

    def create_senate_roster(self):
        self._get_senate_members()
        self._get_senate_committees()
        self._driver.close()
        senate_roster = pandas.DataFrame.from_dict(self._senate_legislators, orient='index')
        senate_committees = pandas.DataFrame(self._senate_committees)
        senate_committees = senate_committees.set_index(['committee', 'district']).sort_values(by=['committee', 'leadership_position', 'last_name'])
        print(senate_committees)
        with pandas.ExcelWriter("Legislator_Roster_2025.xlsx") as writer:
            senate_roster.to_excel(writer, sheet_name="MO Senate", index=False)
            senate_committees.to_excel(writer, sheet_name="Senate Committees")

# TODO: Format committees properly in Excel
# TODO: Joint committees
# TODO: Consider whether to use the data frame earlier?
# TODO: Refactor into separate classes?
