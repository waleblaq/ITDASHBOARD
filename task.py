# +
import os
import time
from RPA.Excel.Files import Files
from datetime import timedelta
from RPA.PDF import PDF
from time import sleep
from RPA.Browser.Selenium import Selenium
import logging

OUTPUT_PATH = "output"

if not os.path.exists(OUTPUT_PATH):
    os.mkdir(OUTPUT_PATH)


# +
class ThoughtfulAutomation:
    Department = []
    headers = []

    def __init__(self):
        self.browser = Selenium()
        self.browser.set_download_directory(os.path.join(os.getcwd(), f"{OUTPUT_PATH}"))
        self.browser.open_available_browser("https://itdashboard.gov/")
        self.files = Files()
        self.pdf = PDF()

# Creating important functions To be used in the Class

    def create_workbook(self):
        try:
            self.files.create_workbook("output/company.xlsx")
        except Exception as e:
            logging.error(f"Failed to create workbook: {str(e)}")

    def get_all_agencies(self):
        self.browser.wait_until_page_contains_element('//a[@class="btn btn-default btn-lg-2x trend_sans_oneregular"]')
        self.browser.find_element('//a[@class="btn btn-default btn-lg-2x trend_sans_oneregular"]').click()
        self.browser.wait_until_page_contains_element(locator="id:agency-tiles-container")
        self.Department = self.browser.find_elements(
            '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        try:
            agency_tile_list = []
            for section in range(1, 50):
                if self.browser.is_element_visible(locator=f"CSS:#agency-tiles-widget > div > div:nth-child({section})"):
                    for agency in range(1, 10):
                        if self.browser.is_element_visible(locator=f"CSS:#agency-tiles-widget > div > div:nth-child({section}) > div:nth-child({agency})"):
                            agency_tile_list.append((
                                self.browser.get_text(locator=f"CSS:#agency-tiles-widget > div > div:nth-child({section}) > div:nth-child({agency}) > div > div > div > div:nth-child(2) > a > span.h4.w200"),
                                self.browser.get_text(locator=f"CSS:#agency-tiles-widget > div > div:nth-child({section}) > div:nth-child({agency}) > div > div > div > div:nth-child(2) > a > span.h1.w900"),
                            ))
                        else:
                            break
                else:
                    break
            return agency_tile_list
        except Exception as e:
            logging.error.error(f"failed to agencies due to: {str(e)}")

    def write_agencies_tiles(self, data):
        try:
            self.files.create_worksheet(name='Agencies_tiles')
            self.files.set_worksheet_value(row=1, column=1, value='Agency')
            self.files.set_worksheet_value(row=1, column=2, value='Tile')
            row = 2
            for Agency, Tile in data:
                self.files.set_worksheet_value(row=row, column=1, value=Agency)
                self.files.set_worksheet_value(row=row, column=2, value=Tile)
                row += 1
        except Exception as e:
            logging.error(f"Failed to write agencies tiles: {str(e)}")

    def save_excel(self):
        self.files.remove_worksheet('Sheet')
        self.files.save_workbook("output/company.xlsx")
        self.files.close_workbook()

    def download_wait(self, path_to_downloads):
        seconds = 0
        d_wait = True
        while d_wait and seconds < 20:
            time.sleep(1)
            dl_wait = False
            for fname in os.listdir(path_to_downloads):
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds += 1
        return seconds

    def same_pdf(self, uii, nametag):
        self.pdf.extract_pages_from_pdf(
            source_path=f"output/{uii}.pdf",
            output_path=f"output/pdfs/page{uii}.pdf",
            pages=1
        )
        conditioner = self.pdf.get_text_from_pdf(f"output/pdfs/page{uii}.pdf")
        for i in conditioner:
            if nametag and uii in conditioner[i]:
                return True
            else:
                return False

    def provide_headers(self):
        while True:
            try:
                heads = self.browser.find_element(
                    '//table[@class="datasource-table usa-table-borderless dataTable no-footer"]'
                ).find_element_by_tag_name(
                    "thead").find_elements_by_tag_name("tr")[1].find_elements_by_tag_name("th")
                if heads:
                    break
            except:
                sleep(2)
        for item in heads:
            self.headers.append(item.text)

    def click_agency(self):
        agency_to_click = "Department of Commerce"
        for section in range(1, 50):
            if self.browser.is_element_visible(f"CSS:#agency-tiles-widget > div > div:nth-child({section})"):
                for agency in range(1, 10):
                    if self.browser.is_element_visible(f"CSS:#agency-tiles-widget > div > div:nth-child({section}) > div:nth-child({agency})"):
                        if agency_to_click == self.browser.get_text(f"CSS:#agency-tiles-widget > div > div:nth-child({section}) > div:nth-child({agency}) > div > div > div > div:nth-child(2) > a > span.h4.w200"):
                            try:
                                self.browser.click_link(f"CSS:#agency-tiles-widget > div > div:nth-child({section}) > div:nth-child({agency}) > div > div > div > div:nth-child(2) > a")
                                return
                            except:
                                logging.critical(f"failed to find this agency: {agency_to_click}")
                    else:
                        break
            else:
                break

    def scrape_table(self, department_to_open):
        self.browser.wait_until_page_contains_element('//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        self.Department = self.browser.find_elements(
            '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        particular_agency = self.Department[department_to_open]
        self.browser.wait_until_page_contains_element(particular_agency)
        Table_link = self.browser.find_element(particular_agency).find_element_by_tag_name("a").get_attribute("href")
        self.browser.go_to(Table_link)
        self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_info"]',
                                                      timeout=timedelta(seconds=50))
        raw_data = self.browser.find_element('//*[@id="investments-table-object_info"]')
        agency_data = raw_data.text.split(" ")
        total_rows = int(agency_data[-2])
        self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_length"]/label/select')
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select').click()
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        self.browser.wait_until_page_contains_element(
            f'//*[@id="investments-table-object"]/tbody/tr[{total_rows}]/td[1]', timeout=timedelta(seconds=20))
        self.provide_headers()
        num_col = len(self.headers)
        headings = []
        columns = []
        column_element = []
        pdf_match = ["Pdf Match", ]
        for a in range(0, num_col):
            self.headers[a] = [self.headers[a]]
            headings.append(self.headers[a])
        for i in range(1, num_col + 1):
            cell = []
            for j in range(1, total_rows + 1):
                self.browser.wait_until_page_contains_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{total_rows}]/td[{num_col}]', timeout=timedelta(seconds=20))
                cells = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{j}]/td[{i}]').text
                cell.append(cells)
            column_element.append(cell)
        for a in range(0, num_col):
            col = headings[a] + column_element[a]
            columns.append(col)
        columns.append(pdf_match)
        for q in range(1, total_rows + 1):
            try:
                url = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{q}]/td[1]').find_element_by_tag_name(
                    "a").get_attribute("href")
            except:
                url = ''
            if url:
                uii_nos = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{q}]/td[1]').text
                investment_title = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{q}]/td[3]').text
                self.browser.go_to(url)
                self.browser.wait_until_page_contains_element('//div[@id="business-case-pdf"]')
                self.browser.find_element('//div[@id="business-case-pdf"]').click()
                self.download_wait(f"{OUTPUT_PATH}")
                self.browser.go_to(Table_link)
                self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_length"]/label/select',
                                                              timeout=timedelta(seconds=20))
                self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_length"]/label/select')
                self.browser.find_element('//*[@id="investments-table-object_length"]/label/select').click()
                self.browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
                self.browser.wait_until_page_contains_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{total_rows}]/td[{num_col}]', timeout=timedelta(seconds=20))
            if url:
                check = self.same_pdf(uii_nos, investment_title)
                if check:
                    match = "pdf is match"
                else:
                    match = "pdf did not match"
            else:
                match = "nan"
            pdf_match.append(match)
        value = {}
        for x in range(0, num_col + 1):
            column_headers = columns[x][0]
            value[column_headers] = columns[x]
        excel_workbook = self.files.create_workbook("output/Agency table.xlsx")
        excel_workbook.append_worksheet("Sheet", value)
        excel_workbook.rename_worksheet("Agency_table", "Sheet")
        excel_workbook.save()


if __name__ == "__main__":
    obj = ThoughtfulAutomation()
    obj.create_workbook()
    AB = obj.get_all_agencies()
    obj.write_agencies_tiles(AB)
    obj.save_excel()
    obj.scrape_table(24)
# -


