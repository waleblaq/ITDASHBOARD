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
class Automatorr:
    department = []
    headers = []
    locator = '//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]'

    def __init__(self):
        self.browser = Selenium()
        self.browser.set_download_directory(os.path.join(os.getcwd(), f"{OUTPUT_PATH}"))
        self.browser.open_available_browser("https://itdashboard.gov/")
        self.files = Files()
        self.pdf = PDF()

# Creating important functions To be used in the Class

    def get_all_agencies(self):
        self.browser.wait_until_page_contains_element("//*[@id='node-23']")
        self.browser.click_element("//*[@id='node-23']")
        self.browser.wait_until_page_contains_element(self.locator)
        self.department = self.browser.find_elements(self.locator)
        try:
            global records
            agency_tile_list = ["Agency", ]
            amount = ["Amount", ]
            for dept in self.department:
                agency_tile_list.append(dept.text.split('\n')[0])
                amount.append(dept.text.split('\n')[2])
        except Exception as e:
            logging.error(f"failed to agencies due to: {str(e)}")
        try:
            records = {"Agency": agency_tile_list,  "Amount": amount}
        except Exception:
            print("Unable to populate the records to worksheet\n")

    def Create_exel_file_for_agency(self):
        excel_workbook = self.files.create_workbook("output/Agency_Amount_table.xlsx")
        excel_workbook.append_worksheet("Sheet", records)
        excel_workbook.save()

    def download_wait(self, path_to_downloads):
        seconds = 0
        d_wait = True
        while d_wait and seconds < 10:
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

    def Open_department(self, department_to_open):
        global Table_link
        particular_agency = self.department[department_to_open]
        self.browser.wait_until_page_contains_element(particular_agency)
        Table_link = self.browser.find_element(particular_agency).find_element_by_tag_name("a").get_attribute("href")
        return Table_link

    def Get_table(self):
        global total_rows
        global agency_tile_list
        self.browser.go_to(self.Open_department(24))
        self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_info"]',
                                                      timeout=timedelta(seconds=50))
        raw_data = self.browser.find_element('//*[@id="investments-table-object_info"]')
        #agency_data = raw_data.text.split(" ")
        total_rows = int(raw_data.text.split(" ")[-2])

    def get_out_all_table_rows(self):
        self.browser.wait_until_page_contains_element('//*[@id="investments-table-object_length"]/label/select', timeout=timedelta(seconds=20))
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select').click()
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        self.browser.wait_until_page_contains_element(
            f'//*[@id="investments-table-object"]/tbody/tr[{total_rows}]/td[1]', timeout=timedelta(seconds=20))

    def Getting_table_element(self):
        global pdf_match
        global columns
        global num_col
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

    def get_url(self):
        global url_list
        global url
        url_list = []
        for q in range(1, total_rows + 1):
            try:
                url = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{q}]/td[1]').find_element_by_tag_name(
                    "a").get_attribute("href")
            except:
                url = ''
            if url:
                url_list.append(url)

    def download_pdf(self):
        for url in url_list:
            self.browser.go_to(url)
            self.browser.wait_until_page_contains_element('//div[@id="business-case-pdf"]')
            self.browser.find_element('//div[@id="business-case-pdf"]').click()
            self.download_wait(f"{OUTPUT_PATH}")

    def populating_pdf_match_column(self):
        self.browser.go_to(Table_link)
        self.get_out_all_table_rows()
        self.browser.wait_until_page_contains_element(
            f'//*[@id="investments-table-object"]/tbody/tr[{total_rows}]/td[{num_col}]', timeout=timedelta(seconds=20))
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
                check = self.same_pdf(uii_nos, investment_title)
                if check:
                    match = "pdf is match"
                else:
                    match = "pdf did not match"
            else:
                match = "nan"
            pdf_match.append(match)
            
    def Create_exel_file(self):
        value = {}
        for x in range(0, num_col + 1):
            column_headers = columns[x][0]
            value[column_headers] = columns[x]
        excel_workbook = self.files.create_workbook("output/Agency table.xlsx")
        excel_workbook.append_worksheet("Sheet", value)
        excel_workbook.rename_worksheet("Agency_table", "Sheet")
        excel_workbook.save()

if __name__ == "__main__":
    obj = Automatorr()
    obj.get_all_agencies()
    obj.Create_exel_file_for_agency()
    obj.Get_table()
    obj.get_out_all_table_rows()
    obj.Getting_table_element()
    obj.get_url()
    obj.download_pdf()
    obj.populating_pdf_match_column()
    obj.Create_exel_file()
# -
# ## 
