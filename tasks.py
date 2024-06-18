from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
import pandas as pd
from time import sleep
from RPA.PDF import PDF
from RPA.Archive import Archive
from pathlib import Path
import logging
import datetime


class OrderRobots:
    def __init__(self):
        self.browser = Selenium()
        self.time_execution = datetime.now().strftime("%Y%m%d%H%M%S")
        self.setup_logging

    def setup_logging(self):
        log_name = f"log_{self.time_execution}.txt"
        log_path = Path("output") / log_name

        logging.basicConfig(
            filename=log_path,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )

    def open_robot_order_website(self, url):
        try:
            self.browser.open_available_browser(url)

            self.browser.wait_until_element_is_visible(
                "//button[@class='btn btn-danger']", timeout=20
            )

            self.browser.click_element("//button[@class='btn btn-danger']")
            sleep(1)

        except Exception as e:
            logging.error(
                f"It was not possible to load the website within the timeout period: {e}"
            )
            raise

    def download_csv_file(self):
        http = HTTP()
        http.download(
            url="https://robotsparebinindustries.com/orders.csv", overwrite=True
        )

    def get_orders(self):
        self.data = pd.read_csv("orders.csv")

    def fill_the_form(self):
        for index, row in self.data.iterrows():

            self.browser.select_from_list_by_value('//*[@id="head"]', str(row["Head"]))
            self.browser.click_element(f'//*[@id="id-body-{str(row["Body"])}"]')
            self.browser.input_text("//input[@type='number']", int(row["Legs"]))
            self.browser.input_text('//*[@id="address"]', str(row["Address"]))

            sleep(1)
            # self.browser.click_element('//*[@id="preview"]')
            self.browser.click_element('//*[@id="order"]')
            sleep(1)
            if self.browser.is_element_visible("//div[@class='alert alert-danger']"):
                while True:
                    self.browser.click_element('//*[@id="order"]')
                    if self.browser.is_element_visible('//*[@id="order-another"]'):
                        break

            self.browser.wait_until_element_is_visible(
                '//*[@id="order-another"]', timeout=10
            )

            # =-=-=-=- Calling Functions -=-=-=-=
            self.store_receipt_as_pdf(str(row["Order number"]))
            self.screenshot_robot(str(row["Order number"]))
            self.embed_screenshot_to_receipt(self.file_name, self.file_pdf)

            # =-=-=-=- Click "Order Another" -=-=-=-=
            self.browser.click_element('//*[@id="order-another"]')

            # =-=-=-=- Close Pop Up -=-=-=-=
            self.browser.click_element("//button[@class='btn btn-danger']")
            sleep(1)

    def store_receipt_as_pdf(self, order_number):
        html_content = self.browser.get_element_attribute(
            '//*[@id="receipt"]', "outerHTML"
        )
        pdf = PDF()

        self.file_pdf = f"output/receipts/{order_number}_receipt.pdf"
        pdf.html_to_pdf(html_content, self.file_pdf)

    def screenshot_robot(self, order_number):
        self.file_name = f"output/receipts/{order_number}.png"
        self.browser.capture_element_screenshot(
            '//*[@id="robot-preview-image"]', filename=self.file_name
        )

    def embed_screenshot_to_receipt(self, screenshot, pdf_file):
        pdf = PDF()
        pdf.add_watermark_image_to_pdf(
            image_path=screenshot, source_path=pdf_file, output_path=pdf_file
        )

    def archive_receipts(self):
        lib = Archive()
        lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

    def main_task(self):
        attempts = 0
        max_attempts = 3
        while attempts < max_attempts:
            try:
                self.open_robot_order_website(
                    "https://robotsparebinindustries.com/#/robot-order"
                )
                self.download_csv_file()
                self.get_orders()
                self.fill_the_form()
                self.archive_receipts()
                logging.info("Robot executed successfully.")
                break
            except Exception as error:
                attempts += 1
                logging.error(f"Attempt {attempts} failed with error: {error}")
            finally:
                self.browser.close_all_browsers()

            if attempts == max_attempts:
                logging.error("Maximum attempts reached, robot execution failed.")


@task
def run_robot():
    robot = OrderRobots()
    robot.main_task()
