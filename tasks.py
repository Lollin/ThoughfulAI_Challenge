from RPA.Browser.Selenium import Selenium
from robocorp.tasks import task
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP
from robocorp import workitems
from datetime import datetime, timedelta
import dateparser
import os
import re

# FILE_NAME = "challenge.xlsx"
# OUTPUT_DIR = Path(os.environ.get("ROBOT_ARTIFACTS"))
# EXCEL_URL = f"https://rpachallenge.com/assets/downloadFiles/{FILE_NAME}"

browser = Selenium()
excel = Files()
http = HTTP()


# Load configurations

for item in workitems.inputs:
    search_phrase = item.payload["search_phrase"]
    print("Search: " + search_phrase)
    news_category = item.payload["news_category"]
    print("Category: " + news_category)
    if item.payload["months"] == 0:
        months = 1
    else:
        months = item.payload["months"]
end_date = datetime.now()
print("Month: " + str(months))


# Define the date range for the news articles
def is_date_within_range(date_text):
    try:
        data_interprated = dateparser.parse(date_text)
        if data_interprated is None:
            return False

        start_date = end_date - timedelta(days=end_date.day - 1)
        start_date = start_date - timedelta(days=start_date.day - 1)
        start_date = start_date.replace(day=1)
        start_date = start_date.replace(month=start_date.month - (months - 1))

        return start_date <= data_interprated <= end_date
    except ValueError:
        return False


def contains_money(text):
    money_patterns = [
        r"\$\d+(?:,\d{3})*(?:\.\d+)?",
        r"\d+(?:,\d{3})*\s*dollars",
        r"\d+(?:,\d{3})*\s*USD",
    ]
    for pattern in money_patterns:
        if re.search(pattern, text):
            return True
    return False


@task
def main():
    """
    Solve the RPA challenge
    """

    print("Opening browser")
    browser.open_browser("https://www.latimes.com/", browser="headlessfirefox")
    browser.maximize_browser_window()
    browser.wait_until_element_is_visible(
        '//span[contains(text(), "Show Search")]', timeout=10
    )
    print("Waiting element to click")
    browser.click_element(
        locator='//span[contains(text(), "Show Search")]//parent::button'
    )

    print("Put phrase of search: " + search_phrase)
    browser.input_text('//input[contains(@name, "q")]', search_phrase)
    browser.click_element(
        locator='//span[contains(text(), "Submit Search")]//parent::button'
    )
    # browser.press_key('//input[contains(@name, "q")]', "Enter")
    print("Wait new page appears")
    browser.wait_until_element_is_visible('//h1[contains(text(), "Search results")]')

    browser.select_from_list_by_value('//select[contains(@class, "select-input")]', "1")

    # Look for articles and webscrapping
    articles = browser.find_elements('//div[contains(@class, "promo-wrapper")]')
    results = []
    print(len(articles))
    for i, article in enumerate(articles):
        try:
            value = i + 1
            print("Reading the News")
            title = browser.find_element(
                locator='(//h3[contains(@class, "promo-title")])[{}]'.format(value)
            ).text
            print("Title is: " + title)
            description = browser.find_element(
                locator='(//p[contains(@class, "promo-description")])[{}]'.format(value)
            ).text
            date_text = browser.find_element(
                locator='(//p[contains(@class, "promo-timestamp")])[{}]'.format(value)
            ).text

            print(date_text)
            if not is_date_within_range(date_text):
                print("The article is old, date: " + date_text)
                continue

            image = browser.find_element(
                locator='(//div[contains(@class, "promo-media")]//a/picture/source)[{}]'.format(
                    value
                )
            )
            image_url = image.get_attribute("srcset").split(" ")[0]
            image_filename = image_url.split(sep="-", maxsplit=2)[2]
            name_base, extension = os.path.splitext(image_filename)
            print(name_base)
            print(extension)
            if not extension or len(extension) > 4:
                print("Apply extension .jpg")
                image_filename = image_filename + ".jpg"

            http.download(image_url, f"output/{image_filename}")

            money_mentioned = contains_money(title) or contains_money(description)
            search_phrase_count = title.lower().count(
                search_phrase.lower()
            ) + description.lower().count(search_phrase.lower())

            results.append(
                {
                    "title": title,
                    "date": date_text,
                    "description": description,
                    "image_filename": image_filename,
                    "search_phrase_count": search_phrase_count,
                    "money_mentioned": money_mentioned,
                }
            )
        except Exception as e:
            print(f"Error to process the article: {e}")

    excel.create_workbook("output/news_data.xlsx")
    headers = [
        "Title",
        "Description",
        "Date",
        "Image Filename",
        "Count Search Phrase",
        "Money",
    ]

    for col, header in enumerate(headers, start=1):
        excel.set_worksheet_value(1, col, header.capitalize())

    excel.set_styles("A1:F1", bold=True, font_name="Arial", size=12)
    excel.set_styles("A2:F12", font_name="Arial", size=10)

    for i, result in enumerate(results, start=2):
        excel.set_worksheet_value(i, 1, result["title"])
        excel.set_worksheet_value(i, 2, result["description"])
        excel.set_worksheet_value(i, 3, result["date"])
        excel.set_worksheet_value(i, 4, result["image_filename"])
        excel.set_worksheet_value(i, 5, result["search_phrase_count"])
        excel.set_worksheet_value(i, 6, result["money_mentioned"])

    excel.save_workbook()


if __name__ == " __main__":
    try:
        main()
    finally:
        browser.close_browser
        print("Done")
