import pandas as pd
from selenium import webdriver
import time
import pandas
import openpyxl
import difflib
import re
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains


chrome_options = webdriver.ChromeOptions()
prefs = {"profile.managed_default_content_settings.images": 2,}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome('chromedriver.exe', chrome_options=chrome_options)
driver.get("http://www.amazon.com")
assert "Amazon" in driver.title
driver.maximize_window()
# AMZ ZIP CODE
elemZip = driver.find_element_by_id("nav-global-location-slot").click()
driver.implicitly_wait(3)
time.sleep(10)

elemZipCode = driver.find_element_by_id("GLUXZipUpdateInput")
elemZipCode.clear()
elemZipCode.send_keys("10001")
elemZipCode.send_keys(Keys.RETURN)
elemZipCode = driver.find_element_by_xpath("//html").click();

elemTitle = driver.find_element_by_name("field-keywords")
elemTitle.clear()
elemTitle.send_keys("Garlic")
elemTitle.send_keys(Keys.RETURN)
assert "No results found." not in driver.page_source

def similarity(s1, s2):
  normalized1 = s1.lower()
  normalized2 = s2.lower()
  matcher = difflib.SequenceMatcher(None, normalized1, normalized2)
  return matcher.ratio()

def num(s):
    try:
        return int(s)
    except ValueError:
        return float(s)

def is_number(s):
    try:
        float(s) or int(s)
        return True
    except ValueError:
        return False

# Change zip code in ebay
driver.get("https://www.ebay.com/itm/183490728837")
driver.find_element_by_xpath("//*[@id='viTabs_1']").click()
driver.find_element_by_name('country').click()
driver.find_element_by_xpath("//*[@id='shCountry']/option[95]").click()
ebayZip = driver.find_element_by_name('zipCode')
ebayZip.send_keys("10001")
ebayZip.send_keys(Keys.RETURN)
driver.find_element_by_name('getRates').click()
time.sleep(5)
# Spreadsheet

excel_data_df = pandas.read_excel('my_books.xlsx', sheet_name='Лист1')
amz_links = excel_data_df['LINKS'].tolist()

book = load_workbook('./my_books.xlsx')
sheet = book.active

counter = 2
for item in amz_links:
    amz_linksSplit = item

    try:
        driver.get('%s' % amz_linksSplit)
    except TimeoutException:
        driver.refresh()
    amz_title = driver.find_element_by_id("productTitle").text
    sheet['B%s' % counter] = amz_title


    # Change brand name
    try:
        amz_brand = driver.find_element_by_id("bylineInfo").text
    except NoSuchElementException:
        amz_brand = "No Brand"
    removal_list = ["Visit", "the", "Store", "Brand:"]
    edit_string_as_list = amz_brand.split()
    amz_brand = [word for word in edit_string_as_list if word not in removal_list]
    amz_brand_name = ' '.join(amz_brand)
    sheet['D%s' % counter] = amz_brand_name

    # Change brand name
    amz_title_removal_list = [amz_brand_name, "for", ":", ";"]
    edit_string_as_list2 = amz_title.split()
    amz_title = [word for word in edit_string_as_list2 if word not in amz_title_removal_list]
    amz_title =  ' '.join(amz_title)

    # AMZ PRICE BLOCK
    try:
        AmzOut = driver.find_element_by_xpath('//*[@id="outOfStock"]/div/div[1]/span[1]').text

    except NoSuchElementException:
        AmzOut = 'есть в наличии'

    if AmzOut == 'Currently unavailable.':
        counter += 1
        continue


    try:
        amz_price = driver.find_element_by_xpath('//*[@id="tp-tool-tip-subtotal-price-value"]/span[1]').text
    except NoSuchElementException:
        print("не нашло")
        amz_price = 0
    if not amz_price or amz_price == 0:
        try:
            amz_price = driver.find_element_by_id('tp_price_block_total_price_ww').text
        except NoSuchElementException:
            print("не нашло")
        if not amz_price:
            try:
                amz_price = driver.find_element_by_id("price_inside_buybox").text
            except NoSuchElementException:
                print("нет цены")
            if not amz_price:
                try:
                    amz_price = driver.find_element_by_xpath('//*[@id="olp_feature_div"]/div[2]/span[1]/a/span[2]').text
                except NoSuchElementException:
                    print('Нет цены')
                if not amz_price:
                    try:
                        amz_price = driver.find_element_by_xpath('//*[@id="corePrice_feature_div"]/div/span/span[1]').text
                    except NoSuchElementException:
                         print("Нет цены")
                    if not amz_price:
                        try:
                            amz_price = driver.find_element_by_xpath('//*[@id="corePrice_feature_div"]/div/span/span[2]').text
                        except NoSuchElementException:
                            print("Нет цены")
                        if not amz_price:
                            try:
                                amz_price = driver.find_element_by_xpath('//*[@id="corePrice_desktop"]/div/table/tbody/tr/td[2]/span[1]/span[1]/span[2]').text
                            except NoSuchElementException:
                                 print("Нет цены")
                            if not amz_price:
                                amz_price = "0"

    amz_price1 = amz_price.replace("$", "")

        # Category filtering on Amazon
    amz_category = driver.find_element_by_xpath('//*[@id="wayfinding-breadcrumbs_feature_div"]/ul/li[1]/span/a').text
    amz_blacklist = ['Electronics','Toys & Games','Automotive']
    if amz_category in amz_blacklist:
        sheet['B%s' % counter] = 'Запрещенная категория'
        counter += 1
        continue
    sheet['C%s' % counter] = amz_price1
    book.save('my_books.xlsx')

    # Grab sellers
    try:
        driver.find_element_by_class_name('olp-text-box').click()
    except NoSuchElementException:
        print(None)
    try:
        driver.find_element_by_xpath('//*[@id="buybox-see-all-buying-choices"]/span/a').click()
    except NoSuchElementException:
        print(None)

    try:
        amz_sellers = driver.find_elements_by_xpath("//*[@id='aod-offer-soldBy']/div/div/div[2]/a")
    except NoSuchElementException:
        amz_sellers = driver.find_element_by_xpath('//*[@id="sellerProfileTriggerId"]')

    with open('NeuroDropper.txt', 'a') as f:
        for item in amz_sellers:
            f.write("%s\n" % item.get_attribute('href'))
    f.close()

 # Check brand TM
    driver.get('https://www.trademarkia.com/')
    try:
        tm = driver.find_element_by_name("ctl00$mainBody$txtSearch")
    except NoSuchElementException:
        driver.refresh()

    tm.clear()
    try:
        tm.send_keys(amz_brand)
    except WebDriverException:
        amz_brand = 'No brand'
        tm.send_keys(amz_brand)
    tm.send_keys(Keys.RETURN)

    amz_brand_tm = driver.find_elements_by_class_name('status-title')

    for btm in amz_brand_tm:

        if btm.text == "registered" or btm.text == "registered and renewed":
            sheet['E%s' % counter] = "TM"
        else:
            print("Не тм")
    book.save('my_books.xlsx')

    # Search supplier
    driver.get("http://www.ebay.com")
    ebaySearch = driver.find_element_by_id('gh-ac')
    ebaySearch.click()
    ebayAdvanceOpt = driver.find_element_by_id('gh-as-a').click()
    ebayAdvanceOptTitle = driver.find_element_by_name('_nkw')
    ebayAdvanceOptTitle.clear()
    ebayAdvanceOptTitle.send_keys(amz_title)  # отправляем текст в строку поиска
    driver.find_element_by_id('LH_TitleDesc').click()
    driver.find_element_by_id('LH_ItemConditionNew').click()
    driver.find_element_by_id('LH_LocatedInRadio').click()
    driver.find_element_by_id('LH_IPP').click()
    driver.find_element_by_xpath('//*[@id="LH_IPP"]/option[1]').click()
    driver.find_element_by_id('searchBtnLowerLnk').click()

    EbayItems = driver.find_elements_by_xpath("//*[@class='lvtitle']/a")

    links = []
    for i in range(len(EbayItems)):

        links.append(EbayItems[i].get_attribute('href'))

    for ei in links[:15]:
        try:
            driver.get(ei)
        except TimeoutException:
            driver.refresh()

        try:
            EbayItemRating = driver.find_element_by_class_name('mbg-l').text
        except NoSuchElementException:
            EbayItemRating = '1000'

        if not EbayItemRating or EbayItemRating == '(':
            EbayItemRating = driver.find_element_by_xpath('//*[@id="vi-slrpres-olp"]/div/div[1]/span/a').text
        if not EbayItemRating:
            EbayItemRating = '1000'

        EbayItemRating = EbayItemRating.replace("(", "")
        EbayItemRating = EbayItemRating.replace(")", "")
        EbayItemRating = int(EbayItemRating)

        # Ebay product title and price
        try:
            EbayItemPrice = driver.find_element_by_id('prcIsum').text
        except NoSuchElementException:
            EbayItemPrice = ''
            if not EbayItemPrice:
                try:
                    EbayItemPrice = driver.find_element_by_class_name('vi-originalPrice').text
                except NoSuchElementException:
                    print()
                if not EbayItemPrice:
                    EbayItemPrice = driver.find_element_by_id('prcIsum_bidPrice').text

        if EbayItemPrice == 'FREE':
            try:
                EbayItemPrice = driver.find_element_by_class_name('vi-originalPrice').text
            except NoSuchElementException:
                print(None)

        EbayItemPrice1 = EbayItemPrice.replace("$", "")
        EbayItemPrice2 = EbayItemPrice1.replace("US", "")
        EbayItemPrice3 = EbayItemPrice2.replace("C", "")
        EbayItemPrice4 = EbayItemPrice3.replace("AU", "")
        EbayItemPrice5 = EbayItemPrice4.replace("GBP", "")
        EbayItemPrice6 = EbayItemPrice5.replace(",", "")
        EbayItemPrice7 = EbayItemPrice6.replace("EUR", "")
        EbayItemPrice8 = EbayItemPrice7.replace("Was:\n", "")
        if 'GBP' in EbayItemPrice8:
            EbayItemPrice8 = driver.find_element_by_id('convbinPrice').text
        EbayItemTitle = driver.find_element_by_xpath('//*[@id="itemTitle"]').text

        # Ebay item Shipping block
        try:
            EbayItemShipping = driver.find_element_by_id('fshippingCost').text
        except NoSuchElementException:
            EbayItemShipping = ""
            if not EbayItemShipping:
                try:
                    EbayItemShipping = driver.find_element_by_class_name('sh_gr_bld_new').text
                except NoSuchElementException:
                    print()
                if not EbayItemShipping:
                    try:
                        EbayItemShipping = driver.find_element_by_id('shSummary').text
                    except NoSuchElementException:
                        print()


        if "$" in EbayItemShipping:
            EbayItemShipping = EbayItemShipping.replace("$", "")
            EbayItemShipping = EbayItemShipping.replace('US', '')

        if "Christmas" in EbayItemShipping:
            EbayItemShipping = EbayItemShipping.replace(" Shipping - Arrives by Christmas", "")
            EbayItemShipping = EbayItemShipping.replace("$", "")
            EbayItemShipping = EbayItemShipping.replace(" | See details", "")

        if 'GBP' in EbayItemShipping or "AU" in EbayItemShipping or "EUR" in EbayItemShipping or "C" in EbayItemShipping:
            try:
                EbayItemShipping = driver.find_element_by_id('convetedPriceId').text
            except NoSuchElementException:
                print()
            EbayItemShipping = EbayItemShipping.replace('US', '')
            EbayItemShipping = EbayItemShipping.replace('$', '')

        if is_number(EbayItemShipping):
            print(EbayItemShipping)
        else:
            EbayItemShipping = 0

        # Ebay Item Available
        try:
            EbayItemAvaliable = driver.find_element_by_id('qtySubTxt').text
        except NoSuchElementException:
            EbayItemAvaliable = "1"

        removal_list1 = ["More", "than", "available"]
        edit_string_as_list1 = EbayItemAvaliable.split()
        EbayItemAvaliable1 = [word for word in edit_string_as_list1 if word not in removal_list1]
        EbayItemAvaliable1 = ' '.join(EbayItemAvaliable1)

        if EbayItemAvaliable1 == 'Last one':
            EbayItemAvaliable2 = 1
        elif EbayItemAvaliable1 == 'Limited quantity':
            EbayItemAvaliable2 = 5
        elif 'lot' in EbayItemAvaliable1:
            EbayItemAvaliable2 = 1
        else:
            EbayItemAvaliable1 = EbayItemAvaliable1.replace(",", "")
            EbayItemAvaliable2 = int(EbayItemAvaliable1)


        DifPercent = similarity(amz_title, EbayItemTitle)
        currentUrl = driver.current_url

        Margin = num(amz_price1) * 0.85 - (num(EbayItemPrice8) + num(EbayItemShipping))
        ROI = Margin / num(EbayItemPrice8) * 100
        print(ROI)

        sheet['F%s' % counter] = num(EbayItemPrice8) + num(EbayItemShipping)

        if ROI > 15:
            if EbayItemAvaliable2 > 5:
                if EbayItemRating >= 1000:
                    if DifPercent > 0.60:
                        if sheet['G%s' % counter].value:
                            sheet['H%s' % counter] = currentUrl
                        elif sheet['H%s' % counter].value:
                            sheet['I%s' % counter] = currentUrl
                        else:
                            sheet['G%s' % counter] = currentUrl
                    else:
                        print("не подходит")

        book.save('my_books.xlsx')
    counter += 1