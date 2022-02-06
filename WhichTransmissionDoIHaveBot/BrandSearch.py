# Imports
from selenium import webdriver
import xlsxwriter
import time

# Define the variables and inputs
line = 1
merch = input("bir marka girin: ")

# Open an excel doc
workbook = xlsxwriter.Workbook("Which Transmission Do I Have " + merch + ".xlsx")
sheet = workbook.add_worksheet("")
sheet.write(0, 0, "Manufacturer")
sheet.write(0, 1, "Model")
sheet.write(0, 2, "Year")
sheet.write(0, 3, "Engine")
sheet.write(0, 4, "Article Number")
sheet.write(0, 5, "Transmission Brand")
sheet.write(0, 6, "Transmission Type")

# Open the website
browser = webdriver.Chrome()
browser.get("https://www.automaticchoice.com/de-ch/getriebeteile/welches-getriebe-habe-ich")
brandslist = []
brandsvaluelist = []
time.sleep(5)

# Get the manufacturer names and values
for tag in browser.find_elements_by_css_selector("select[name = manufacturer] option"):
    brandvalue = tag.get_attribute('value')
    brandsvaluelist.append(brandvalue)
    brand = tag.text
    brandslist.append(brand)
del brandsvaluelist[0]
del brandsvaluelist[116:]
del brandslist[0]
del brandslist[116:]
brandslist.pop()
brandsvaluelist.pop()
branddict = dict(zip(brandsvaluelist, brandslist))
merchdict = dict(zip(brandslist, brandsvaluelist))


# Reach to the model page
bl = merchdict[merch]
browser.get("https://www.automaticchoice.com/de-ch/getriebeteile/welches-getriebe-habe-ich?manufacturer="+bl+"&type=transmission")
modellist = []
modelvaluelist = []
time.sleep(20)
# Get the model names and values
for m in browser.find_elements_by_css_selector("select[name = model] option"):
    modelvalue = m.get_attribute("value")
    modelvaluelist.append(modelvalue)
    model = m.text
    modellist.append(model)
del modellist[0]
del modelvaluelist[0]
modellist.pop()
modelvaluelist.pop()
modeldict = dict(zip(modelvaluelist, modellist))
# Reach to the year page
for ml in modelvaluelist:
    browser.get("https://www.automaticchoice.com/de-ch/getriebeteile/welches-getriebe-habe-ich?manufacturer="+bl+"&model="+ml+"&type=transmission")
    yearlist = []
    yearvaluelist = []
    time.sleep(3)
    # Get the year names and values
    for y in browser.find_elements_by_css_selector("select[name = year] option"):
        yearvalue = y.get_attribute("value")
        yearvaluelist.append(yearvalue)
        year = y.text
        yearlist.append(year)
    del yearvaluelist[0]
    del yearlist[0]
    yearvaluelist.pop()
    yearlist.pop()
    yeardict = dict(zip(yearvaluelist, yearlist))

    # Reach to the engine page
    for yl in yearvaluelist:
        browser.get("https://www.automaticchoice.com/de-ch/getriebeteile/welches-getriebe-habe-ich?manufacturer="+bl + "&model="+ml+"&year="+yl+"&type=transmission")
        enginelist = []
        enginevaluelist = []
        time.sleep(1)

        # Get the engine names and values
        for e in browser.find_elements_by_css_selector("select[name = engine] option"):
            enginevalue = e.get_attribute("value")
            enginevaluelist.append(enginevalue)
            engine = e.text
            enginelist.append(engine)
        del enginevaluelist[0]
        del enginelist[0]
        enginevaluelist.pop()
        enginelist.pop()
        enginedict = dict(zip(enginevaluelist, enginelist))

        # Reach to the result page
        for el in enginevaluelist:
            browser.get("https://www.automaticchoice.com/de-ch/getriebeteile/welches-getriebe-habe-ich?manufacturer="+bl+"&model="+ml+"&year="+yl+"&engine="+el+"&type=transmission")
            resultlist = []
            time.sleep(1)
            # Listing the article number, brand name and transmission type from result list
            for r in browser.find_elements_by_css_selector("ul[class = results] li"):
                result = r.text
                resultlist.append(result)
            # Write the datas to excel file
            a = len(resultlist)/6
            b = 0
            c = 1
            while c <= a:
                sheet.write(line, 0, branddict[bl])
                sheet.write(line, 1, modeldict[ml])
                sheet.write(line, 2, yeardict[yl])
                sheet.write(line, 3, enginedict[el])
                sheet.write(line, 4, resultlist[1 + b])
                sheet.write(line, 5, resultlist[2 + b])
                sheet.write(line, 6, resultlist[3 + b])
                line = line + 1
                b += 6
                c += 1
workbook.close()
browser.close()


# Written by SRU
