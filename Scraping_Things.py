from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import html
import sys, os
from datetime import datetime
import Global_var
from Global_var import Process_End
from Insert_On_databse import insert_in_Local
import ctypes
import string
import html
import wx
app = wx.App()

def ChromeDriver():
    chrome_options = Options()
    chrome_options.add_extension('C:\\BrowsecVPN.crx')
    browser = webdriver.Chrome(executable_path=str(f"C:\\chromedriver.exe"),chrome_options=chrome_options)
    browser.maximize_window()
    # browser.get("""https://chrome.google.com/webstore/detail/browsec-vpn-free-and-unli/omghfjlpggmjjaagoclmmobgdodcjboh?hl=en" ping="/url?sa=t&amp;source=web&amp;rct=j&amp;url=https://chrome.google.com/webstore/detail/browsec-vpn-free-and-unli/omghfjlpggmjjaagoclmmobgdodcjboh%3Fhl%3Den&amp;ved=2ahUKEwivq8rjlcHmAhVtxzgGHZ-JBMgQFjAAegQIAhAB""")
    wx.MessageBox(' -_-  Add Extension and Select Proxy Between 10 SEC -_- ', 'Info', wx.OK | wx.ICON_WARNING)
    time.sleep(15)  # WAIT UNTIL CHANGE THE MANUAL VPN SETtING
    browser.get("https://eproc.punjab.gov.pk/ActiveTenders.aspx")
    time.sleep(2)
    Scrap_data(browser)


def Scrap_data(browser):
    a = True
    while a == True:
        try:
            Pages = ''  # For Range purpose
            for Pages in browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolderSRIS_rdgrdManageTender_ctl00"]/tfoot/tr/td/table/tbody/tr/td/div[5]'):
                Pages = Pages.get_attribute("innerText")
                Pages = Pages.partition("in")[2].partition("page")[0].replace('<strong>', '').replace('</strong>',                                                                                                     '').strip()
                break
            if Pages == "":
                Pages = "1"
            else:pass
            for Next_page_Tab in range(int(Pages)):
                for add in range(1, 100, 1):
                    SagField = []
                    for data in range(45):
                        SagField.append('')
                    v = False
                    while v == False:
                        try:
                            for publish_Date in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[4]'):
                                publish_Date = publish_Date.get_attribute("innerText")
                                datetime_object = datetime.strptime(publish_Date, '%d %b %Y')
                                publish_date1 = datetime_object.strftime("%Y-%m-%d")
                                if publish_date1 >= Global_var.From_Date:
                                    print("Publish Date Alive")
                                    break
                                else:
                                    print("Publish Date Dead")
                                    ctypes.windll.user32.MessageBoxW(0, "Total: " + str(
                                        Global_var.Total) + "\n""Duplicate: " + str(
                                        Global_var.duplicate) + "\n""Expired: " + str(
                                        Global_var.expired) + "\n""Inserted: " + str(
                                        Global_var.inserted) + "\n""Skipped: " + str(
                                        Global_var.skipped) + "\n""Deadline Not given: " + str(
                                        Global_var.deadline_Not_given) + "\n""QC Tenders: "
                                                                     + str(Global_var.QC_Tenders) + "",
                                                                     "eproc.punjab.gov.pk",
                                                                     1)
                                    Global_var.Process_End()
                                    browser.close()
                                    sys.exit()

                            #  ====================================== THIS INFORMATION FOR HTML FILE PURPOSE ================================================================================
                            Global_var.Total += 1
                            for Department in browser.find_elements_by_xpath('/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(add) + ']/td[6]'):
                                Department = Department.get_attribute("innerText").upper()
                                if Department == "THQ HOSPITAL, SHORKOT":
                                    Address = "Jhang Shorkot Multan Road, Shorkot, Jhang, Punjab 35010, Pakistan" + "<br>\n" + "Phone: " + "+92 47 5310971"
                                    SagField[2] = Address.strip()

                                elif Department == "MAYO HOSPITAL LAHORE":
                                    Address = "near New Anarkali Neela Gumbad Anarkali Bazaar, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99211129"
                                    SagField[2] = Address.strip()

                                elif Department == "ARID AGRICULTURE UNIVERSITY RAWALPINI":
                                    Address = "Shamsabad, Muree Road Shamsabad, Rawalpindi, Punjab 46000, Pakistan" + "<br>\n" + "Phone: " + "+92 51 9292122"
                                    SagField[2] = Address.strip()

                                elif Department == "PUBLIC HEALTH ENGINEERING DIVISION SHEIKHUPURA":
                                    Address = "2 Lake Road, Mozang Chungi, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99212830"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF AGRICULTURE FAISALABAD":
                                    Address = "University Main Rd, Faisalabad, Punjab 38000, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9200161"
                                    SagField[2] = Address.strip()

                                elif Department == "SAHIWAL":
                                    Address = "Sahiwal, Sahiwal District, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 300 6931200"
                                    SagField[2] = Address.strip()

                                elif Department == "FAISALABAD INDUSTRIAL ESTATE DEVELOPMENT AND MANAGEMENT COMPANY":
                                    Address = "1st Floor, Chamber of Commerce Building Canal Rd, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9230234"
                                    SagField[2] = Address.strip()

                                elif Department == "THQ HOSPITAL THAL (MNS), LAYYAH.":
                                    Address = "Layyah, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "LAYYAH.":
                                    Address = "Layyah, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "ARID AGRICULTURE UNIVERSITY RAWALPINI.":
                                    Address = "Shamsabad, Muree Road Shamsabad, Rawalpindi, Punjab 46000, Pakistan" + "<br>\n" + "Phone: " + "+92 51 9292122"
                                    SagField[2] = Address.strip()

                                elif Department == "LOCAL GOVERNMENT AND COMMUNITY DEVELOPMENT FAISALABAD":
                                    Address = "Anarkali Bazaar, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99210013"
                                    SagField[2] = Address.strip()

                                elif Department == "WATER AND SANITATION AGENCY MULTAN":
                                    Address = "Rohtak House Multan, 316A, Bosan Rd, Shamsabad Colony Multan, Punjab 60000, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB SEED CORPORATION":
                                    Address = "4- Lytton Road, Lahore, Pakistan" + "<br>\n" + "Phone: " + "+92 44 2552399"
                                    SagField[2] = Address.strip()

                                elif Department == "PAKISTAN KIDNEY LIVER INSTITUTE AND RESEARCH CENTER TRUST":
                                    Address = "Knowledge City 1, PKLI Avenue, Opposite DHA Phase 6 Sector E Chota Mota Singh, Lahore, Pakistan" + "<br>\n" + "Phone: " + "+92 42 38107554"
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB HUMAN ORGAN TRANSPLANT AUTHORITY, LAHORE":
                                    Address = "39, Shadman 1, Near Shadman Market, Lahore, Pakistan" + "<br>\n" + "Tel: " + "+92-42-99206046-7" + "<br>\n" + "Fax: " + "+92-42-99206048"
                                    SagField[2] = Address.strip()

                                elif Department == "DHQ HOSPITAL FAISALABAD":
                                    Address = "Mall Road, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9200140"
                                    SagField[2] = Address.strip()

                                elif Department == "KHAWAJA FAREED UNIVERSITY OF ENGINEERING AND INFORMATION TECHNOLOGY, RAHIM YAR KHAN.":
                                    Address = "Abu Dhabi Rd, Rahim Yar Khan, Punjab 64200, Pakistan" + "<br>\n" + "Phone: " + "+92 68 5882400"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF VETERINARY AND ANIMAL SCIENCES, LAHORE.":
                                    Address = "Shaykh Abdul Qadir Jilani Rd, Data Gunj Buksh Town, Lahore, Kasur, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99211374"
                                    SagField[2] = Address.strip()

                                elif Department == "SIR GANGA RAM HOSPITAL LAHORE":
                                    Address = "Shara-I-Fatima Jinnah Queen's Road, Jubilee Town, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99200572"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT JAIL, PAKPATTAN":
                                    Address = "Pakpattan-Kamir Rd, Chak 36SP, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 457 380240"
                                    SagField[2] = Address.strip()

                                elif Department == "WATER AND SANITATION AGENCY LAHORE":
                                    Address = "31-B Zahoor Elahi Rd, Block B Gulberg 2, Lahore, Punjab 54660, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99332100"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF VETERINARY AND ANIMAL SCIENCES, LAHORE.":
                                    Address = "Shaykh Abdul Qadir Jilani Rd, Data Gunj Buksh Town, Lahore, Kasur, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99211374"
                                    SagField[2] = Address.strip()

                                elif Department == "ARID AGRICULTURE UNIVERSITY RAWALPINI":
                                    Address = "Shamsabad, Muree Road Shamsabad, Rawalpindi, Punjab 46000, Pakistan" + "<br>\n" + "Phone: " + "+92 51 9292122"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF AGRICULTURE FAISALABAD":
                                    Address = "University Main Rd, Faisalabad, Punjab 38000, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9200161"
                                    SagField[2] = Address.strip()

                                elif Department == "GOVERNMENT COLLEGE UNIVERSITY LAHORE":
                                    Address = "Katchery Rd, Anarkali Bazaar, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 111 000 010"
                                    SagField[2] = Address.strip()

                                elif Department == "AYUB AGRICULTURE RESEARCH INSTITUTE FAISALABAD":
                                    Address = "Jhang Rd, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9201671"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF GUJRAT":
                                    Address = "Jalalpur Jattan Road, Gujrat, Punjab 50700, Pakistan" + "<br>\n" + "Phone: " + "+92 53 3643112"
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB INDUSTRIAL ESTATE DEVELOPMENT AND MANAGEMENT COMPANY":
                                    Address = "Sundar Industrial Estate, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 35297203"
                                    SagField[2] = Address.strip()

                                elif Department == "PARKS AND HORTICULTURE AUTHORITY LAHORE":
                                    Address = "Gate No-4 Jillani Park 2 Jail Rd, G.O.R. - I, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99200908"
                                    SagField[2] = Address.strip()

                                elif Department == "GOVT. THQ HOSPITAL, KOTLI SATTIAN":
                                    Address = "Kotli Sattian, Rawalpindi, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "ALLAMA IQBAL MEDICAL COLLEGE AND JINNAH HOSPITAL, LAHORE":
                                    Address = "Quaid-i-Azam Campus, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99231443"
                                    SagField[2] = Address.strip()

                                elif Department == "LIVESTOCK AND DAIRY DEVELOPMENTE":
                                    Address = "Anarkali Bazaar, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99210527"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF OKARA":
                                    Address = "2-KM, Renala Khurd - Okara Road, G.T. Road, Okara N-5 Okara, Punjab 56300, Pakistan" + "<br>\n" + "Phone: " + "+92 44 2552399"
                                    SagField[2] = Address.strip()

                                elif Department == "LOCAL GOVERNMENT AND COMMUNITY DEVELOPMENT, GUJRANWALA":
                                    Address = "Office of Deputy Director LG&CD Dept, Civil Lines, Gujranwala, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 55 9200646"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF SARGODHA":
                                    Address = "University Road, Sargodha, Punjab 40100, Pakistan" + "<br>\n" + "Phone: " + "+92 48 111 867 111"
                                    SagField[2] = Address.strip()

                                elif Department == "DHQ HOSPITAL (SOUTH CITY), OKARA":
                                    Address = "Okara, Punjab 56300, Pakistan" + "<br>\n" + "Phone: " + "+92 44 2703458"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF PUNJAB":
                                    Address = "Canal Rd, Quaid-i-Azam Campus, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99231098"
                                    SagField[2] = Address.strip()

                                elif Department == "THQ LEVEL HOSPITAL, KOT SULTAN":
                                    Address = "Unnamed Road, Kot Sultan, Layyah, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF ENGINEERING AND TECHNOLOGY LAHORE":
                                    Address = "G.T Road, Staff Houses Engineering University Lahore, Lahore, Punjab 54890, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99029227"
                                    SagField[2] = Address.strip()

                                elif Department == "MUHAMMAD NAWAZ SHAREEF UNIVERSITY OF AGRICULTURE":
                                    Address = "Old Shujabad Road, Rangilpur, Multan, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 61 9201680"
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB PENSION FUND":
                                    Address = "112 Shakir Ali Ln, Tipu Block Garden Town, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 35882960"
                                    SagField[2] = Address.strip()

                                elif Department == "LAHORE DEVELOPMENT AUTHORITY, LAHORE":
                                    Address = "467-DII Khayaban-e-Firdousi, M.A Block D 2 Phase 1 Johar Town, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 111 111 532"
                                    SagField[2] = Address.strip()

                                elif Department == "LAWRENCE COLLEGE, GHORA GALI MURREE":
                                    Address = "Lawrence College Rd, Murree, Rawalpindi, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 51 3751001"
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB VOCATIONAL TRAINING COUNCIL":
                                    Address = "134-A, Madr-e- Millat Road, Industrial Area Quaid-e-Azam Industrial Estate Green Town, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 35209200"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT EDUCATION AUTHORITY, JHANG":
                                    Address = "p-647/54 Jhang - Toba Tek Singh Road, Jhang Sadar, Jhang, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 47 9200128"
                                    SagField[2] = Address.strip()

                                elif Department == "MUNICIPAL COMMITTEE, HAFIZABAD":
                                    Address = "Municipal Committee Hafizabad, Pakistan" + "<br>\n" + "Tel: " + "0547-541249" + "<br>\n" + "Fax: " + "0547-524429"
                                    SagField[2] = Address.strip()

                                elif Department == "MULTAN DEVELOPMENT AUTHORITY, MULTAN":
                                    Address = "Shahraae Iwan-e-Sanato Tijarat MDA Chowk Multan, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 61 9200827"
                                    SagField[2] = Address.strip()

                                elif Department == "DIVISIONAL FOREST OFFICER, SIALKOT FOREST DIVISION, SIALKOT":
                                    Address = "Mudassar Shaheed Rd, Sialkot Cantonment, Sialkot, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "GUJRANWALA WASTE MANAGEMENT COMPANY":
                                    Address = "2nd Floor Gujranwala Chamber of Commerce &Industry Chamber Plaza, Aiwan-e-Tijarat Road Gujranwala, 52250, Pakistan" + "<br>\n" + "Phone: " + "+92 55 9200890"
                                    SagField[2] = Address.strip()

                                elif Department == "BOARD OF INTERMEDIATE AND SECONDARY EDUCATION, GUJRANWALA":
                                    Address = "Sialkot Bypass, Faisal Town, Lohianwala, Gujranwala, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 55 9200751"
                                    SagField[2] = Address.strip()

                                elif Department == "QUAID-E-AZAM MEDICAL COLLEGE, BAHAWALPUR":
                                    Address = "Bahawalpur Cantt, Bahawalpur, Punjab 63100, Pakistan" + "<br>\n" + "Phone: " + "+92 62 2731042"
                                    SagField[2] = Address.strip()

                                elif Department == "GOVERNMENT COLLEGE WOMEN UNIVERSITY FAISALABAD":
                                    Address = "Arfa Kareem Rd, Block Z Madina Town, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9220068"
                                    SagField[2] = Address.strip()

                                elif Department == "THE BANK OF PUNJAB":
                                    Address = "BOP Tower, 10-B,Block E-II, Main Boulevard Gulberg III Lahore, Pakistan" + "<br>\n" + "Tel: " + " (042) 35783700-10" + "<br>\n" + "Fax: " + "(042) 35783713-35783975"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "feedback@bop.com.pk"
                                    SagField[8] = "www.bop.com.pk"

                                elif Department == "ADAPTIVE RESEARCH FARM, GUJRANWALA":
                                    Address = "Agriculture House, 21- Sir Sultan Agha Khan Soeim Road, Lahore, Pakistan" + "<br>\n" + "Phone: " + "(042) 99200732" + "<br>\n" + "Fax: " + "(042) 99200743"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "dgaextar@gmail.com"
                                    SagField[8] = "ext.agripunjab.gov.pk"

                                elif Department == "TEVTA":
                                    Address = "TEVTA Secretariat, 96-H, Gulberg Road, Lahore." + "<br>\n" + "Phone: " + "+92 -42 -99263055-59"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "info@tevta.gop.pk"
                                    SagField[8] = "www.tevta.gop.pk"

                                elif Department == "THQ HOSPITAL, TAXILA, RAWALPINDI":
                                    Address = "Taxila, Rawalpindi, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 51 9315474"
                                    SagField[2] = Address.strip()

                                elif Department == "THE WOMEN UNIVERSITY, MULTAN":
                                    Address = "Kutchery Campus, L.M.Q. Road, Multan, Pakistan" + "<br>\n" + "Phone: " + "061-9200811" + "<br>\n" + "Fax: " + "061-4500950"
                                    SagField[2] = Address.strip()
                                    SagField[1] = " info@wum.edu.pk"
                                    SagField[8] = "wum.edu.pk"

                                elif Department == "BATTALION COMMANDER, BATTALION NO. II PC RAWAT, RAWALPINDI":
                                    Address = "GT Road Riwat Rawalpindi, Pakistan" + "<br>\n" + "Phone: " + "+92 51 4610293"
                                    SagField[2] = Address.strip()

                                elif Department == "THQ HOSPITAL, PINDI GHEB":
                                    Address = "Pindi Gheb-Attock Rd, Pindi Gheb Town, Pindi Gheb, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 57 2350355"
                                    SagField[2] = Address.strip()

                                elif Department == "LABOUR AND HUMAN RESOURCE":
                                    Address = "62-D New Muslim Town, off Wahdat Road ,Lahore " + "<br>\n" + "Phone: " + "042-99230348"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "labour.punjab.gov.pk"

                                elif Department == "DHQ HOSPITAL FAISALABAD":
                                    Address = "Mall Road, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9200140"
                                    SagField[2] = Address.strip()

                                elif Department == "FAISALABAD INDUSTRIAL ESTATE DEVELOPMENT AND MANAGEMENT COMPANY":
                                    Address = "1st Floor, Chamber of Commerce Building Canal Rd, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9230234"
                                    SagField[2] = Address.strip()

                                elif Department == "THE ISLAMIA UNIVERSITY OF BAHAWALPUR.":
                                    Address = "University Chowk, Gulshan Colony, Bahawalpur, Punjab 63100, Pakistan" + "<br>\n" + "Phone: " + "+92-62-9250235" + "<br>\n" + "Fax: " + "+92-62-9250235"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "info@iub.edu.pk"
                                    SagField[8] = "www.iub.edu.pk"

                                elif Department == "FORESTRY WILD LIFE AND FISHERIES":
                                    Address = "Forest Colony Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()
                                    SagField[8] = "fmis.pk"

                                elif Department == "PUNJAB DAANISH SCHOOLS AND CENTERS OF EXCELLENCE AUTHORITY, LAHORE":
                                    Address = "Quaid-i-Azam Campus, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99231740"
                                    SagField[2] = Address.strip()

                                elif Department == "ALLIED HOSPITAL, FAISALABADE":
                                    Address = "Dr. Tusi Rd, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9210082"
                                    SagField[2] = Address.strip()

                                elif Department == "ALLIED HOSPITAL, FAISALABAD":
                                    Address = "Dr. Tusi Rd, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9210082"
                                    SagField[2] = Address.strip()

                                elif Department == "PROVINCIAL HIGHWAYS CIRCLE GUJRANWALA":
                                    Address = "N5, Sharifpura, Civil Lines, Gujranwala, Punjab 52250, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "SIR GANGA RAM HOSPITAL LAHORE":
                                    Address = "Shara-I-Fatima Jinnah Queen's Road, Jubilee Town, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99200572"
                                    SagField[2] = Address.strip()

                                elif Department == "WATER AND SANITATION AGENCY FAISALABAD":
                                    Address = "Opposite Allied Hospital, Dr. Tusi Rd, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 419210049-50"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "info@wasafaisalabad.gop.pk"
                                    SagField[8] = "wasafaisalabad.gop.pk"

                                elif Department == "DIRECTORATE OF AGRICULTURAL INFORMATION, PUNJAB":
                                    Address = "21-Sir Agha Khan-III Road, Lahore." + "<br>\n" + "Phone: " + "92-42-99200729" + "<br>\n" + "Fax: " + "92-42-99202911"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "dainformation@gmail.com"
                                    SagField[8] = "dai.agripunjab.gov.pk"

                                elif Department == "POPULATION WELFARE":
                                    Address = "8 abu bakar block new garden Abu Bakar Block Garden Town, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "(042) 99232436-8" + "<br>\n" + "Fax: " + "(042) 99232439"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "consultant@pwd.punjab.gov.pk"
                                    SagField[8] = "pwd.punjab.gov.pk"

                                elif Department == "PUNJAB SOCIAL SECURITY HEALTH MANAGMENT COMPANY":
                                    Address = "Manga - Raiwind Road Lahore, Kasur, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "WATER AND SANITATION AGENCY RAWALPINDI":
                                    Address = "Manga - Raiwind Road Lahore, Kasur, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()
                                    SagField[8] = "wasarwp.punjab.gov.pk"

                                elif Department == "BOARD OF INTERMEDIATE AND SECONDARY EDUCATION, LAHORE":
                                    Address = "86 Mozang Rd, Block B Jubilee Town, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99200192-197" + "<br>\n" + "Fax: " + "92 42 99200113"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.biselahore.com"
                                    SagField[1] = "info@biselahore.com"

                                elif Department == "NISHTAR MEDICAL COLLEGE AND HOSPITAL MULTAN":
                                    Address = "Nishtar Rd, Gillani Colony, Multan, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 61 9200238"
                                    SagField[2] = Address.strip()

                                elif Department == "INFRASTRUCTURE DEVELOPMENT AUTHORITY OF THE PUNJAB - IDAP":
                                    Address = "Ground Floor, 7 C-1 MM Alam Rd, Block C 1 Gulberg III, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99268324"
                                    SagField[2] = Address.strip()

                                elif Department == "MINES AND MINERALS":
                                    Address = "PMDC 13-H/9, Islamabad, Pakistan" + "<br>\n" + "Fax: " + "92-51-9265127-28"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.pmdc.gov.pk"
                                    SagField[1] = "pmdc@isb.comsats.net.pk"

                                elif Department == "WATER AND SANITATION AGENCY, LDA, GANG BAKHSH TOWN, LAHORE":
                                    Address = "31-B Zahoor Elahi Rd, Block B Gulberg 2, Lahore, Punjab 54660, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99332100"
                                    SagField[2] = Address.strip()

                                elif Department == "MINES AND MINERALS":
                                    Address = "TEVTA Secretariat, 96-H, Gulberg Road, Lahore, Pakistan" + "<br>\n" + "Phone: " + "+92-49-99263055-59"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "tevta.gop.pk"
                                    SagField[1] = "info@tevta.gop.pk"

                                elif Department == "PUNJAB SEED CORPORATION":
                                    Address = "4- Lytton Road, Lahore, Pakistan" + "<br>\n" + "Phone: " + "+042-99212571-75" + "<br>\n" + "Fax: " + "042-99212570"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.psc.agripunjab.gov.pk"
                                    SagField[1] = "info@psc.punjab.gov.pk"

                                elif Department == "FAISALABAD DEVELOPMENT AUTHORITY, FAISALABAD":
                                    Address = "Mall Road, Faisalabad, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 41 9200097"
                                    SagField[2] = Address.strip()

                                elif Department == "DIRECTORATE GENERAL SOIL SURVEY OF PUNJAB":
                                    Address = "P.O Awan Town, Multan Road, Lahore" + "<br>\n" + "Phone: " + "042-99330015" + "<br>\n" + "Fax: " + "042-99330011"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "soil.punjab.gov.pk"
                                    SagField[1] = "soil.punjab@yahoo.com"

                                elif Department == "PUBLIC HEALTH ENGINEERING DIVISION, LAHORE":
                                    Address = "2 Lake Road, Mozang Chungi, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99212830"
                                    SagField[2] = Address.strip()

                                elif Department == "BUFFALO RESEARCH INSTITUTE, PATTOKI DISTRICT KASUR":
                                    Address = "National Hwy 5, Kasur, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 300 6601560"
                                    SagField[2] = Address.strip()

                                elif Department == "GOVERNMENT SADIQ COLLEGE WOMEN UNIVERSITY BAHAWALPUR":
                                    Address = "Govt sadiq college women University Girls College Rd, Anwar Colony, Bahawalpur, Punjab 63100, Pakistan" + "<br>\n" + "Phone: " + "+92 62 9250517"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "soil.punjab.gov.pk"
                                    SagField[1] = "soil.punjab@yahoo.com"

                                elif Department == "GOVERNMENT SADIQ COLLEGE WOMEN UNIVERSITY BAHAWALPUR":
                                    Address = "Govt sadiq college women University Girls College Rd, Anwar Colony, Bahawalpur, Punjab 63100, Pakistan" + "<br>\n" + "Phone: " + "+92 62 9250517"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.gscwu.edu.pk"
                                    SagField[1] = "INFO@GSCWU.EDU.PK,WEBMASTER@GSCWU.EDU.PK"

                                elif Department == "CH. PERVAIZ ELAHI INSTITUTE OF CARDIOLOGY, MULTAN":
                                    Address = "Abdali Rd, Tipu Sultan Colony, Dera Adda, Multan, Punjab 60000, Pakistan" + "<br>\n" + "Phone: " + "+92 61 9201045"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "cpeic.gop.pk"
                                    SagField[1] = "cpeic.mic@gmail.com"

                                elif Department == "AGRICULTURE DELIVERY UNIT, LAHORE":
                                    Address = "Agriculture Delivery Unit (ADU), Agriculture Department Punjab, 21 Davis Road, Garhi Shahu, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + " +92 42 99205549"
                                    SagField[2] = Address.strip()

                                elif Department == "SOIL AND WATER TESTING LABORATORY, LAHORE":
                                    Address = "Near D.H.Q Hospital, Attock, Pakistan" + "<br>\n" + "Phone: " + "057-9316135"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.agripunjab.gov.pk"
                                    SagField[1] = "swtlatk@gmail.com"

                                elif Department == "MUNICIPAL COMMITTEE, JALALPUR JATTAN":
                                    Address = "Jalalpur Jattan, Gujrat, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "SENIOR CIVIL JUDGE (CIVIL DIVISION), LODHRAN":
                                    Address = "New Judicial Complex, Chak No. 100/M, Lodhran" + "<br>\n" + "Phone: " + "0608-921009" + "<br>\n" + "Fax: " + "0608-921008"
                                    SagField[2] = Address.strip()

                                elif Department == "THQ HOSPITAL, CHUNIAN":
                                    Address = "Chunian, Kasur, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 49 4311085"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT JAIL, PAKPATTAN":
                                    Address = "Pakpattan-Kamir Rd, Chak 36SP, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 457 380240"
                                    SagField[2] = Address.strip()

                                elif Department == "PROJECT MANAGEMENT UNIT, SPORTS BOARD PUNJAB, LAHORE":
                                    Address = "National Hockey Stadium, Ferozepur Road Lahore." + "<br>\n" + "Phone: " + "+92-42-99232518,+92-42-99230855"
                                    SagField[2] = Address.strip()

                                elif Department == "DHQ HOSPITAL, PAKPATTAN":
                                    Address = "Pākpattan, Pakpattan, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 457 373486"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "sportsboard.punjab.gov.pk"
                                    SagField[1] = "it@sportsboard.punjab.gov.pk"

                                elif Department == "DIVISIONAL FOREST OFFICER, KASUR":
                                    Address = "Kasur Forest Division Changa Manga," + "<br>\n" + "Phone: " + "(049) 4381561"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "dfoliaqatgill@gmail.com"

                                elif Department == "DIVISIONAL FOREST OFFICER, KASUR":
                                    Address = "Shahra-e-Quaid-e-Azam,Lahore, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99212951-66"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.lhc.gov.pk"

                                elif Department == "MUNICIPAL COMMITTEE CHINIOT":
                                    Address = "Ghazi Cotton Link Rd, Jhang, Chiniot, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.mcchiniot.lgpunjab.org.pk"

                                elif Department == "ON FARM WATER MANAGEMENT WING OF AGRICULTURE":
                                    Address = "Agriculture House, 1st Floor, 21-Davis Road, Lahore. Pakistan." + "<br>\n" + "Phone: " + "(042) 99200703, 99200713" + "<br>\n" + "Fax: " + "(042) 99200702, 99200710"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "ofwm.agripunjab.gov.pk"
                                    SagField[1] = "pipipwm@gmail.com"

                                elif Department == "EXECUTIVE ENGINEER HIGHWAYS DIVISION, SIALKOT.":
                                    Address = "SIALKOT, Pakistan" + "<br>\n" + "Tel: " + "0304-0920065, 9250451-52"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "dcosialkot@punjab.gov.pk"

                                elif Department == "UNIVERSITY OF HEALTH SCIENCES, LAHORE":
                                    Address = "Khayaban-e-Jamia Punjab, Block D Muslim Town, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 111 333 366"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "info@uhs.edu.pk"
                                    SagField[8] = "www.uhs.edu.pk"

                                elif Department == "CARDIAC CENTRE, CHUNIAN":
                                    Address = "Allahabad - Chunian Rd, Zaheerabad, Chunian, Kasur, Punjab 55220, Pakistan" + "<br>\n" + "Phone: " + "+92 49 4310015"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT SESSIONS COURT TOBA TEK SINGH":
                                    Address = "Toba Tek Singh, Toba Tek Singh District, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 46 2515851"
                                    SagField[2] = Address.strip()

                                elif Department == "COMMUNICATION AND WORKS":
                                    Address = "DHQ Complex, 'H' Block , Old Secretariat Muzaffarabad, Pakistan" + "<br>\n" + "Phone: " + "+(05822)-921410"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "info@cwd.gok.pk"
                                    SagField[8] = "cwd.gok.pk"

                                elif Department == "PROVINCIAL BUILDINGS DIVISION GOR, LAHORE":
                                    Address = "Shadman, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+(05822)-921410"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF ENGINEERING AND TECHNOLOGY TAXILA":
                                    Address = "UET, Taxila, Rawalpindi, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 51 9047500"
                                    SagField[2] = Address.strip()

                                elif Department == "MUNICIPAL COMMITTEE, PAKPATTAN":
                                    Address = "Pākpattan, Pakpattan, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 457 372415"
                                    SagField[2] = Address.strip()

                                elif Department == "CATTLE MARKET MANAGEMENT COMPANY, SAHIWAL DIVISION":
                                    Address = "District, Punjab, Ganj Shakar Colony Sahiwal, Sahiwal District, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 40 4270022"
                                    SagField[2] = Address.strip()

                                elif Department == "MUNICIPAL CORPORATION, SARGODHA":
                                    Address = "Kitchen Road, Sargodha, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 48 9230148"
                                    SagField[2] = Address.strip()

                                elif Department == "IRRIGATION AND POWER DEPARTMENT":
                                    Address = "Punjab Irrigation Department, Old Anarkali, Lahore, Pakistan" + "<br>\n" + "Phone: " + "(042) 99213595 - 7"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT HEALTH AUTHORITY, GUJRAT":
                                    Address = "EDOs Complex, DCO Office Gujrat, Pakistan" + "<br>\n" + "Phone: " + "+92 53 9260105"
                                    SagField[2] = Address.strip()

                                elif Department == "MUNICIPAL CORPORATION, SAHIWAL":
                                    Address = "Municipal Corporation Sahiwal ( Dulla Bhatti Chowk ) Pakistan" + "<br>\n" + "Tel: " + "068-58012145" + "<br>\n" + "Fax: " + "068-5801224"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.mcsahiwal.lgpunjab.org.pk"

                                elif Department == "GOVT. COLLEGE OF TECHNOLOGY KAMALIA.":
                                    Address = "Near GCC Kamalia.، Kamalia, Toba Tek Singh District, Punjab 36350, Pakistan" + "<br>\n" + "Phone: " + "+92 46 3411025"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT POLICE OFFICE, ATTOCK":
                                    Address = "Attock, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 57 9316353"
                                    SagField[2] = Address.strip()

                                elif Department == "UNIVERSITY OF ENGINEERING AND TECHNOLOGY, FAISALABAD CAMPUS":
                                    Address = "Khurianwala Bypass, Khurrian Wala, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 41 4360006"
                                    SagField[2] = Address.strip()

                                elif Department == "SERVICES HOSPITAL, JAIL ROAD, LAHORE":
                                    Address = "Ghaus-ul-Azam, Shadman Jail Road, Punjab، Lahore, 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99203402"
                                    SagField[2] = Address.strip()

                                elif Department == "PROVINCIAL DISASTER MANAGEMENT AUTHORITY, PUNJAB":
                                    Address = "40-A، Lawrence Road, Jubilee Town, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99203301"
                                    SagField[2] = Address.strip()

                                elif Department == "DISTRICT POLICE OFFICE, BAHAWALPUR":
                                    Address = "Farid Gate Rd, Bahawalpur Cantt, Bahawalpur, Punjab 63100, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "TEVTA DISTRICT OFFICE GUJRANWALA AND HAFIZABAD":
                                    Address = "TEVTA Secretariat, 96-H, Gulberg Road, Lahore." + "<br>\n" + "Phone: " + "+92 -42 -99263055-59"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "info@tevta.gop.pk"
                                    SagField[8] = "www.tevta.gop.pk"

                                elif Department == "D. I. G. SECURITY DIVISION, LAHORE":
                                    Address = "Inchrage Monitoring Branch, EPF HQrs, Pakistan" + "<br>\n" + "Phone: " + "042-99205044"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "epfmonhq@gmail.com"

                                elif Department == "GOVT. COLLEGE FOR WOMEN, ATTOCK":
                                    Address = "Attock, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 57 9316164"
                                    SagField[2] = Address.strip()

                                elif Department == "POLICE WIRELESS TRAINING SCHOOL, BAHAWALPUR":
                                    Address = "BAHAWALPUR, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "LOCAL GOVERNMENT AND COMMUNITY DEVELOPMENT LAHORE":
                                    Address = "Anarkali Bazaar, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99210013"
                                    SagField[2] = Address.strip()

                                elif Department == "MUNICIPAL COMMITTEE, KAMOKE":
                                    Address = "Main G.T.Road Near Rajbah, Kamoke, Pakistan" + "<br>\n" + "Tel: " + "055-6811542"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "mckamoki@gmail.com"

                                elif Department == "SOIL AND WATER TESTING LABORATORY, SAHIWAL":
                                    Address = "Sahiwal, Sahiwal District, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 40 4400658"
                                    SagField[2] = Address.strip()

                                elif Department == "RAWALPINDI MEDICAL COLLEGE, RAWALPINDI.":
                                    Address = "Tipu Rd, Chamanzar Colony, Rawalpindi, Punjab 46000, Pakistan" + "<br>\n" + "Phone: " + "+92 51 9281011"
                                    SagField[2] = Address.strip()

                                elif Department == "THQ HOSPITAL, SARAI ALAMGIR":
                                    Address = "Mandi Bahauddin - Sarai Alamgir Rd, Mandi Bahauddin, Gujrat, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 544 654242"
                                    SagField[2] = Address.strip()

                                elif Department == "PROVINCIAL HIGHWAY CIRCLE, D.G.KHAN":
                                    Address = "Jampur Rd, Khayaban-e-Sarwar, Dera Ghazi Khan, Punjab, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()

                                elif Department == "CENTRAL JAIL, LAHORE":
                                    Address = "Central Jail Kot Lakhpat Lahore, Pakistan" + "<br>\n" + "Phone: " + "(042) 35400674" + "<br>\n" + "Fax: " + "(042) 35215685"
                                    SagField[2] = Address.strip()
                                    SagField[1] = "	cj.lahore@gmail.com"

                                elif Department == "BOARD OF INTERMEDIATE SECONDARY EDUCATION, FAISALABAD":
                                    Address = "Jhang Rd, Air Avenue City, Faisalabad, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 412517705/6"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.bisefsd.edu.pk"

                                elif Department == "PUNJAB HIGHWAY DEPARTMENT, HIGHWAYS DIVISION, RAHIM YAR KHAN":
                                    Address = "Rahim Yar Khan, Punjab 64200, Pakistan" + "<br>\n" + "Phone: " + "+92 68 9330005"
                                    SagField[2] = Address.strip()

                                elif Department == "DAANISH SCHOOL (BOYS) HARNOLI, MIANWALI":
                                    Address = "Mianwali Muzaffargarh Road, Mianwali, Punjab, Pakistan" + "<br>\n" + "Phone: " + "0300 9666261"
                                    SagField[2] = Address.strip()

                                elif Department == "THQ HOSPITAL PIPLAN":
                                    Address = "Hafiz Wala Road, Piplan, Mianwali, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 459 201143"
                                    SagField[2] = Address.strip()

                                elif Department == "QUAID-E-AZAM ACADEMY FOR EDUCATIONAL DEVELOPMENT, CHISHTIAN, BAHAWALNAGAR":
                                    Address = "Wahdat Colony, Wahdat Road,LAHORE, Pakistan" + "<br>\n" + "Phone: " + "+92 042 99260108"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "qaed.edu.pk"
                                    SagField[1] = "info@qaed.edu.pk"

                                elif Department == "UNIVERSITY OF EDUCATION":
                                    Address = "College Road, Block C Phase 1 Township, Lahore, Punjab 54770, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99262231"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.ue.edu.pk"
                                    SagField[1] = "vc@ue.edu.pk"

                                elif Department == "MUNICIPAL CORPORATION, SIALKOT":
                                    Address = "Sialkot, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 347 4976562"
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB LAND RECORDS AUTHORITY, LAHORE":
                                    Address = "Punjab Land Records Authority - PLRA Government of the Punjab 2-Kilometer Main Multan Road, Opposite EME-DHA Housing Society, Lahore" + "<br>\n" + "Phone: " + "(042) 99330111-112"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.punjab-zameen.gov.pk"
                                    SagField[1] = "info@punjab-zameen.gov.pk"

                                elif Department == "LAHORE GENERAL HOSPITAL, LAHORE":
                                    Address = "Ferozpur Road، near Chungi، Amar Sidhu Ismail Nagar, Lahore, Punjab 54000, Pakistan" + "<br>\n" + "Phone: " + "+92 42 99268801"
                                    SagField[2] = Address.strip()

                                elif Department == "PUNJAB SKILLS DEVELOPMENT FUND":
                                    Address = "Dr Mateen Fatima Rd, Block H Gulberg 2, Lahore, Punjab, Pakistan" + "<br>\n" + "Phone: " + "+92 800 486 27"
                                    SagField[2] = Address.strip()
                                    SagField[8] = "www.psdf.org.pk"
                                    SagField[1] = "Info@psdf.org.pk"

                                elif Department == "DISTRICT OFFICER BUILDINGS, CHAKWAL":
                                    Address = "Chakwal, Pakistan" + "<br>\n" + "Phone: " + ""
                                    SagField[2] = Address.strip()
                                else:
                                    SagField[2] = "Punjab, Pakistan\n<br>[Disclaimer : For Exact Organization/Tendering Authority details, please refer the tender notice.]"

                                SagField[12] = Department.strip()

                            Procurement_Title = ""
                            Type = ""
                            Publish_Date = ""
                            Department = ""
                            Status = ""

                            for Procurement_Title in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[1]'):
                                Procurement_Title = Procurement_Title.get_attribute("innerText").strip()
                                Procurement_Title = string.capwords(str(Procurement_Title))
                                Procurement_Title = html.unescape(str(Procurement_Title))

                            for Procurement_Name in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[2]'):
                                Procurement_Name = Procurement_Name.get_attribute("innerText").replace(
                                    "View Tender Detail", "")
                                Procurement_Name = string.capwords(str(Procurement_Name))
                                Procurement_Name = html.unescape(str(Procurement_Name))
                                SagField[19] = Procurement_Name.strip()

                            for Type in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[3]'):
                                Type = Type.get_attribute("innerText").strip()
                                Type = string.capwords(str(Type))
                                Type = html.unescape(str(Type))

                            for Publish_Date in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[4]'):
                                Publish_Date = Publish_Date.get_attribute("innerText").strip()

                            for Close_Date in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[5]'):
                                Close_Date = Close_Date.get_attribute("innerText")
                                datetime_object = datetime.strptime(Close_Date, '%d %b %Y')
                                Close_Date1 = datetime_object.strftime("%Y-%m-%d")
                                SagField[24] = Close_Date1.strip()

                            for Status in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[7]'):
                                Status = Status.get_attribute("innerText").replace("&nbsp;", "").strip()
                                Status = string.capwords(str(Status))
                                Status = html.unescape(str(Status))

                            for Tender_Notice in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[8]/a'):
                                Tender_Notice = Tender_Notice.get_attribute("href")
                                SagField[4] = Tender_Notice

                            for Bidding_Document in browser.find_elements_by_xpath(
                                    '/html/body/form/div[3]/div[2]/div/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr[5]/td/div/table/tbody/tr[' + str(
                                            add) + ']/td[9]/a'):
                                Bidding_Document = Bidding_Document.get_attribute("href")
                                SagField[5] = Bidding_Document

                            Tender_Details = "Title: " + str(SagField[19]) + "<br>\n""Procurement Title: " + str(
                                Procurement_Title) + "<br>\n""Type: " + str(Type) + "<br>\n""Publish Date: " + str(
                                Publish_Date) + "<br>\n""Department: " + str(Department) \
                                             + "<br>\n""Status: " + str(Status)
                            SagField[18] = Tender_Details.strip()

                            SagField[7] = "PK"
                            SagField[20] = ""
                            SagField[21] = ""
                            SagField[22] = "0.0"
                            SagField[26] = "0.0"
                            SagField[27] = "0"  # Financier
                            SagField[28] = "https://eproc.punjab.gov.pk/ActiveTenders.aspx"
                            SagField[31] = "eproc.punjab.gov.pk"
                            SagField[42] = SagField[7]
                            SagField[43] = ''
                            
                            for SegIndex in range(len(SagField)):
                                print(SegIndex, end=' ')
                                print(SagField[SegIndex])
                                SagField[SegIndex] = html.unescape(str(SagField[SegIndex]))
                                SagField[SegIndex] = str(SagField[SegIndex]).replace("'", "''")
                            check_date(SagField)
                            print(" Total: " + str(Global_var.Total) + " Duplicate: " + str(
                                        Global_var.duplicate) + " Expired: " + str(
                                        Global_var.expired) + " Inserted: " + str(
                                        Global_var.inserted) + " Skipped: " + str(
                                        Global_var.skipped) + " Deadline Not given: " + str(
                                        Global_var.deadline_Not_given) + " QC Tenders: " + str(Global_var.QC_Tenders),"\n")
                            v = True
                        except Exception as b:
                            exc_type, exc_obj, exc_tb = sys.exc_info()
                            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                            print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(b), "\n", exc_type, "\n",fname, "\n", exc_tb.tb_lineno)
                            v = False
                for Next_pages in browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolderSRIS_rdgrdManageTender_ctl00"]/tfoot/tr/td/table/tbody/tr/td/div[3]/input[1]'):
                    Next_pages.click()
                    time.sleep(2)
                    break

            a = False
        except Exception as e:
            a = True
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(e), "\n", exc_type, "\n", fname, "\n", exc_tb.tb_lineno)


def check_date(SagField):
    tender_date = str(SagField[24])
    nowdate = datetime.now()
    date2 = nowdate.strftime("%Y-%m-%d")
    try:
        if tender_date != '':
            deadline = time.strptime(tender_date , "%Y-%m-%d")
            currentdate = time.strptime(date2 , "%Y-%m-%d")
            if deadline > currentdate:
                insert_in_Local(SagField)
            else:
                print("Tender Expired")
                Global_var.expired += 1
        else:
            print("l")
            Global_var.skipped += 1
            Global_var.deadline_Not_given += 1
    except Exception as e:
        exc_type , exc_obj , exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname , "\n" ,exc_tb.tb_lineno)


ChromeDriver()
