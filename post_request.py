import requests
from bs4 import BeautifulSoup as bs
import re
import xlsxwriter
import tkinter
import tkinter.filedialog
import time
import datetime
def get_pages(stateId):
    headers = {
            "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "Accept-Encoding":"gzip, deflate, br",
            "Accept-Language":"en-US,en;q=0.9",
            "Cache-Control":"max-age=0",
            "Connection":"keep-alive",
            "Content-Length":"646",
            "Content-Type":"application/x-www-form-urlencoded",
            "Cookie":"BNI_persistence=6UwsW5Ik0TJvBoF7ogAktLGHclNmXIhVacPamy5USSZ134oMI4K6rsr793GzcjyoqtiUc5cKV2-9ZcOfgA4pRA==; ASP.NET_SessionId=kwtdrf5vfkkqwphnrm5tisbt; BNES_ASP.NET_SessionId=dLcKfI9njoeK5milXI5ebbBicgD/d89VPsUmYaPJNA4yM6tjgM+/AB0xThcye2Eb041+v7SbUGgBIkRztLQqDH+BG29YvREajjhoC5kSrD3m9gt4nckGjw==; __RequestVerificationToken_L2Fkcw2=fDvL4biKCDz0ShbbyNVHuU1Xe2JpoNNSCa1VXnRklRnvfE4lByZhAfafmRh199IMFe6A5QrAlxiZv6mzPzVIPz7CApY1; BNES___RequestVerificationToken_L2Fkcw2=nzbTK9XzeZ4GV4ifKKOFvyhUGvYn5zRLJxnhN4SkoLlZbxQkTsDc9E6LdckYeKJbkgK8qO6rbMuan7PLzkV6XUT9XMaP702sbAXfokC7Roo2Ug3ZMHvhTuG3PVqDZsZ19YuR1uSvob4Nsnr+7zbmHxmjcA5HU32ePD87yy8Bj+Lh/tms/S+0fmCpp5Foy+KfFI0HEZY5voa82VTK+aEEZr1JDa8PQwMlhP9l9gzlyew=",
            "Host":"apps.acgme.org",
            "Origin":"https://aps.acgme.org",
            "Referer":"https://apps.acgme.org/ads/Public/Programs/Search",
            "Sec-Fetch-Dest":"document",
            "Sec-Fetch-Mode":"navigate",
            "Sec-Fetch-Site":"same-origin",
            "Sec-Fetch-User":"?1",
            "Upgrade-Insecure-Requests":"1",
            "User-Agent":"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36"
        }


    payload = {"stateId":stateId,
               "specialtyId":"",
               "specialtyCategoryTypeId":"",
               "numCode":"",
               "city":"",
               "g-recaptcha-response":"03AGdBq25Q8u8QuKcu3owEghbVLjqu0p-_fSBf_KNiBKf_5ws5pDUeRc1TNnyYLepUaUP9palO1Ekj8TPTU6BuM91wwjsHsxi_ryvegPPdFIp-vDQtQbjMaxQIXxhTpC4B_6zftn8xCNW0ZpBfDUcaGxaFnivWfCerhlDct-5IUypC8e42zj-2DXS8izWq2qk9Y4yq6kXp1Q-P85YAXAUFa34ecbEpbx0bDSMhTPDfERnC2Arbg51-jlJQNo2S8qYd4Q-Su1crv9xHrKWdGT_AySjbFgtJ_vabc5h6QwVfMv03CFVoawb3qCYqDYC-Jgn6a0_n1e6f4IMGdm9Lh_oe3rQfgBIawVwt88kWjcERTHeyGyQ9NSfzaPrm9a0kANjNrSd1b-eKn5YK54YzY8_t1H-WI4SwvLbXV9w8gPLmY7AVyLwQH4_Bcy8",
               "__RequestVerificationToken":"amvVIUkWU0a3-N0-8mpeyZsHKLXQiuEqELFvYa_Ccdw_jwJpTPfA9tD3ZQXkSDtNO_5p9cn6LuOpyJd6xsMKf2PHs_s1"
               }
    url = "https://apps.acgme.org/ads/Public/Programs/Search"
    r=requests.post(url,data=payload,headers=headers)

    soup = bs(r.content,'html5lib')

    table = soup.find(
            'table', attrs={'id': 'programsListView-listview'}).find('tbody')
    #print(table)
    row_urls = list(table.findAll('a'))
    row_urls = list(filter(
        lambda x: '/ads/Public/Programs/Detail?programId=' in x['href'], row_urls))
    urls = ['{}{}'.format("https://apps.acgme.org", row['href']) for row in row_urls]
    return urls

def add_to_excel(sheet1,y,program_number, institution, specialty, location, city, state, zip_code, address, website,code,phone,email):
    sheet1.write('A' + str(y),specialty.lstrip())
    sheet1.write('B' + str(y), program_number)
    # print(institution)
    sheet1.write('C' + str(y), " ".join(institution))
    sheet1.write('D' + str(y), str(code))
    sheet1.write('E' + str(y), city.lstrip())
    sheet1.write('F' + str(y), state)
    sheet1.write('G' + str(y), zip_code)
    sheet1.write('H' + str(y), location.lstrip())
    sheet1.write('I' + str(y), address)
    sheet1.write('J' + str(y), website)
    sheet1.write('K' + str(y), phone)
    sheet1.write('L' + str(y), email)
    return y + 1

def scrape_program(url,sheet1,y, stateId):
    headers = {
            "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "Accept-Encoding":"gzip, deflate, br",
            "Accept-Language":"en-US,en;q=0.9",
            "Connection":"keep-alive",
            "Content-Type":"application/x-www-form-urlencoded",
            "Cookie":"BNI_persistence=6UwsW5Ik0TJvBoF7ogAktLGHclNmXIhVL5bkrQKczHKnxVNTljqyLXatFauQnq4dveDlbEmDZTQrv_jUI_NCDg==; ASP.NET_SessionId=rm1adcgqimpxmhgxxn52vx0b; BNES_ASP.NET_SessionId=IrTCpfNUt9DoJl7GMQPwkkvPFfOlw1YcKW3aFzHGHpaOdfY3fPo+LmvAYSGiWw7oDMvv6Cdi5kw/sH5Gzl6KEVknN3oZ9UHB/3WrLV4J0QPVTTArGZd+fQ==; __RequestVerificationToken_L2Fkcw2=UzY0wf1KFaQod1hqxkduX1eGASwbRe1rtdLiGKIzj5B6rGdye2HkmqVK8mjDbJg0wHj7JtBts5cz0Zx1qRSup8rAYzE1; BNES___RequestVerificationToken_L2Fkcw2=sG9Fa+yr2eEueV42gepERYz/+o+adR5zxfC6jLo14yIDExYYkzIt3GafEeCzMp4+qfjIK8SeQEDQ3gfHqvIKh6huElfcW3SmLMkiNSq/Jnu9oGDVB+0mjaVv+rNrfRQh4wLmLoyyJZ6AdgGUMhpgIiuNNvcCNzmOoA6MIhvWkHsQtCV0fN5tZ2HTnaarHyP12isNyg4b8Ob7mD0Fo7LlFAogN7g9W4k0Oias7ZboNhQ=",
            "Host":"apps.acgme.org",
            "Referer":"https://apps.acgme.org/ads/Public/Programs/Search",
            "Sec-Fetch-Dest":"document",
            "Sec-Fetch-Mode":"navigate",
            "Sec-Fetch-Site":"same-origin",
            "Sec-Fetch-User":"?1",
            "Upgrade-Insecure-Requests":"1",
            "User-Agent":"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36"
        }
    x = url + "&ReturnUrl=https%3A%2F%2Fapps.acgme.org%2Fads%2FPublic%2FPrograms%2FSearch"
    timeStart = datetime.datetime.now()

    page1 = requests.get(x, headers=headers)
    timeEnd = datetime.datetime.now()
    print(timeEnd-timeStart)

    soup = bs(page1.content,'html5lib')
   # print(soup)
    h1 = soup.find('h1').getText().split('\n')
    program_number = h1[0].split(' ')[0]
    institution = h1[0].split(' ')[2:]
    specialty = " ".join(h1[2].split(' ')[0:(len(h1[2].split(' ')) - 1)])
    location = h1[3]
    print(h1)
    address = soup.find('address').getText().split('\n')
    city=""
    state=""
    zip_code = ""
    try:
        city = address[len(address) - 2].split(", ")[0]
    except:
        city = "No city found"
    # print(city.lstrip())
    # exit()
    try:
        state = address[len(address) - 2].split(",")[1][1:3]
    except:
        state = "State ID: " + str(stateId)
    try:
        zip_code = address[len(address) - 2].split(",")[1][4:]
    except:
        zip_code = "No zip code found"
    for x in range(0,len(address)):
        address[x] = address[x].lstrip()
    address = "|".join(address[1:len(address)-1])

    code = int(
        soup.find('dl', attrs={'class': "dl-horizontal"}).find('dd').find_next('a')
            .contents[0].strip())
    # print(code)
    phone = ''
    email = ''
    try:
        phone = list(soup.findAll('dl', attrs={'class': "dl-horizontal"}))[2].find('dd').contents[0].strip()
    except:
        phone = 'Could not find!'
    try:
        email = list(soup.findAll('dl', attrs={'class': "dl-horizontal"}))[2].find('dd').find_next('a').contents[0].strip()
    except:
        email = 'Could not find!'
    print(phone)
    print(email)
    website = soup.findAll('a')[2].get('href')
    
    y = add_to_excel(sheet1,y,program_number, institution, specialty, location, city, state, zip_code, address, website,code,phone,email)
    return y 

print("Setting up excel sheet")
finput = tkinter.filedialog.askopenfilename()
f = xlsxwriter.Workbook(finput)
sheet1 = f.add_worksheet()
print("Finished! Scraping pages..")
sheet1.write('A1', 'Specialty')
sheet1.write('B1', 'Program number')
sheet1.write('C1', 'Organization')
sheet1.write('D1', 'Code')
sheet1.write('E1', 'City')
sheet1.write('F1', 'State')
sheet1.write('G1', 'Zip Code')
sheet1.write('H1', 'City/State')
sheet1.write('I1', 'Full Address')
sheet1.write('J1', 'Website')
sheet1.write('K1', 'Phone')
sheet1.write('L1', 'Email')
sheet1.write('M1', 'Email')
sheet1.write('N1', 'Director')
sheet1.write('O1', 'Director Appointment Date')
sheet1.write('P1', 'Cordinator')
sheet1.write('Q1', 'Cordinator Phone Number')
sheet1.write('R1', 'Cordinator Email')
sheet1.write('S1', 'Osteopathic Recognition')
y = 2

for x in range(1,53):
    # print(x)
    print("Fetching pages from the ACGME Program search")
    pages = get_pages(x)
    print("Finished! Setting up excel sheet...")
    for x in range(0,len(pages)):
        if x % 10 == 0:
            print("Scraping page " + str(x) + " of " + str(len(pages)))
        timeStart = datetime.datetime.now()
        # time.sleep(30)
        try:
            y = scrape_program(pages[x], sheet1,y,x)
        except:
            f.close()
            exit()
        timeEnd = datetime.datetime.now()
        # print(timeEnd-timeStart)
f.close()
