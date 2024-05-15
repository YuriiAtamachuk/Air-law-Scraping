import openpyxl.styles
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import Select
from openpyxl.styles import Alignment, Font
from time import sleep
from bs4 import BeautifulSoup
import fontstyle
from docx import Document
import re

def wait_url(driver: webdriver.Chrome, url: str):
    while True:
        cur_url = driver.current_url
        if cur_url == url:
            break
        sleep(0.1)

def find_element(driver: webdriver.Chrome, whichBy, unique: str) -> WebElement:
    while True:
        try:
            element = driver.find_element(whichBy, unique)
            break
        except:
            pass
        sleep(1)
    return element

def find_elements(driver : webdriver.Chrome, whichBy, unique: str) -> list[WebElement]:
    while True:
        try:
            elements =driver.find_elements(whichBy, unique)
            break
        except:
            pass
        sleep(1)
    return elements

remote_debugging_address = "127.0.0.1:9024"
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("debuggerAddress", remote_debugging_address)
driver = webdriver.Chrome(options=chrome_options)

url ="https://app.atplquestions.com/test"

email = "kakou324@gmail.com"
password = "Upwork123@"
match_num = 0


# driver.get(url)
wait_url(driver, url)
sleep(1)

#  save to docx
doc = Document()
table = doc.add_table(rows=1, cols=8)

header = table.rows[0].cells
header[0].text = "Question number"
header[1].text = "Question"
header[2].text = "Option1"
header[3].text = "Option2"
header[4].text = "Option3"
header[5].text = "Option4"
header[6].text = "Answer"
header[7].text = "Explanation"


questions_num = find_element(driver, By.CLASS_NAME, "maxSize")
questions_num_int = int(questions_num.text)
print(questions_num_int)

start_study = find_element(driver, By.CLASS_NAME, "startStudy")
start_study.click()
print("Started!")


file_name = find_element(driver, By.CLASS_NAME, "testname").text
print(file_name)
parts = file_name.split('|')
# Extract the second part (index 1) and split it by space
code_parts = parts[1].strip().split()
# The code is the first part of the second split
code = code_parts[0]
print(code)

data = []

for _ in range(questions_num_int):
    each_ques_data = []
    sleep(1)
    ques_mum =find_element(driver, By.CLASS_NAME, "q-number").text
    each_ques_data.append(ques_mum)
    quesiton = find_element(driver, By.CLASS_NAME, "questionText").text
    each_ques_data.append(quesiton)
    quesiton_options = find_elements(driver, By.CLASS_NAME, "questionOption")
    for index, question_each_option in enumerate(quesiton_options):
        question_option = question_each_option.find_element(By.CLASS_NAME, "texter")
        if index == 0:
            option1 = question_option.text
            each_ques_data.append(option1)
        elif index == 1:
            option2= question_option.text
            each_ques_data.append(option2)
        elif index == 2:
            option3= question_option.text
            each_ques_data.append(option3)
        elif index == 3:
            option4= question_option.text
            each_ques_data.append(option4)

    user_navs = find_elements(driver, By.CLASS_NAME, "nav-link")
    for explanation in user_navs:
        if explanation.text == "Explanation":
            explanation.click()
            sleep(2)
    explanation_txt = find_element(driver, By.CLASS_NAME, "exp-text").get_attribute('innerHTML')
    print(explanation_txt)
    soup = BeautifulSoup(explanation_txt, 'html.parser')
    plain_text = ""
    for element in soup.recursiveChildGenerator():
        if isinstance(element, str):
            plain_text += element
        elif element.name == 'strong':
            plain_text += '<bold>' + element.text + '</bold>'
        elif element.name == 'u':
            plain_text += '<underline>' + element.text + '</underline>'
        elif element.name == 'em':
            plain_text += '<italic>' + element.text + '</italic>'
    # for element in soup.find_all():
    #     if element.name == 'strong':
    #         plain_text += fontstyle.apply(element.get_text(separator="\n"), 'bold')
    #     elif element.name == 'em':
    #         plain_text += fontstyle.apply(element.get_text(separator="\n"), 'italic')
    #     elif element.name == 'u':
    #         plain_text += fontstyle.apply(element.get_text(separator="\n"), 'underline')
    #     else:
    #         plain_text += element.get_text(separator="\n")
        # plain_text += '\n'
    cleand_text = re.sub(r'\x1b\[[0-9;]*m', '', plain_text)

    print(plain_text)
    each_ques_data.append("")
    each_ques_data.append(cleand_text)
    # sheet[f'H{match_num+1}'] = explanation_txt
    # sheet[f'H{match_num+1}'] = soup.get_text()
    # for element in soup.find_all():
    #     text = element.get_text()
    #     if element.name == 'span':
    #         if 'font-weight' in element.get('style', ''):
    #             sheet[f'H{match_num+1}'].font = Font(bold=True)
    #         if 'color' in element.get('style', ''):
    #             sheet[f'H{match_num+1}'].font = Font(color=element['style'].split('color:')[1].split(';')[0])
    #     # sheet[f'H{match_num+1}'].value = text
    find_element(driver, By.CLASS_NAME, "next").click()
    # print("each===========>",each_ques_data)
    data.append(each_ques_data)
    # workbook.save(f'{code}.xlsx')

print(data)
for row in data:
    row_cells = table.add_row().cells
    for index, cell_value in enumerate(row):
        row_cells[index].text = cell_value

doc.save(f'{code}.docx')

print("data====>", data)