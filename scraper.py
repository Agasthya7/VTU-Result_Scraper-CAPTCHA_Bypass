#Importing the necessary libraries
import cv2
import numpy as np
import pytesseract
import time
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from captcha_solver import captcha_og
counter=0
usn_input_style=input('Enter the way you want to enter the USN data.\nIf you have a text file containing USN PRESS 1 \n If you have your usn list specified in the USN_Data.txt file PRESS 2 \n If you are providing range of USN PRESS 3\n')

if usn_input_style == '1':
    while True:
        try:
            locn = input('Provide the location of text file without any quotes\n') 
            with open(locn, 'r') as f:
                my_list = [line.strip() for line in f]
            break
        except FileNotFoundError:
            print("Invalid file location. Please try again.")

elif usn_input_style =='2':
            with open(r"Data\USN_Data.txt", 'r') as f:
                my_list = [line.strip() for line in f]
            break

elif usn_input_style == '3':
    first_bit = input('Enter the first bit of your college usn\n')
    college = input("Enter the college code\n").upper()
    year = input('Enter the year\n')
    branch = input('Please enter the branch\n').upper()

    while True:
        try:
            low = int(input('Enter starting USN. If it starts from 001 enter it as 1\n'))
            high = int(input('Enter last USN excluding Lateral Entry USN\n'))
            lat_high = int(input('Enter last USN of lateral entry students (EX:417). If there are no lateral entry students enter 0\n'))
            break
        except ValueError:
            print("Invalid input. Please enter integers only.")

    my_list = []
    if branch not in ['CS', 'EC', 'ME', 'CV', 'IS']:
        print("Invalid branch code. Please try again.")
    elif not year.isdigit() or int(year) < 17 or int(year) > 23:
        print("Invalid year. Please enter an integer between 17 and 23.")
    else:
        for i in range (low, high+1):
            if i < 10:
                usn = first_bit + college + year + branch + '00' + str(i)
            elif i < 100:
                usn = first_bit + college + year + branch + '0' + str(i)
            else:
                usn = first_bit + college + year + branch + str(i)
            my_list.append(usn)
        if lat_high != 0:
            for i in range (400, lat_high+1):
                usn = first_bit + college + str(int(year)+1) + branch + str(i)
                my_list.append(usn)

    
#function to solve captcha
def solve_captcha(driver):
    div_element = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[2]/img')
    div_element.screenshot(r'Data\Captcha\unsolved.png')

    # load imge and set the bounds
    img = cv2.imread(r'Data\Captcha\unsolved.png')
    lower =(102, 102, 102) # lower bound for each channel
    upper = (125,125, 125) # upper bound for each channel

    # create the mask and use it to change the colors
    mask = cv2.inRange(img, lower, upper)
    img[mask != 0] = [0,0,0]

    # Save it
    cv2.imwrite(r'Data\Captcha\semisolved.png',img)

    img = Image.open(r'Data\Captcha\semisolved.png') # get image
    pixels = img.load() # create the pixel map

    for i in range(img.size[0]): # for every pixel:
        for j in range(img.size[1]):
            if pixels[i,j] != (0,0,0): # if not black:
                pixels[i,j] = (255, 255, 255) # change to white
    #img.show()
    img.save(r'Data\Captcha\solved.png')


    # read image
    img = cv2.imread(r'Data\Captcha\solved.png')

    # configurations
    config = ('-l eng --oem 1 --psm 3')

    # pytessercat
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    text = pytesseract.image_to_string(img, config=config)

    # print text
    tes_captcha = text.split('\n')[0]
    captcha=captcha_og()
    if len(captcha) < 6 :
        return tes_captcha

    return captcha

student_no=len(my_list)

for i in range(student_no):
    count=1
    usn_no=my_list[i]
    print("Currently trying to grab the results of "+usn_no)
    repeat=True
    while repeat:
        #repeat=False
        # configure web driver
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

# launch browser
        driver = webdriver.Chrome('C:\chromedriver\chromedriver.exe', options=options)

# load url
        url = 'https://results.vtu.ac.in/JJEcbcs23/resultpage.php' 
        driver.get(url)
        # wait for page to load
        time.sleep(2)

# solve captcha
        captcha=solve_captcha(driver)
# wait for result to load
        time.sleep(2)

# check if captcha is incorrect
    #captcha_text_field = driver.find_element_by_name('captchacode')
        if len(captcha) != 6:
    # if captcha is incorrect, click refresh button and solve again
            refresh_button = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[3]/p/a')
            refresh_button.click()
            time.sleep(2)
            captcha=solve_captcha(driver)
    
# enter usn
        usn_text_field = driver.find_element_by_name('lns')
        usn_text_field.send_keys(usn_no)# CHANGE THIS LATER ON FOR DEBUGGING
        captcha_text_field = driver.find_element_by_name('captchacode')
        captcha_text_field.send_keys(captcha)

# submit form
        submit_button = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[3]/div[1]/input')
        submit_button.click()
        
        try:
            alert = driver.switch_to.alert
            if alert.text == "University Seat Number is not available or Invalid..!":
                #print("No results for : " + usn_no +" going to the next USN")
                alert.accept()
                driver.quit()
                repeat = False # proceed to next USN
                break
            elif alert.text == "Invalid captcha code !!!":
                #print("Invalid CAPTCHA Detected for USN : " + usn_no +".Retrying for the same USN")
                count=count+1
                alert.accept()
                driver.quit()
                continue # repeat program for same USN again
        except:
            repeat = False # proceed to next USN
# wait for result to load
        #time.sleep(8)
        #element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]')))
        try:
            element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]')))
        except:
            # if timeout exception is raised, repeat the program for the same USN
            count += 1
            repeat=True
            driver.quit()
            continue # repeat program for same USN again

        stud_element = driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]')
        usn_element = driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]')
        table_element = driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div')
        sub_elements = table_element.find_elements_by_xpath('div')
        num_sub_elements = len(sub_elements)
        stud_text = stud_element.text
        usn_text = usn_element.text

        with open(r"Data\Results.txt", 'a') as file:
            for i in range (2,num_sub_elements+1):
                file.write(stud_text)
                file.write(',')
                file.write(usn_text)
                for j in range (1,7):
                    data=driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div['+str(i)+']/div['+str(j)+']')
                    data_text=data.text
                    file.write(',')
                    file.write(data_text)
                file.write('\n')
            file.write('\n')                
        print('Ran the program for '+str(count)+' times for this USN')
        counter=counter+count
        with open(r"Data\Completed.txt", 'a') as file:
            file.write(usn_no)
            file.write('\n')
        file_path = r'Data\USN_Data.txt'
        line_to_delete = usn_no+'\n' # Replace with the line you want to delete

        # Read the content of the file and store it in a list
        with open(file_path, 'r') as file:
            lines = file.readlines()
    
        # Remove the line you want to delete from the list
        lines = [line for line in lines if line != line_to_delete]

        # Write the modified list back to the file
        with open(file_path, 'w') as file:
            file.writelines(lines)
    # close browser
    driver.quit()
print("Ran for a Total of "+str(counter)+" times")
