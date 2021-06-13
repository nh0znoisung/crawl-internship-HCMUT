from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep
from selenium.webdriver.common.keys import Keys
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import timeit
import progressbar
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import os
import requests
import shutil
from pathlib import Path
from docx2pdf import convert
from os import walk
########################################
options = webdriver.ChromeOptions()

dir_link = os.getcwd() + "\\Internship\\Info"

preferences = {"download.default_directory": dir_link, "safebrowsing.enabled": "false"}
options.add_experimental_option("prefs", preferences)
browser = webdriver.Chrome(executable_path="./chromedriver", options=options)
browser.maximize_window()
url = 'https://internship.cse.hcmut.edu.vn/internship'
browser.get(url)


def dl_img(img_url, folder, file_name):
    Path(folder).mkdir(parents=True, exist_ok=True)

    response = requests.request("GET", img_url, stream = True)
    response.raw.decode_content = True
    file_img = open(folder+'/'+file_name, "wb")
    shutil.copyfileobj(response.raw,file_img)

def convert_word_to_pdf(dir):
	(_,_,filenames) = list(walk(dir))[0]

	# print(filenames)
	for file in filenames:
		if file.split(".")[-1] == "docx" :
			convert(dir + "\\" + file, dir + "\\" + "{}.pdf".format(file.split(".")[0]))
			os.remove(dir + "\\" + "{}.docx".format(file.split(".")[0]))

data_sponsor_url = []

data_id = []
data_name = []
data_introduction = []
data_field = []
data_max_register = []
data_max_stu_accept = []
data_stu_registed = []
data_stu_accepted = []
data_location = []
data_email = []

def main(dl_images = True, dl_files = True):
	try:
		WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.TAG_NAME, "figure"))) #since wifi, so we have a little bit delay
		start = timeit.default_timer()
		print("##################### CRAWL ALL COMPANY INTERNSHIP #######################")
		comp = browser.find_elements_by_tag_name("figure")
		n = len(comp)
		# print(n)
		idx = 0
		bar = progressbar.ProgressBar(maxval= n, widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()])
		bar.start()
		#### DEAL WITH NORMAL COMPANY
		for i in range(n):
			data_id.append(comp[i].get_attribute("data-id"))

			if comp[i].find_element_by_xpath('..').find_element_by_xpath('..').tag_name == 'a':
				# a => href
				link_href = comp[i].find_element_by_xpath('..').find_element_by_xpath('..').get_attribute("href")
				data_sponsor_url.append(link_href)
			else:
				idx += 1
				bar.update(idx)
				#click and def 
				sleep(3)
				comp[i].click() 

				turn_off = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[3]/button")
				sleep(3)
				# Name of company

				name = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/h4").text
				# print(name)
				data_name.append(name)

				# Introduction 
				intro = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/p[1]").text
				# print(intro)
				if intro == "":
					data_introduction.append("NULL")
				else:
					data_introduction.append(intro)

				# Field
				field = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/p[2]").text
				# print(field)
				if field == "":
					data_field.append("NULL")
				else:
					data_field.append(field)

				# Max register
				max_register = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/div[1]/div[1]/h6/span").text
				# print(max_register)
				if max_register == "":
					data_max_register.append("NULL")
				else:
					data_max_register.append(max_register)

				# Max accept
				max_stu_accept = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/div[1]/div[2]/h6/span").text
				# print(max_stu_accept)
				if max_stu_accept == "":
					data_max_stu_accept.append("NULL")
				else:
					data_max_stu_accept.append(max_stu_accept)

				# Registed
				stu_registed = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/div[2]/div[1]/h6/span").text
				# print(stu_registed)
				if stu_registed == "":
					data_stu_registed.append("NULL")
				else:
					data_stu_registed.append(stu_registed)


				# Accepted
				stu_accepted = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/div[2]/div[2]/h6/span").text
				# print(stu_accepted)
				if(stu_accepted == ""):
					data_stu_accepted.append("NULL")
				else:
					data_stu_accepted.append(stu_accepted)

				# LOcation
				location = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/p[3]").text
				# print(location)
				if(location == ""):
					data_location.append("NULL")
				else:
					data_location.append(location)

				# Email 
				# Crawl in many click <li> or just <ul>
				mail = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/ul").text
				# print(mail)
				if mail == "":
					data_email.append("NULL")
				else:
					data_email.append(mail)

				

				# Information (word, pdf) => download
				if dl_files == True:
					dl_file = browser.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div/h6[5]/a")
					dl_file.click()


				# data-id of figure .png
				#/img/company/5caef42002c2e72195b4cea8.png?t=74257893
				if dl_images == True:
					img_url_ = comp[i].find_element_by_xpath('..').get_attribute("style").split(";")[-3].split('"')[-2]
					img_url = "https://internship.cse.hcmut.edu.vn" + img_url_
					# print(img_url)
					dl_img(img_url, os.getcwd() + "\\Internship\\Logo", "{}.{}".format(name, img_url.split("?")[-2].split(".")[-1]))
					
				turn_off.click()


		#### DEAL WITH SPONSOR COMPANY
		#n = 76, m = 8 => 84 (2/6/2021)
		# m = len(data_sponsor_url)
		# print(m)

		# https://internship.cse.hcmut.edu.vn/sponsor/vng-corporation
		for sponsor_url in data_sponsor_url:
			idx += 1
			bar.update(idx)
			browser.get(sponsor_url)
			
			# Name
			sleep(3)
			name = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/h4").text
			# print(name)
			data_name.append(name)

			# Introduction
			intro = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/p[1]").text
			# print(intro)
			if intro == "":
				data_introduction.append("NULL")
			else:
				data_introduction.append(intro)

			# Field
			field = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/p[2]").text
			# print(field)
			if field == "":
				data_field.append("NULL")
			else:
				data_field.append(field)

			# Max register
			max_register = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/div[1]/div[1]/h6/span").text
			# print(max_register)
			if max_register == "":
				data_max_register.append("NULL")
			else:
				data_max_register.append(max_register)

			# Max accept
			max_stu_accept = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/div[1]/div[2]/h6/span").text
			# print(max_stu_accept)
			if max_stu_accept == "":
				data_max_stu_accept.append("NULL")
			else:
				data_max_stu_accept.append(max_stu_accept)

			# Registed
			stu_registed = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/div[2]/div[1]/h6/span").text
			# print(stu_registed)
			if stu_registed == "":
				data_stu_registed.append("NULL")
			else:
				data_stu_registed.append(stu_registed)


			# Accepted
			stu_accepted = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/div[2]/div[2]/h6/span").text
			# print(stu_accepted)
			if(stu_accepted == ""):
				data_stu_accepted.append("NULL")
			else:
				data_stu_accepted.append(stu_accepted)

			# Location
			location = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/p[3]").text
			# print(location)
			if(location == ""):
				data_location.append("NULL")
			else:
				data_location.append(location)

			# Email 
			# Crawl in many click <li> or just <ul>
			mail = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/ul").text
			# print(mail)
			if mail == "":
				data_email.append("NULL")
			else:
				data_email.append(mail)

			

			# Information (word, pdf) => download
			if dl_files == True:
				dl_file = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[2]/div/div/div/h6[5]/a")
				dl_file.click()


			# data-id of figure .png 
			# a little bit change: in <img>.src
			if dl_images == True:
				img_url = browser.find_element_by_xpath("/html/body/div/div[2]/div/div/div[1]/div/div/div/img").get_attribute("src")
				# print(img_url)
				dl_img(img_url, os.getcwd() + "\\Internship\\Logo", "{}.{}".format(name, img_url.split("?")[-2].split(".")[-1]))



		dic1 = {"Company ID": data_id, "Company name": data_name, "Introduction": data_introduction, "Field": data_field, "Max register": data_max_register, "Max student accept": data_max_stu_accept, "Number of students registed": data_stu_registed, "Number of students accepted": data_stu_accepted, "Location": data_location, "E-mail": data_email}
		df = pd.DataFrame(dic1)
		df.to_excel("Internship/Data_Intern.xlsx", encoding = 'utf-8-sig')

		dic2 = {"Company ID": data_id, "Company name": data_name, "Max register": data_max_register, "Max student accept": data_max_stu_accept, "Number of students registed": data_stu_registed, "Number of students accepted": data_stu_accepted, "Location": data_location}
		df = pd.DataFrame(dic2)
		df.to_csv("Internship/Dataset.csv", encoding = 'utf-8-sig')

		bar.finish()
			
		browser.close()
		stop = timeit.default_timer()
		print('Time: ', stop - start)

		# Change it to pdf
		if dl_files == True:
			print("################### CONVERT WORD TO PDF #######################")
			convert_word_to_pdf(dir_link)
		
		print("Everything finished")
	except TimeoutException:
		print("Loading took too much time!")

if __name__ == '__main__':
	main(True, True)


# 86 companies (6/5/2021): Run in 570s ~ 10 min







