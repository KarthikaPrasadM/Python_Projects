import requests
import pandas as pd
import os

path = input("Enter excel path: ")
# path = (r"C:\Users\nandu.prasad\Desktop\url Validation Check\sample_input")
excel_path_files = []
for file in os.listdir(path):
    if file.endswith(".xlsx"):
        pass

os.chdir(path)
df = pd.read_excel(file, skiprows = 1)

url_list = df["Website"].tolist()


Corrected_Final_urls = []
url_status = []


corrected_urls_list = []

for url in url_list:
    if url.startswith("Https://www.") or url.startswith("https://www.") or url.startswith("http://www.") or url.startswith("https://") or url.startswith("http://"):
        if url.endswith("/"):
            corrected_urls_list.append(url)
        else:
            url = url + "/"
            corrected_urls_list.append(url)
    elif url.startswith("www.") or url.startswith("WWW.") or url.startswith("Www."):
        url = "https://" + url
        if url.endswith("/"):
            corrected_urls_list.append(url)
        else:
            url = url + "/"
            corrected_urls_list.append(url)
    elif not (url.startswith("Https://www.") or url.startswith("https://www.") or url.startswith("http://www.") or url.startswith("https://") or url.startswith("http://") or url.startswith("www.") or url.startswith("WWW.") or url.startswith("Www.")):
        url = "https://www." + url
        if url.endswith("/"):
            corrected_urls_list.append(url)
        else:
            url = url + "/"
            corrected_urls_list.append(url)

final_url_list = []
exception_from_corrected_urls_list = []

for item in corrected_urls_list:
    try:
        response = requests.get(item)
        # print("URL is valid and exists on the internet")
        Corrected_Final_urls.append(item)
        url_status.append("Valid")
        final_url_list.append(item)
    except requests.ConnectionError as exception:
        # print("URL is not valid in the internet")
        Corrected_Final_urls.append(item)
        url_status.append("Not Valid")
        exception_from_corrected_urls_list.append(item)

print("Valid URLs: ", final_url_list)
# print("Not valid URLs: ", exception_from_corrected_urls_list)
print("Not valid URLs:")
for item in exception_from_corrected_urls_list:
    print(item)

df.insert(3, 'Corrected URLs', Corrected_Final_urls)
df.insert(4, 'URL Status', url_status)

df.to_excel("output.xlsx")
print("Please check Output file in location: ", path)