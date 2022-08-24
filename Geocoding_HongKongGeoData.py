# Tang Justin Hayse Chi Wing G.
# Updating the Point of interest without coordinate

import urllib.parse
import pandas as pd
import time
import requests

print("Start updating the point of interest without coordinate")
start_time = time.time()

file_name = 'updated_updated2_No. of Failed Trees_Mangkhut.csv'
df_raw = pd.read_csv(file_name, encoding='utf-8', header=0)
df = df_raw.fillna('')
num_records = len(df)
error_index = 0
i = 0

while True:
    for i in range(num_records):
        if i < error_index:
            continue
        address_encode = urllib.parse.quote(df.xs(i)['Venue'])
        response = requests.get("https://geodata.gov.hk/gs/api/v1.0.0/locationSearch?q=" + address_encode)
        if response.status_code != 200:
            error_index = i
            print("Ops! Start again from Row ID:" + str(error_index))
            time.sleep(5)
            break
        json_response = response.json()
        if i % 200 == 0:
            print(i)
        try:
            df.xs(i)['x'] = json_response[0]['x']
            df.xs(i)['y'] = json_response[0]['y']
        except:
            print("Troublemaker row ID:" + str(i))
            pass
    if i == num_records - 1:
        break

df.to_csv("updated_" + file_name, encoding='utf-8', index=False)
print(time.time() - start_time)
