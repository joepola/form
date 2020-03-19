# import xlwings as xw

# xw.Book('Book2')

# for item in range(1, 10):
#     cellspot = "A{}".format(item)
#     if xw.Range(cellspot).value == 9:
#         xw.Range(cellspot).value = "here"
import json
import time
from requests import get, post

def fucntiontest():
# Endpoint URL
    endpoint = "https://formrecognizer123.cognitiveservices.azure.com/"
    apim_key = "7b0d2566b8ad483f8ca74324289aa5b1"
  
    post_url = endpoint + "/formrecognizer/v2.0-preview/prebuilt/receipt/analyze"
    #source = r"file:///C:/Users/rc08462/Pictures/Saved%20Pictures/Sample-Invoice-printable.webp"
    source = "invoice.jpg"

    headers = {
        # Request headers
        'Content-Type': 'image/jpeg',
        'Ocp-Apim-Subscription-Key': apim_key,
    }

    params = {
        "includeTextDetails": True
    }

    with open(source, "rb") as f:
        data_bytes = f.read()

    try:
        resp = post(url = post_url, data = data_bytes, headers = headers, params = params)
        if resp.status_code != 202:
            #print("POST analyze failed:\n%s" % resp.text)
            quit()
        #print("POST analyze succeeded:\n%s" % resp.headers)
        get_url = resp.headers["operation-location"]
    except Exception as e:
        #print("POST analyze failed:\n%s" % str(e))
        quit()

    n_tries = 10
    n_try = 0
    wait_sec = 6
    while n_try < n_tries:
        try:
            resp = get(url = get_url, headers = {"Ocp-Apim-Subscription-Key": apim_key})
            resp_json = json.loads(resp.text)
            if resp.status_code != 200:
                #print("GET Layout results failed:\n%s" % resp_json)
                quit()
            status = resp_json["status"]
            if status == "succeeded":
                #print("Layout Analysis succeeded:\n%s" % resp_json)
                return resp_json
            if status == "failed":
                #print("Analysis failed:\n%s" % resp_json)
                quit()
            # Analysis still running. Wait and retry.
            time.sleep(wait_sec)
            n_try += 1     
        except Exception as e:
            msg = "GET analyze results failed:\n%s" % str(e)
            #print(msg)
            quit()
    return False

pleasework = fucntiontest()

# print(pleasework)
# print(pleasework["analyzeResult"]["documentResults"][0]["fields"]["Tax"]["valueNumber"])
print(pleasework["analyzeResult"]["documentResults"])
