import requests

api_key = "378beee4fcd36a0518efcb724fefe2aace424601"
corp_codes = "00626011,00137997"
url = "https://opendart.fss.or.kr/api/fnlttMultiAcnt.xml"
params = {
    "crtfc_key": api_key,
    "corp_code": corp_codes,
    "bsns_year": "2025",
    "reprt_code": "11011"
}

res = requests.get(url, params=params)
print("XML Status Code:", res.status_code)
# write first 1500 chars of XML response
print(res.text[:1500])
