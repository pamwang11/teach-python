from bs4 import BeautifulSoup
import urllib3
import pprint
URL = 'https://nvd.nist.gov/vuln/search/results?adv_search=true&form_type=advanced&results_type=overview&cvss_version=3&cvss_v3_severity=HIGH&startIndex=0'
http = urllib3.PoolManager()
ret = http.request('GET', URL )



soup = BeautifulSoup(ret.data, 'html.parser')

na = soup.find_all('a')

for na_get in na:
    if 'CVE' in str(na_get.string):
        print(na_get.string)
        print('https://nvd.nist.gov/vuln/detail/'+ na_get.string)

