import json
import requests
import codecs

def TryExpandAsJSON(d):
    for key, value in d.items():
        if type(value) is str:
            try:
                d[key] = json.loads(value)
            except:
                pass
            
        if type(d[key]) is dict:
            TryExpandAsJSON(d[key])
                

def FetchJSON(URL, asFile):
    print('Fetching and saving data to %s ... '%(asFile), end='', flush=True)
    r = requests.get(URL)
    if r.status_code == 200:
        JsonDict = r.json()

    if type(JsonDict) is dict:
        TryExpandAsJSON(JsonDict)

    f = codecs.open(asFile, "w", 'utf-8')
    json.dump(JsonDict, f, ensure_ascii=False)
    f.close()
    print('done. ')

FetchJSON(
    URL    = "https://is.snssdk.com/forum/ncov_data/?city_code=%5B%22420000%22%5D&data_type=%5B2%5D&src_type=province"
    # "https://i.snssdk.com/forum/ncov_data/?activeWidget=20&city_name=%E4%B8%8A%E6%B5%B7&data_type=%5B2%2C4%2C8%5D&src_type=map",
    # 'https://i.snssdk.com/forum/ncov_data/?country_id=["USA"]&country_name=美国&click_from=overseas_epidemic_tab_list&data_type=[5,4]&policy_scene=USA&src_type=country',
    ,
    asFile = 'Province.json' 
)