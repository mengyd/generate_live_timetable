import os, json

def loadParams(params_path):
    f = open(params_path,'r', encoding='UTF-8')
    params_data = json.load(f)
    return params_data

def loadAllias(allias_path):
    f = open(allias_path,'r', encoding='UTF-8')
    dict_allias = json.load(f)
    return dict_allias

def loadConfig(choice = None):
    workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
    params = loadParams(workpath+'/config.json')
    allias = loadAllias(workpath+'/allias.json')
    if choice == 'params':
        return params
    if choice == 'allias':
        return allias
    return params, allias

