import base64
import win32com.client
import requests
import json
from datetime import datetime,timezone
import pandas as pd
import sqlite3
from pathlib import Path
from ismp_lib.config import *

def db_init():
    dbcon = sqlite3.connect(DB_NAME)
    create_query = '''
        CREATE TABLE IF NOT EXISTS tokens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inn TEXT NOT NULL,
        token TEXT NOT NULL,
        expire_date TEXT NOT NULL
        ) 
    '''
    cursor = dbcon.cursor()
    cursor.execute(create_query)
    dbcon.commit()
    dbcon.close()
    
# Подписание данных, упакованных в base64
def signData(b64data,cert):
    CADESCOM_BASE64_TO_BINARY = 1 # Входные данные пришли в Base64
    CADESCOM_CADES_TYPE = 1 # Тип усовершенствованной подписи
    CAPICOM_AUTHENTICATED_ATTRIBUTE_SIGNING_TIME = 0 # Атрибут штампа времени подписи
    oSigner = win32com.client.Dispatch("CAdESCOM.CPSigner")

    oSigner.Certificate = cert
    oSigningTimeAttr = win32com.client.Dispatch("CAdESCOM.CPAttribute")
    oSigningTimeAttr.Name = CAPICOM_AUTHENTICATED_ATTRIBUTE_SIGNING_TIME
    oSigningTimeAttr.Value = datetime.now()
    oSigner.AuthenticatedAttributes2.Add(oSigningTimeAttr)
    oSignedData = win32com.client.Dispatch("CAdESCOM.CadesSignedData")
    oSignedData.ContentEncoding = CADESCOM_BASE64_TO_BINARY
    oSignedData.Content = b64data
    EncodingType = 0
    sSignedMessage = oSignedData.SignCades(oSigner,CADESCOM_CADES_TYPE,False,EncodingType);

    return sSignedMessage

# AI generated (sorry v_v)
# Разбор строки субъекта в сертификате
def parse_cert_parameter(parameter_str:str):
    subject_dict = {}
    current_pair = []
    in_quotes = False
    i = 0
    n = len(parameter_str)
    
    while i < n:
        char = parameter_str[i]
        
        # Если встретили кавычку
        if char == '"':
            # Проверяем на удвоенную кавычку "" внутри кавычек
            if in_quotes and i + 1 < n and parameter_str[i+1] == '"':
                current_pair.append('"') # Добавляем одну кавычку вместо двух
                i += 2
                continue
            else:
                in_quotes = not in_quotes
                i += 1
                continue
                
        # Если встретили запятую-разделитель (только вне кавычек)
        if char == ',' and not in_quotes:
            pair_str = "".join(current_pair).strip()
            if '=' in pair_str:
                k, v = pair_str.split('=', 1)
                subject_dict[k.strip()] = v
            current_pair = []
        else:
            current_pair.append(char)
        i += 1
        
    # Дописываем последнюю пару после цикла
    if current_pair:
        pair_str = "".join(current_pair).strip()
        if '=' in pair_str:
            k, v = pair_str.split('=', 1)
            subject_dict[k.strip()] = v
            
    return subject_dict

# Определяет ИНН организации в сертификате
def extract_cert_inn(subject_dict:dict):
    inn = ''
    if 'ОГРН' in subject_dict and 'ИНН ЮЛ' in subject_dict:
        inn = subject_dict['ИНН ЮЛ']
    elif 'ИНН' in subject_dict:
        inn = subject_dict['ИНН']
    return inn

# Определяет вид владельца сертификата
def kind_of_owner(subject_dict:dict):
    kind = ''
    if 'ОГРН' in subject_dict:
        kind = 'ЮЛ'
    elif 'ОГРНИП' in subject_dict:
        kind = 'ИП'
    else:
        kind = 'ФЛ'
    return kind

# Проверка наличия действующего токена
def check_available_auth_token(inn):
    dbcon = sqlite3.connect(DB_NAME)
    select_query = f"""
        SELECT token,expire_date FROM tokens
        WHERE inn == {inn}
    """
    cursor = dbcon.cursor()
    cursor.execute(select_query)
    tokens = cursor.fetchall()
    dbcon.close()
    if len(tokens) != 0:
        token = tokens[0][0]
        expire_date = tokens[0][1]
        if (datetime.fromisoformat(expire_date) - datetime.now(timezone.utc)).total_seconds() >= 300: # обновляем токен, если скоро истечёт
            return token
        else:
            return None
    else:
        return None

# Сохранение токена для дальнейшего использования
def save_token(token,inn,expire_date):
    dbcon = sqlite3.connect(DB_NAME)
    cursor = dbcon.cursor()
    query = f"DELETE FROM tokens WHERE inn={inn}"
    cursor.execute(query)
    query = f"INSERT INTO tokens (token,expire_date,inn) VALUES ('{token}','{expire_date}','{inn}')"
    cursor.execute(query)
    dbcon.commit()
    dbcon.close()

# Поиск доступных сертификатов
def availableCertList():
    CAPICOM_CURRENT_USER_STORE = 2
    CAPICOM_MY_STORE = "My"
    CAPICOM_STORE_OPEN_READ_ONLY = 0
    oStore = win32com.client.Dispatch("CAdESCOM.Store")

    oStore.Open(CAPICOM_CURRENT_USER_STORE, CAPICOM_MY_STORE,CAPICOM_STORE_OPEN_READ_ONLY)
    # поиск сертификата по отпечатку
    # CAPICOM_CERTIFICATE_FIND_SHA1_HASH = 0;
    # Certificates = oStore.Certificates.Find(CAPICOM_CERTIFICATE_FIND_SHA1_HASH, fingerprint);
    # cert = Certificates.Item(1);

    certList = []
    for cert in oStore.Certificates: 
        if cert.ValidToDate >= datetime.now(timezone.utc): # Игнорируем истёкшие
            subject_name_dict = parse_cert_parameter(cert.SubjectName)
            inn = extract_cert_inn(subject_name_dict)
            if inn == '' or inn.startswith('00'): # Игнорируем неподходящие подписи
                continue
            kind = kind_of_owner(subject_name_dict)
            cn = subject_name_dict['CN']

            certData = {}
            certData["CertObj"] = cert
            certData['Subject'] = cert.SubjectName
            certData['CN'] = cn
            certData['INN'] = inn
            certData['Kind'] = kind
            certData["ValidFromDate"] = cert.ValidFromDate.strftime('%d.%m.%Y')
            certData["ExpireDate"] = cert.ValidToDate.strftime('%d.%m.%Y')
            certData["SerialNumber"] = cert.SerialNumber
            certData["Fingerprint"] = cert.Thumbprint
            certList.append(certData)
    return certList

def strToB64(dataString: str):
    return base64.b64encode(dataString.encode('UTF-8')).decode('ascii')

def getAuthToken(cert):
    subject_dict = parse_cert_parameter(cert.SubjectName)
    inn = extract_cert_inn(subject_dict)
    token = check_available_auth_token(inn)

    if token is None:
        url = ENDPOINT_URL + '/api/v3/true-api/auth/key'
        headers = {"Accept": "application/json"}
        resolve = requests.get(url,headers=headers)
        resolveDeserialised = json.loads(resolve.content)
        uuid = resolveDeserialised['uuid']
        dataForSighing = resolveDeserialised['data']
        dataForSighing = strToB64(dataForSighing)
        data = signData(dataForSighing,cert)
        body = {"uuid":uuid,"data":data,"unitedToken":True}
        headers['Content-Type'] = 'application/json; charset=UTF-8'
        url = ENDPOINT_URL + '/api/v3/true-api/auth/simpleSignIn'

        resolve = requests.post(url=url,headers=headers,json=body)

        token = ''
        if resolve.status_code == 200:
            resolve_dict = json.loads(resolve.content)
            token = resolve_dict['uuidToken']
            expireDate = resolve_dict['expireDate']
            save_token(token,inn,expireDate)
        return token
    else:
        return token

def getCisInfo(cis,auth_token):
    url = ENDPOINT_URL + '/api/v3/true-api/cises/info'
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json; charset=UTF-8",
        "Authorization": "Bearer " + auth_token
        }
    body = [cis]
    resolve = requests.post(url=url,headers=headers,json=body)
    cisInfo = None
    if resolve.status_code == 200:
        cisInfo = json.loads(resolve.content.decode('utf-8'))
    return cisInfo


# MAIN

# Создание БД, если ещё нет
db_path = Path(DB_NAME)
if not db_path.is_file():
    db_init()

certList = availableCertList()
df = pd.DataFrame(certList,columns=['CN','INN','Kind','ValidFromDate','ExpireDate'])

print(df)
certIndex = int(input("Номер сертификата: "))
cert = certList[certIndex]['CertObj']
while True:
    cis = input("Код маркировки: ")
    auth_token = getAuthToken(cert)
    cisInfo = getCisInfo(cis,auth_token)
    if cisInfo is not None:
        cisInfoStr = json.dumps(cisInfo,indent=4,ensure_ascii=False)
        print(cisInfoStr)