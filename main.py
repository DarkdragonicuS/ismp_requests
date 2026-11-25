import base64
import win32com.client
import requests
import json
from datetime import datetime,timezone
import sqlite3
from pathlib import Path
from ismp_lib.config import *
import flet as ft
import asyncio
from ismp_lib.devices import DeviceBarcodeScanner
import ismp_lib.devices as devices
from ismp_lib.barcodes import parse_barcode_data

# AI generated (sorry v_v)
# Для биндинга данных со значением поля flet
class DataModel:
    def __init__(self, initial_value=""):
        self._value = initial_value
        self._listeners = []

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, new_val):
        self._value = new_val
        # Уведомляем все привязанные элементы управления
        for listener in self._listeners:
            listener(new_val)

    def bind_to(self, callback):
        self._listeners.append(callback)

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
            return ''
    else:
        return ''

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

    if token == '':
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

# Получение полной информации о КИ
def getCisInfo(cis,auth_token,inn=''):
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

# Получение информации о GTIN
def getProductInfo(gtin,auth_token,inn=''):
    url = ENDPOINT_URL + '/api/v4/true-api/product/info'
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json; charset=UTF-8",
        "Authorization": "Bearer " + auth_token
        }
    body = {'gtins':[gtin]}
    resolve = requests.post(url=url,headers=headers,json=body)
    productInfo = None
    if resolve.status_code == 200:
        productInfo = json.loads(resolve.content.decode('utf-8'))
    return productInfo

# Получение списка места осуществления деятельности организации
def getModsInfo(inn,auth_token):
    url = ENDPOINT_URL + '/api/v3/true-api/mods/list'
    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer " + auth_token
        }
    body = {'inns':[inn]}
    resolve = requests.get(url=url,headers=headers,params=body)
    modsInfo = None
    if resolve.status_code == 200:
        modsInfo = json.loads(resolve.content.decode('utf-8'))
    return modsInfo

# Выводит информацию о марке в читаемом виде
def requestedCisNormalizedInfo(cisItem):
    result = {}
    result['requestedCis'] = cisItem['requestedCis']                                                    # Запрошенный КИ
    result['applicationDate'] = cisItem['applicationDate'] if 'applicationDate' in cisItem else None    # Дата нанесения
    result['introducedDate'] = cisItem['introducedDate'] if 'introducedDate' in cisItem else None       # Дата ввода в оборот
    result['gtin'] = cisItem['gtin'] if 'gtin' in cisItem else None                                     # GTIN

# MAIN

# Создание БД, если ещё нет
db_path = Path(DB_NAME)
if not db_path.is_file():
    db_init()

certList = availableCertList()

def main(page: ft.Page):
    
    selected_cert = None
    selected_inn = DataModel('')

    # Сброс флажков на других сертификатах
    def cb_certSelector_on_change(e):
        table_certs_rows = e.control.parent.parent.parent.rows
        for row in table_certs_rows:
            if e.control != row.cells[7].content:
                row.cells[7].content.value = False
        nonlocal selected_cert
        selected_cert = e.control.data
        subject_name_dict = parse_cert_parameter(selected_cert.SubjectName)
        nonlocal selected_inn
        selected_inn.value = extract_cert_inn(subject_name_dict)

    
    #lv_certs_controls = []
    rowNum = 0

    columns = [
                ft.DataColumn(label=ft.Text(value='N')),
                ft.DataColumn(label=ft.Text(value='CN')),
                ft.DataColumn(label=ft.Text(value='ИНН')),
                ft.DataColumn(label=ft.Text(value='Тип')),
                ft.DataColumn(label=ft.Text(value='Начало действия')),
                ft.DataColumn(label=ft.Text(value='Окончание действия')),
                ft.DataColumn(label=ft.Text(value='Токен')),
                ft.DataColumn(label=ft.Text(value=''))
    ]

    rows = []

    for cert in certList:

        rowNum += 1
        cert_obj = cert['CertObj']
        
        text_rowNum = ft.Text(value=str(rowNum),data=cert_obj)
        text_cn = ft.Text(value=cert['CN'],data=cert_obj)
        text_INN = ft.Text(value=cert['INN'],data=cert_obj)
        text_kind = ft.Text(value=cert['Kind'],data=cert_obj)
        text_ValidFromDate = ft.Text(value=cert['ValidFromDate'],data=cert_obj)
        text_ExpireDate = ft.Text(value=cert['ExpireDate'],data=cert_obj)
        text_auth_token = ft.TextField(check_available_auth_token(cert['INN']),read_only=True,expand=True,data=cert_obj)
        cb_certSelector = ft.Checkbox(value=False,data=cert_obj,on_change=cb_certSelector_on_change)

        row = ft.DataRow(
            cells = [
                ft.DataCell(text_rowNum),
                ft.DataCell(text_cn,expand=True),
                ft.DataCell(text_INN),
                ft.DataCell(text_kind),
                ft.DataCell(text_ValidFromDate),
                ft.DataCell(text_ExpireDate),
                ft.DataCell(text_auth_token),
                ft.DataCell(cb_certSelector)                                
            ]
        )

        rows.append(row)

    table_certs = ft.DataTable(columns=columns,rows=rows)
    table_container = ft.Column(height=300,controls=ft.Row(controls=table_certs,expand=True),scroll=ft.ScrollMode.ALWAYS)

    field_result = ft.TextField(multiline=True,read_only=True,expand=True,max_lines=30, min_lines=10)

    # Заполнение вкладки "Инфо о КИ"
    def tab_ki_info_content():
       
        # Получение инфо о КИ
        def btn_km_info_on_click(e: ft.ControlEvent):
            cert = selected_cert
            cis = field_km.value
            auth_token = getAuthToken(cert)
            update_auth_token_fields()
            cisInfo = getCisInfo(cis,auth_token)
            if cisInfo is not None:
                cisInfoStr = json.dumps(cisInfo,indent=4,ensure_ascii=False)
            else:
                cisInfoStr = 'Something went wrong'
            field_result.value = cisInfoStr

        # Получение инфо о КИ сканером
        async def btn_scanner_read_on_click(e: ft.ControlEvent):
            e.control.disabled = True
            yield
            field_km.value = devices.get_ki(device=scanner)
            e.control.disabled = False
            yield

        field_km = ft.TextField(keyboard_type=ft.KeyboardType.TEXT,label='КИ',expand=True)
        btn_scanner_read = ft.Button(content='Отсканировать',on_click=btn_scanner_read_on_click)
        btn_km_info = ft.Button(content='Получить',on_click=btn_km_info_on_click)

        return ft.Column(
                    controls=[
                        ft.Row(
                            controls=[
                                field_km,
                                btn_scanner_read,
                                btn_km_info
                            ]
                        )
                    ]
                )

    # Заполнение вкладки "Инфо о GTIN"
    def tab_product_info_content():
        
        # Получение инфо о КИ
        def btn_product_info_on_click(e: ft.ControlEvent):
            cert = selected_cert
            gtin = field_gtin.value
            auth_token = getAuthToken(cert)
            update_auth_token_fields()
            productInfo = getProductInfo(gtin,auth_token)
            if productInfo is not None:
                productInfoStr = json.dumps(productInfo,indent=4,ensure_ascii=False)
            else:
                productInfoStr = 'Something went wrong'
            field_result.value = productInfoStr

        # Получение инфо о КИ сканером
        async def btn_gtin_scanner_read_on_click(e: ft.ControlEvent):
            e.control.disabled = True
            yield
            field_gtin.value = devices.get_gtin(device=scanner)
            e.control.disabled = False
            yield

        field_gtin = ft.TextField(keyboard_type=ft.KeyboardType.TEXT,label='GTIN',expand=True)
        btn_gtin_scanner_read = ft.Button(content='Отсканировать',on_click=btn_gtin_scanner_read_on_click)
        btn_product_info = ft.Button(content='Получить',on_click=btn_product_info_on_click)

        return ft.Column(
                    controls=[
                        ft.Row(
                            controls=[
                                field_gtin,
                                btn_gtin_scanner_read,
                                btn_product_info
                            ]
                        )
                    ]
                )

    # Заполнение вкладки "Инфо о МОД"
    def tab_mod_info_content():
        # Получение инфо о МОДах
        def btn_mods_info_on_click(e):
            cert = selected_cert
            inn = selected_inn.value
            auth_token = getAuthToken(cert)
            update_auth_token_fields()
            modsInfo = getModsInfo(inn,auth_token)
            if modsInfo is not None:
                modsInfoStr = json.dumps(modsInfo,indent=4,ensure_ascii=False)
            else:
                modsInfoStr = 'Something went wrong'
            field_result.value = modsInfoStr

        field_inn = ft.TextField(keyboard_type=ft.KeyboardType.NUMBER,label='ИНН',expand=True,value=selected_inn.value)

        # AI generated (sorry v_v)
        # При изменении поля ввода обновляем прибинденное значение
        def field_inn_on_change(e):
            selected_inn.value = e.control.value
        field_inn.on_change = field_inn_on_change
        # Привязываем автоматическое обновление UI к изменению
        def update_ui(new_val):
            if field_inn.value != new_val:
                field_inn.value = new_val
            page.update()
        selected_inn.bind_to(update_ui)

        btn_mods_info = ft.Button(content='Получить',on_click=btn_mods_info_on_click)

        return ft.Column(
            controls=[
                ft.Row(
                    controls=[
                        field_inn,
                        btn_mods_info
                    ]
                )
            ]
        )

    page.add(
        ft.SafeArea(
            content=ft.Column(
                expand=True,
                controls=[
                    table_container,
                    field_result,
                    ft.Container(content=
                        ft.Tabs(
                            selected_index=0,
                            length=2,
                            expand=True,
                            content=ft.Column(
                                expand=True,
                                controls=[
                                    ft.TabBar(
                                        tabs=[
                                            ft.Tab(label="Инфо о КИ"),
                                            ft.Tab(label="Инфо о GTIN"),
                                            ft.Tab(label="Инфо о МОД")
                                        ]
                                    ),
                                    ft.TabBarView(
                                        expand=True,
                                        controls=[
                                            # Инфо о КИ
                                            tab_ki_info_content(),
                                            # Инфо о GTIN
                                            tab_product_info_content(),
                                            # Инфо о МОДах
                                            tab_mod_info_content()
                                        ]
                                    )
                                ]    
                            )
                        )
                    )
                ]
            )
                        
        )
    )
    page.scroll = ft.ScrollMode.AUTO

    def update_auth_token_fields():
        for row in table_certs.rows:
            row.cells[6].content.value = check_available_auth_token(row.cells[2].content.value)

if __name__ == "__main__":    
    scanner = DeviceBarcodeScanner()
    scanner.mode = BARCODE_SCANNER_MODE
    scanner.port = BARCODE_SCANNER_PORT
    scanner.baudrate = BARCODE_SCANNER_RATE
    scanner.suffix = BARCODE_SCANNER_SUFFIX

    ft.run(main)