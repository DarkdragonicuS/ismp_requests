import biip

def get_ki(raw_data:str):
    return parse_barcode_data(raw_data)['ki']

def get_gtin(raw_data:str):
    return parse_barcode_data(raw_data)['gtin']

def parse_barcode_data(raw_data:str):
    parsed_data = biip.parse(raw_data)
    if parsed_data.gtin_error is None and len(parsed_data.gtin.value) == 14:
        gtin = parsed_data.gtin.value
        barcode = parsed_data.gtin.payload
        serial = parsed_data.gs1_message.element_strings.get(ai="21").value
        ki = parsed_data.gs1_message.element_strings[0].ai.ai + gtin + "21" + serial
        data = {
            'gtin': gtin,
            'serial': serial,
            'barcode': barcode,
            'ki': ki
        }
    else:
        barcode = parsed_data.value
        gtin = barcode.zfill(14)
        data = {
            'gtin': gtin,
            'barcode': barcode
        }
    return data
    