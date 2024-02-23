import os
import requests
import xml.etree.ElementTree as ET

class strSplit:
    def __init__(self):
        self.resString = ""

def create_soap_envelope(index, strAzione, strCodiceCliente, strCodiceLottoAffido):
    soap_envelope = f'''<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:eic="https://services.engie.it/ws/EICreditMgmtCM26.ws.provider:EAI_CM26">
        <soapenv:Header/>
        <soapenv:Body>
            <eic:retrieveCreditPosition>
                <Input>
                    <Codice_AdR>AXTR2505</Codice_AdR>
                    <Azione>{strAzione}</Azione>
                    <Codice_Cliente>{strCodiceCliente}</Codice_Cliente>
                    <Codice_LottoAffido>{strCodiceLottoAffido}</Codice_LottoAffido>
                </Input>
            </eic:retrieveCreditPosition>
        </soapenv:Body>
    </soapenv:Envelope>'''
    return soap_envelope

def get_data_dir():
    return os.getcwd()

def create_web_request(url, action):
    headers = {"Content-Type": "text/xml;charset=\"utf-8\"", "Accept": "text/xml"}
    web_request = requests.post(url, headers=headers)
    return web_request

def insert_soap_envelope_into_web_request(soap_envelope_xml, web_request):
    web_request.data = soap_envelope_xml
    return web_request

def selector(cell):
    if cell is None:
        return ""
    cell_type = type(cell)
    if cell_type == float:
        return str(cell)
    elif cell_type == str:
        return cell
    elif cell_type == bool:
        return str(cell)
    else:
        return "unknown"

def remove_duplicates(values):
    return list(set(values))

def set_string_array(int_count, value):
    return [value] * int_count

def main():
    # WSDL Endpoint
    url = "https://some.com/wsdl"
    # WSDL action
    action = "EICreditMgmtCM26_ws_EAI_CM26_Port"
    
    try:
        data_dir = get_data_dir()
        files = [f for f in os.listdir(data_dir) if f.endswith(".xlsx")]

        if len(files) > 1:
            log_file = os.path.join(data_dir, f"logError_{str.replace(str.replace(str(datetime.datetime.now()), ':', '-'), '/', '-')}.log")
            with open(log_file, "w") as f:
                f.write("There is more than 1 source file\n")
            return

        if not files:
            log_file = os.path.join(data_dir, f"logError_{str.replace(str.replace(str(datetime.datetime.now()), ':', '-'), '/', '-')}.log")
            with open(log_file, "w") as f:
                f.write("Absent source Excel file\n")
            return

        input_file = os.path.join(data_dir, files[0])
        workbook = openpyxl.load_workbook(input_file)
        sheet = workbook.active

        soap_envelopes = []
        soap_results = []

        for row in sheet.iter_rows(min_row=1, max_col=10, max_row=1000):
            strAzione, strCodiceCliente, strCodiceLottoAffido = (selector(cell.value) for cell in row)
            soap_envelope = create_soap_envelope(0, strAzione, strCodiceCliente, strCodiceLottoAffido)
            soap_envelopes.append(soap_envelope)

        for soap_envelope in soap_envelopes:
            web_request = create_web_request(url, action)
            web_request = insert_soap_envelope_into_web_request(soap_envelope, web_request)
            response = requests.post(url, data=soap_envelope)
            soap_results.append(response.text)

        united_output_file = os.path.join(data_dir, "united_output_file.xml")
        with open(united_output_file, "w") as f:
            f.write("<?xml version=\"1.0\"?><messaggi>\n")
            for result in soap_results:
                f.write(result)
            f.write("</messaggi>\n")

    except Exception as e:
        log_file = os.path.join(data_dir, f"logError_{str.replace(str.replace(str(datetime.datetime.now()), ':', '-'), '/', '-')}.log")
        with open(log_file, "w") as f:
            f.write("There was an error: {}\n".format(str(e)))

if __name__ == "__main__":
    main()
