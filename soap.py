import os
import requests
import datetime
import openpyxl  # Fixed missing import


def create_soap_envelope(strAzione, strCodiceCliente, strCodiceLottoAffido):
    return f'''<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" 
        xmlns:eic="https://services.engie.it/ws/EICreditMgmtCM26.ws.provider:EAI_CM26">
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


def get_data_dir():
    return os.getcwd()


def create_web_request(url, soap_envelope):
    headers = {"Content-Type": "text/xml;charset=utf-8", "Accept": "text/xml"}
    return requests.post(url, headers=headers, data=soap_envelope)


def selector(cell):
    return str(cell) if cell is not None else ""


def main():
    url = "https://some.com/wsdl"
    action = "EICreditMgmtCM26_ws_EAI_CM26_Port"

    try:
        data_dir = get_data_dir()
        files = [f for f in os.listdir(data_dir) if f.endswith(".xlsx")]

        if len(files) != 1:
            error_message = "There is more than 1 source file" if len(files) > 1 else "Absent source Excel file"
            log_file = os.path.join(data_dir, f"logError_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log")
            with open(log_file, "w") as f:
                f.write(error_message + "\n")
            return

        input_file = os.path.join(data_dir, files[0])
        workbook = openpyxl.load_workbook(input_file)
        sheet = workbook.active

        soap_results = []

        for row in sheet.iter_rows(min_row=1, max_col=3, max_row=1000):  # Adjust max_col based on actual data structure
            strAzione, strCodiceCliente, strCodiceLottoAffido = (selector(cell.value) for cell in row[:3])
            soap_envelope = create_soap_envelope(strAzione, strCodiceCliente, strCodiceLottoAffido)
            response = create_web_request(url, soap_envelope)
            soap_results.append(response.text)

        output_file = os.path.join(data_dir, "united_output_file.xml")
        with open(output_file, "w") as f:
            f.write("<?xml version=\"1.0\"?><messaggi>\n")
            f.writelines(soap_results)
            f.write("</messaggi>\n")

    except Exception as e:
        log_file = os.path.join(data_dir, f"logError_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log")
        with open(log_file, "w") as f:
            f.write(f"There was an error: {e}\n")


if __name__ == "__main__":
    main()
