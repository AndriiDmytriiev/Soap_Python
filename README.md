This script reads an Excel file containing client data, constructs SOAP requests for each row, sends them to a web service, and stores the responses in an XML file. Here's a breakdown of what it does:

1. Read an Excel file
It looks for .xlsx files in the current directory.
If there is not exactly one Excel file, it logs an error and exits.
2. Process the Excel file
It loads the first (and only) Excel file it finds.
It iterates through up to 1000 rows, extracting three values per row:
strAzione (Action)
strCodiceCliente (Client Code)
strCodiceLottoAffido (Lot Code)
3. Construct and send SOAP requests
For each row, it generates a SOAP XML envelope containing the extracted values.
It sends the SOAP request to a predefined web service (url = "https://some.com/wsdl").
It collects the responses.
4. Save responses to an XML file
It writes all SOAP responses into a single XML file named united_output_file.xml.
5. Error Handling
If multiple Excel files exist or none are found, it logs an error.
If an exception occurs during processing, it logs the error with a timestamp.
Potential Issues
The script assumes the Excel file has at least 3 columns in every row.
It does not check the validity of the SOAP responses.
The web service URL (https://some.com/wsdl) is likely a placeholder.
There is no retry mechanism for failed requests.
