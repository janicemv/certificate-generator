# Certificate Generator

This project generates personalized certificates from an Excel spreadsheet and a Word template. It replaces placeholders in the template with real data and saves the output in both DOCX and PDF formats.

I used this code in a more robust way to generate the certificates for the [European Ovary Workshop](https://europeanovaryworkshop.com/), with specific templates for posters, presentations, attendees and organisers.

## Features
- Reads participant data from an Excel file
- Uses a Word template with placeholders
- Generates certificates in DOCX format
- Converts DOCX certificates to PDF

## Prerequisites
Ensure you have the following installed:
- Python 3.x
- Microsoft Word
- PDF Reader (to see your generated files)
- Required Python libraries:
  ```sh
  pip install pandas python-docx comtypes
  ```

## Folder Structure
```
project-root/
│-- data/
│   ├── example.xlsx  # Excel file with participant data
│-- templates/
│   ├── certificate.docx  # Word template with placeholders
│-- certificates_docx/  # Generated DOCX certificates
│-- certificates_pdf/  # Generated PDF certificates
│-- script.py  # Main script
```

## Usage
1. Place the participant data in `data/example.xlsx` with the following structure:

   | first_name | last_name  |
   |------------|------------|
   | Lyra       | Nightbloom |
   | Elowen     | Stormrider |

2. Create a Word template (`certificate.docx`) with placeholders:
   ```
   {{FIRST_NAME}} {{LAST_NAME}}
   Date: {{DATE}}
   ```

3. Run the script:
   ```sh
   python script.py
   ```

4. The generated certificates will be stored in `certificates_docx/` and `certificates_pdf/`.

## Notes
- Uncomment line 44 and comment the following lines to check the output before generating the files. If the response is correct, uncomment the code to generate your certificates normally. Otherwise, fix the placeholders to match the script.
- A delay (`time.sleep(2)`) is added before converting to PDF to ensure the DOCX file is saved properly.
- To ensure images are correctly placed, I had to insert the signature as a watermark in the template. Future versions will address this issue by enabling direct image insertion into the DOCX template.

## License
This project is open-source and free to use.
