from docx import Document
import requests



doc = Document("docx_migration_test_file.docx")

html_content = "<html><body>\n"

for para in doc.paragraphs:
    text = para.text.strip()

    if text:
        html_content += f"<p>{text}</p>\n"

html_content += "</body></html>"




html_file = "naresh_migration_output.html"

with open(html_file, "w", encoding="utf-8") as f:
    f.write(html_content)

print("HTML file generated:", html_file)




with open(html_file, "r", encoding="utf-8") as f:
    html_data = f.read()



url = "https://apihub.document360.io/v2/articles"

headers = {
    "api_token": "TvrDSpEEHxJVgfim9Gbqpw9ZZ5Vz2XwCFnPFqQ2DMFBl/A+ZI4PAfRO9qOGyG14nflFJ5n/8HuMsM2wiNzlBmT0QTuT1ktlNh0ueEqCKaRaSwygmVAbl3u5oKyyt3IYcKTLXz8XbxfYcXyG36rCFvA==",
    "Content-Type": "application/json"
}

payload = {
    
    "Title": "Naresh Migration Article",
    "Content": html_data,
    "UserId": "9bc90a13-4c7a-45e6-b51d-839bc3a5a209",
    "status": 1
}





response = requests.post(url, headers=headers, json=payload)

print("Status Code:", response.status_code)
print("Response:", response.text)