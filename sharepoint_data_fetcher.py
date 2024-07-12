from flask import Flask, jsonify
from sqlalchemy import create_engine
import pandas as pd
import io
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

app = Flask(__name__)

# Database tilkobling
engine = create_engine('DATABASE_CONNECTION_STRING')  # Bytt til din skybaserte database

# Funksjon for å oppdatere data fra SharePoint
def update_data():
    url = 'https://company.sharepoint.com/Shared%20Documents/Folder/Target_Excel_File.xlsx'
    username = 'Dumby_account@company.com'
    password = 'Password!'

    ctx_auth = AuthenticationContext(url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print("Authentication successful")

    response = File.open_binary(ctx, url)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0) 

    df = pd.read_excel(bytes_file_obj, sheet_name=None)
    for sheet_name, sheet_data in df.items():
        sheet_data.to_sql(sheet_name, con=engine, if_exists='replace', index=False)
    print("Data updated successfully")

@app.route('/data/<sheet_name>', methods=['GET'])
def get_data(sheet_name):
    df = pd.read_sql(f'SELECT * FROM {sheet_name}', engine)
    return jsonify(df.to_dict(orient='records'))

@app.route('/update', methods=['POST'])
def manual_update():
    try:
        update_data()
        return jsonify({"status": "success", "message": "Data updated successfully"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run()
