import json
import streamlit as st
import requests

def supabase_request(method, table, params=None, json_data=None):
    url = st.secrets["supabase"]["url"] + "/rest/v1/" + table
    headers = {
        "apikey": st.secrets["supabase"]["key"],
        "Authorization": f"Bearer {st.secrets['supabase']['key']}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    if method == "GET":
        return requests.get(url, headers=headers, params=params).json()
    elif method == "POST":
        return requests.post(url, headers=headers, json=json_data).json()
    elif method == "PATCH":
        return requests.patch(url, headers=headers, params=params, json=json_data).json()

def save_financial_data(email, data_dict):
    if not email:
        return False
        
    company_name = data_dict.get("company_name", "Unknown")
    year = data_dict.get("year", "Unknown")
    json_data_str = json.dumps(data_dict)
    
    # Check if exists to know whether to insert or update
    res = supabase_request("GET", "user_files", params={
        "select": "id", 
        "email": f"eq.{email}",
        "company_name": f"eq.{company_name}",
        "year": f"eq.{year}"
    })
    
    try:
        if res and isinstance(res, list) and len(res) > 0:
            supabase_request("PATCH", "user_files", params={"id": f"eq.{res[0]['id']}"}, json_data={
                "json_data": json_data_str
            })
        else:
            supabase_request("POST", "user_files", json_data={
                "email": email,
                "company_name": company_name,
                "year": year,
                "json_data": json_data_str
            })
        return True
    except Exception as e:
        print(f"Error saving to Supabase: {e}")
        return False

def load_user_financial_data(email):
    if not email:
        return {}
        
    res = supabase_request("GET", "user_files", params={
        "select": "company_name, year, json_data", 
        "email": f"eq.{email}"
    })
    
    companies = {}
    if res and isinstance(res, list):
        for row in res:
            key = f"{row['company_name']} - {row['year']}"
            try:
                companies[key] = json.loads(row["json_data"])
            except Exception:
                pass
                
    return companies
