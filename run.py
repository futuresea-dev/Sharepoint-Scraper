import requests
import csv
from urllib.parse import urlparse, parse_qs, unquote
import os
import time


base_path = os.getcwd()
file_path = "output/database.csv"
file_content_path = "output/database_content.csv"

def acquire_token_func():
    class Token:
        def __init__(self, access_token, token_type):
            self.accessToken = access_token
            self.tokenType = token_type

    tenant_id = os.environ.get('tenant_id', '')
    client_id = os.environ.get('client_id', '')
    client_secret = os.environ.get('client_secret', '')

    # Get the access token using client credentials
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    body = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(url, headers=headers, data=body)
    token_data = response.json()

    if "access_token" not in token_data:
        raise ValueError(f"Error obtaining access token: {token_data.get('error_description')}")
    token = Token(token_data["access_token"], token_data["token_type"])
    return token


def extract_last_element(url):
    # Parse the URL
    parsed_url = urlparse(url)

    # Get the path and split it by '/'
    path_segments = parsed_url.path.split('/')

    # Return the last element
    return path_segments[-1] if path_segments else None


def make_folder_path(url):
    # Parse the URL
    parsed_url = urlparse(url)

    path = parsed_url.path.replace("/Shared%20Documents/", "Document/")

    path = unquote(path)  # Decoding URL-encoded characters
    return path.lstrip('/')


def scrape_files(drive_id, headers, children):
    url = children['webUrl']
    name = children['name']
    modified = children['lastModifiedDateTime']
    item_type = 'file'
    status = check_record_in_csv(url, modified)
    if status == False:
        update_or_add_record(name, url, item_type, modified)
        child_id = children['id']
        item_url = f'https://graph.microsoft.com/v1.0/sites/root/drives/{drive_id}/items/{child_id}/content'
        response = requests.get(item_url, headers=headers)

        if response.status_code == 200:
            url = children['webUrl']
            base_path = os.getcwd()
            folder_path = make_folder_path(url)
            local_file_path = os.path.join(base_path, folder_path)
            # Create directories if they do not exist
            if not os.path.exists(os.path.dirname(local_file_path)):
                try:
                    os.makedirs(os.path.dirname(local_file_path))
                except OSError as e:
                    print(f"An error has occurred: {e}")
                    raise
            # Open the local file in binary write mode
            with open(local_file_path, 'wb') as f:
                # Write the file in chunks
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"File downloaded successfully: {local_file_path}")
        else:
            print(f"Failed to download file. Status code: {response.status_code}")


def update_or_add_record(name, url, record_type, modified_datetime):
    records = []
    record_found = False

    # Read the existing records from the CSV file
    with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            if row['URL'] == url:
                # Update the existing record
                row['modified'] = modified_datetime
                record_found = True
            # Append the row (updated or not) to records
            records.append(row)

    # If the record was not found, add the new record
    if not record_found:
        records.append({
            'Name': name,
            'URL': url,
            'Type': record_type,
            'modified': modified_datetime
        })

    # Write the updated records back to the CSV file
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        fieldnames = ['Name', 'URL', 'Type', 'modified']
        csv_writer = csv.DictWriter(file, fieldnames=fieldnames)
        csv_writer.writeheader()
        csv_writer.writerows(records)


def update_or_add_content_record(name, title, page_layout, promotion_kind, description, url, thumb_url, modified):
    records = []
    record_found = False

    # Read the existing records from the CSV file
    with open(file_content_path, mode='r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            if row['URL'] == url:
                # Update the existing record
                row['Modified'] = modified
                record_found = True
            # Append the row (updated or not) to records
            records.append(row)

    # If the record was not found, add the new record
    if not record_found:
        records.append({
            'Name': name,
            'Title': title,
            'PageLayout': page_layout,
            'PromotionKind': promotion_kind,
            "Description": description,
            "URL": url,
            "ThumbnailWebUrl" : thumb_url,
            "Modified": modified

        })

    # Write the updated records back to the CSV file
    with open(file_content_path, mode='w', newline='', encoding='utf-8') as file:
        fieldnames = ["Name", "Title", "PageLayout", "PromotionKind", "Description", "URL", "ThumbnailWebUrl", "Modified"]
        csv_writer = csv.DictWriter(file, fieldnames=fieldnames)
        csv_writer.writeheader()
        csv_writer.writerows(records)


def check_record_in_csv(url_to_check, new_modified_value):
    # Open the CSV file for reading
    with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)

        # Check each row in the CSV
        for row in csv_reader:
            # Compare the URL and modified values
            if row['URL'] == url_to_check:
                if row['modified'] != new_modified_value:
                    return False
                else:
                    return True
        return False

def check_record_in_content_csv(url_to_check, new_modified_value):
    # Open the CSV file for reading
    with open(file_content_path, mode='r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)

        # Check each row in the CSV
        for row in csv_reader:
            # Compare the URL and modified values
            if row['URL'] == url_to_check:
                if row['Modified'] != new_modified_value:
                    return False
                else:
                    return True
        return False


def search_folder(drive_id, headers, children):
    child_id = children['id']
    item_url = f'https://graph.microsoft.com/v1.0/sites/root/drives/{drive_id}/items/{child_id}/children'
    response = requests.get(item_url, headers=headers)
    if response.status_code == 200:
        lists = response.json()['value']
        for item in lists:
            url = item['webUrl']
            name = item['name']
            modified = item['lastModifiedDateTime']
            status = check_record_in_csv(url, modified)
            if not status:
                if 'file' in item:
                    scrape_files(drive_id, headers, item)
                elif 'folder' in item:
                    item_type = 'folder'
                    update_or_add_record(name, url, item_type, modified)
                    search_folder(drive_id, headers, item)


def scrape_content(headers):
    drive_url = f'https://graph.microsoft.com/v1.0/sites/root'
    response = requests.get(drive_url, headers=headers)

    # Check the response status code and content
    if response.status_code == 200:
        res = response.json()
        site_id = res['id']
        page_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/pages'
        response = requests.get(page_url, headers=headers)
        if response.status_code == 200:
            res = response.json()
            page_lists = res['value']
            for page_list in page_lists:
                name = page_list['name']
                title = page_list['title']
                page_layout = page_list['pageLayout']
                promotion_kind = page_list['promotionKind']
                if 'description' in page_list:
                    description = page_list['description']
                else:
                    description = ""
                modified = page_list['lastModifiedDateTime']
                url = page_list['webUrl']
                thumb_url = page_list['thumbnailWebUrl']
                status = check_record_in_content_csv(url, modified)
                if not status:
                    update_or_add_content_record(name, title, page_layout, promotion_kind, description, url, thumb_url, modified)

    else:
        print("Error:", response.status_code, response.text)


def main():
    token = acquire_token_func()

    if not os.path.isfile(file_path):
        header = ["Name", "URL", "Type", "modified"]
        with open(file_path, 'w', newline='') as file:
            csv.writer(file).writerow(header)
    if os.path.isfile(file_content_path) == False:
        header = ["Name", "Title", "PageLayout", "PromotionKind", "Description", "URL", "ThumbnailWebUrl", "Modified"]
        with open(file_content_path, 'w', newline='') as file:
            csv.writer(file).writerow(header)
    # Bearer token
    # Set the headers with the Bearer token
    headers = {
        "Authorization": f'{token.tokenType} {token.accessToken}'
    }
    # scrape content
    scrape_content(headers)
    drive_url = f'https://graph.microsoft.com/v1.0/sites/root/drives'
    response = requests.get(drive_url, headers=headers)

    # Check the response status code and content
    if response.status_code == 200:
        res = response.json()
        drive_id = res['value'][0]['id']
        # Check the response status code and content
        children_url = f'https://graph.microsoft.com/v1.0/sites/root/drives/{drive_id}/root/children'
        response = requests.get(children_url, headers=headers)
        if response.status_code == 200:
            res_children = response.json()['value']
            for children in res_children:
                if 'file' in children:
                    scrape_files(drive_id, headers, children)
                elif 'folder' in children:
                    search_folder(drive_id, headers, children)

    else:
        print("Error:", response.status_code, response.text)

    print("Task executed at:", time.strftime("%Y-%m-%d %H:%M:%S"))


if __name__ == "__main__":
    main()