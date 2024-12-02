# Sharepoint-Scraper

this is a Python script designed to act as a scraper for a specific Microsoft Office 365 SharePoint site.

The project aims to synchronize with Google Chat and Google Meet integration for Graph8's virtual office application. It involves a series of tasks such as debugging Google OAuth 2.0, resolving SSL/TLS issues, and integrating seamless chat and meet functionality based on Google API.

It uses the requests library to interact with the Microsoft Graph API and retrieves a list of files, folders, and content pages from the SharePoint site. The data acquired (including name, URL, type, and last modification) are regularly updated and stored in two .csv files.

Additionally, the script handles downloading files from SharePoint to the local system, preserving the original SharePoint folders structure for every newly added or updated file.

The code heavily relies on Google OAuth 2.0 for user authorization, Netlify Functions for serverless operations, and Firebase for the database aspect of the project.