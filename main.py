import os
from dotenv import load_dotenv
from utils.apify_client import scrape_linkedin_jobs
from utils.excel import save_to_excel

# Load environment variables from .env file
load_dotenv()

APIFY_TOKEN = os.getenv("APIFY_API_TOKEN")
EXCEL_FILE_NAME = os.getenv("EXCEL_FILE_NAME")

# Define the job URLs to scrape
analyst_1hr = "https://www.linkedin.com/jobs/search-results/?f_TPR=r3600&geoId=102713980&keywords=%22Analyst%22&origin=JOBS_HOME_SEARCH_BUTTON"
analyst_3hr = "https://www.linkedin.com/jobs/search-results/?f_TPR=r10800&geoId=102713980&keywords=%22Analyst%22&origin=JOBS_HOME_SEARCH_BUTTON"
analyst_24hrs = "https://www.linkedin.com/jobs/search-results/?f_TPR=r86400&geoId=102713980&keywords=%22Analyst%22&origin=JOBS_HOME_SEARCH_BUTTON"
analyst_1week = "https://www.linkedin.com/jobs/search-results/?f_TPR=r604800&geoId=102713980&keywords=%22Analyst%22&origin=JOBS_HOME_SEARCH_BUTTON"

# Prepare the Actor Input
run_input = {
    "urls" : [
        analyst_1hr,
        analyst_3hr,
        analyst_24hrs,
        analyst_1week
    ],
    "count": 100
}

if __name__ == "__main__":
    df = scrape_linkedin_jobs(APIFY_TOKEN, run_input)
    save_to_excel(df, EXCEL_FILE_NAME)
    print(f"✅ Jobs saved to '{EXCEL_FILE_NAME}' successfully.")