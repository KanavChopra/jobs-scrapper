from apify_client import ApifyClient
from datetime import datetime
import pytz
import pandas as pd

def scrape_linkedin_jobs(apify_token: str, actor_input: dict) -> pd.DataFrame:
    """Scrapes job listings from LinkedIn using Apify's LinkedIn Jobs Scraper actor.

    Args:
        apify_token (str): Apify API token.
        actor_input (dict): Input configuration for the Apify actor.

    Returns:
        pd.DataFrame: DataFrame containing the scraped job listings.
    """
    client = ApifyClient(apify_token)

    print("🚀 Starting LinkedIn Job Scraper...")
    # Start the Actor
    run = client.actor("curious_coder/linkedin-jobs-scraper").call(run_input=actor_input)

    # Wait for the Actor to finish and get the results
    dataset_id = run["defaultDatasetId"]
    print(f"💾 Data stored in dataset: {dataset_id}")
    print(f"🔗 View data at: https://console.apify.com/storage/datasets/{dataset_id}")

    dataset = client.dataset(dataset_id)
    items = list(dataset.list_items().items)

    # Convert the results to a DataFrame
    df = pd.DataFrame(items)
    print(df.head())

    # Add a timestamp column
    tz = pytz.timezone('Asia/Kolkata')
    current_time = datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')
    df['current_timestamp'] = current_time

    return df