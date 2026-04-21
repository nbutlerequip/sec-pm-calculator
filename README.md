# SEC PM Calculator

Preventive Maintenance quoting tool for Southeastern Equipment Company.

## Setup

1. Install dependencies: `pip install -r requirements.txt`
2. Configure Google Sheets credentials in `.streamlit/secrets.toml`
3. Run: `streamlit run app.py`

## Deployment

Connected to Streamlit Cloud via GitHub. Pushes to `main` auto-deploy.

## Google Sheets Setup

1. Create a Google Cloud service account
2. Share your Google Sheet with the service account email
3. Add credentials to Streamlit Cloud secrets

### Sheet Columns (auto-created on first save):
Date, Customer Name, Branch, Rep, Service Type, Make, Model, Category, Serial, Machine Age, Machine Hours, Travel Time, Parts Cost, Labor Cost, Travel Cost, Total Cost, Annual PM Price, Margin %, Notes
