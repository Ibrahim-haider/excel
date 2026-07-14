# Silver Line Streamlit Portal

A full Streamlit + SQLite portal for RD Electronics Silver Line partner referrals.

## Features

- Admin login
- Partner login with unique partner code and password
- Admin dashboard with KPIs, charts, leaderboard and CSV export
- Admin can add partners and generate partner login codes
- Admin can add referrals for any partner
- Partner can add customer referrals
- Partner can view own referrals and commission status
- SQLite database created automatically on first run

## Demo Login

Admin:

- Username: `admin`
- Password: `admin123`

Demo partner passwords are `partner001`, `partner002`, etc. Their usernames are their partner codes, visible inside the admin partner directory.

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, `database.sql`, and this README.
3. Go to Streamlit Cloud.
4. Select your repo and set main file as `app.py`.
5. Deploy.

## Important

Change the default admin password before using it with real data. This demo stores data in local SQLite. For serious production use, connect PostgreSQL/Supabase/MySQL instead.
