"""Generate a fake, messy HubSpot-style contacts export (sample data only)."""
import csv
import random
from datetime import datetime, timedelta

random.seed(42)

FIRST = ["Olivia", "Liam", "Emma", "Noah", "Ava", "Ethan", "Sophia", "Mason", "Isabella", "Lucas",
         "Mia", "Logan", "Amelia", "James", "Harper", "Benjamin", "Evelyn", "Elijah", "Abigail", "Daniel",
         "Priya", "Arjun", "Mei", "Wei", "Fatima", "Omar", "Chloe", "Jack", "Grace", "Ryan"]
LAST = ["Smith", "Johnson", "Brown", "Tremblay", "Martin", "Roy", "Wilson", "Gagnon", "Lee", "Taylor",
        "Campbell", "Anderson", "Patel", "Chen", "Nguyen", "Singh", "White", "Clark", "Lewis", "Walker"]
COMPANIES = ["Northwind Energy", "Maple Ridge Financial", "Bluewater Logistics", "Summit Health Group",
             "Granite Insurance", "Lakeshore Retail", "Polaris Tech", "Cedar Manufacturing",
             "Harbor Capital", "Aurora Telecom", "Pinecrest Foods", "Ironbridge Construction",
             "Silverline Pharma", "Redwood Credit Union", "Evergreen Utilities", "Atlas Mining"]
INDUSTRY = {"Northwind Energy": "Energy", "Maple Ridge Financial": "Financial Services",
            "Bluewater Logistics": "Transportation", "Summit Health Group": "Healthcare",
            "Granite Insurance": "Insurance", "Lakeshore Retail": "Retail", "Polaris Tech": "Technology",
            "Cedar Manufacturing": "Manufacturing", "Harbor Capital": "Financial Services",
            "Aurora Telecom": "Telecommunications", "Pinecrest Foods": "Food & Beverage",
            "Ironbridge Construction": "Construction", "Silverline Pharma": "Healthcare",
            "Redwood Credit Union": "Financial Services", "Evergreen Utilities": "Energy",
            "Atlas Mining": "Mining"}
TITLES = ["CFO", "Chief Financial Officer", "VP Finance", "Director of Finance", "Finance Manager",
          "Controller", "CEO", "COO", "Head of Operations", "HR Director", "Marketing Manager",
          "IT Manager", "Procurement Lead", "Treasurer", "Senior Accountant"]
STAGES = ["subscriber", "lead", "marketingqualifiedlead", "salesqualifiedlead", "opportunity",
          "customer", "evangelist", "other"]
STAGE_W = [10, 30, 20, 12, 8, 12, 3, 5]
COUNTRIES = ["Canada", "canada", "CA", "United States", "USA", "US", "United Kingdom", "UK"]
COUNTRY_W = [30, 8, 6, 18, 6, 4, 6, 3]
CITIES = {"Canada": ["Toronto", "Vancouver", "Calgary", "Montreal", "Ottawa"],
          "United States": ["New York", "Chicago", "Boston", "Dallas"],
          "United Kingdom": ["London", "Manchester"]}
OWNERS = ["Sarah K.", "Mike D.", "Jenna L.", ""]
SOURCES = ["Offline Sources", "Organic Search", "Email Marketing", "Paid Social", "Referrals", "Direct Traffic"]
NORM = {"canada": "Canada", "ca": "Canada", "united states": "United States", "usa": "United States",
        "us": "United States", "united kingdom": "United Kingdom", "uk": "United Kingdom"}


def messy(s):
    """Randomly mess up casing and whitespace, like real CRM data."""
    r = random.random()
    if r < 0.08:
        return s.upper()
    if r < 0.16:
        return s.lower()
    if r < 0.24:
        return f"  {s} "
    return s


rows = []
start = datetime(2023, 1, 1)
for i in range(1, 1501):
    fn, ln = random.choice(FIRST), random.choice(LAST)
    comp = random.choice(COMPANIES)
    country = random.choices(COUNTRIES, COUNTRY_W)[0]
    city = random.choice(CITIES[NORM[country.lower()]])
    domain = comp.lower().replace(" ", "").replace("&", "") + ".com"
    email = f"{fn.lower()}.{ln.lower()}{random.randint(1, 99)}@{domain}"
    if random.random() < 0.04:
        email = ""
    created = start + timedelta(days=random.randint(0, 1000))
    rows.append({
        "Record ID": 10000000 + i * 7,
        "First Name": messy(fn),
        "Last Name": messy(ln),
        "Email": messy(email).strip() if random.random() < 0.9 else email.upper(),
        "Phone Number": random.choice(["", f"+1 416 555 {random.randint(1000, 9999)}",
                                       f"(647) 555-{random.randint(1000, 9999)}"]),
        "Job Title": messy(random.choice(TITLES)) if random.random() > 0.05 else "",
        "Company Name": messy(comp),
        "Industry": INDUSTRY[comp] if random.random() > 0.1 else "",
        "City": city,
        "Country/Region": country,
        "Lifecycle Stage": random.choices(STAGES, STAGE_W)[0],
        "Lead Status": random.choice(["NEW", "OPEN", "IN_PROGRESS", "CONNECTED", "UNQUALIFIED", ""]),
        "Contact owner": random.choice(OWNERS),
        "Original Source": random.choice(SOURCES),
        "Marketing emails unsubscribed": random.choices(["false", "true", ""], [80, 12, 8])[0],
        "Create Date": created.strftime("%Y-%m-%d %H:%M"),
        "Last Activity Date": (created + timedelta(days=random.randint(0, 300))).strftime("%Y-%m-%d %H:%M")
        if random.random() > 0.15 else "",
    })

# Add duplicates (same person exported twice, slightly different)
for r in random.sample(rows, 120):
    d = dict(r)
    d["Record ID"] = r["Record ID"] + 3
    d["Email"] = r["Email"].upper() if r["Email"] else ""
    d["Last Activity Date"] = ""
    rows.append(d)
random.shuffle(rows)

with open("sample_hubspot_contacts_export.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader()
    w.writerows(rows)
print(f"Wrote {len(rows)} rows")
