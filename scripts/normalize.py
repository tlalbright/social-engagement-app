#!/usr/bin/env python3
"""
Limelight Spreadsheet Normalizer
Processes Limelight LinkedIn engagement exports for import into Ringlead/HubSpot/Salesforce.

Usage:
    python normalize.py <input_file.xlsx> [output_file.xlsx]

If output_file is omitted, saves as <input_name>_normalized.xlsx in the same folder.
"""

import sys
import json
import os
import re
import csv
import uuid
from pathlib import Path
from datetime import datetime

# Fixed namespace for deterministic engagement IDs (UUID5)
# Same LinkedIn URL + Post URL + Engagement Type always produces the same ID.
_ENGAGEMENT_NS = uuid.UUID("6ba7b810-9dad-11d1-80b4-00c04fd430c8")

try:
    import openpyxl
except ImportError:
    print("Installing openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl --break-system-packages -q")
    import openpyxl

# ─────────────────────────────────────────────
# COLUMN CONFIGURATION
# Based on Limelight export format (columns are fixed by template)
# Colors in Row 1 of the source file:
#   RED   (s=3) = delete column entirely
#   GREEN (s=4) = keep, rename header to Row 2 value
#   BLUE  (s=5) = keep as-is with original header
# ─────────────────────────────────────────────

# 0-indexed column positions in the raw Limelight file
RED_COLS = {0, 3, 4, 5, 9, 11, 16, 17, 18, 19, 20, 21, 28}
# A=0  D=3  E=4  F=5  J=9  L=11  Q=16  R=17  S=18  T=19  U=20  V=21  AC=28

GREEN_RENAME = {
    1:  "First Name",              # B: firstName
    2:  "Last Name",               # C: lastName
    6:  "Job Title",               # G: title1
    7:  "LinkedIn URL-Contact",    # H: linkedinProfileUrl
    8:  "Company Name",            # I: companyName
    10: "Employees- Do Not Use",   # K: staffCount
    12: "Website URL",             # M: companyHomepageUrl
    13: "Industry",                # N: companyIndustry
    14: "City",                    # O: companyCity
    15: "Country/Region",          # P: companyCountry
    22: "Event Task Notes",        # W: engmt1Post
    29: "Salesforce Campaign ID",  # AD
}

# BLUE cols kept with original header (X=23, Y=24, Z=25, AA=26, AB=27)
BLUE_COLS = {23, 24, 25, 26, 27}

# ─────────────────────────────────────────────
# COUNTRY CODE → FULL NAME
# ─────────────────────────────────────────────
COUNTRY_MAP = {
    "AE": "United Arab Emirates",
    "AF": "Afghanistan",
    "AL": "Albania",
    "AM": "Armenia",
    "AO": "Angola",
    "AR": "Argentina",
    "AT": "Austria",
    "AU": "Australia",
    "AZ": "Azerbaijan",
    "BA": "Bosnia and Herzegovina",
    "BD": "Bangladesh",
    "BE": "Belgium",
    "BF": "Burkina Faso",
    "BG": "Bulgaria",
    "BH": "Bahrain",
    "BJ": "Benin",
    "BO": "Bolivia",
    "BR": "Brazil",
    "BY": "Belarus",
    "CA": "Canada",
    "CD": "Democratic Republic of the Congo",
    "CH": "Switzerland",
    "CI": "Ivory Coast",
    "CL": "Chile",
    "CM": "Cameroon",
    "CN": "China",
    "CO": "Colombia",
    "CR": "Costa Rica",
    "CY": "Cyprus",
    "CZ": "Czech Republic",
    "DE": "Germany",
    "DK": "Denmark",
    "DO": "Dominican Republic",
    "DZ": "Algeria",
    "EC": "Ecuador",
    "EE": "Estonia",
    "EG": "Egypt",
    "ES": "Spain",
    "ET": "Ethiopia",
    "FI": "Finland",
    "FR": "France",
    "GA": "Gabon",
    "GB": "United Kingdom",
    "GE": "Georgia",
    "GH": "Ghana",
    "GR": "Greece",
    "GT": "Guatemala",
    "HK": "Hong Kong",
    "HN": "Honduras",
    "HR": "Croatia",
    "HU": "Hungary",
    "ID": "Indonesia",
    "IE": "Ireland",
    "IL": "Israel",
    "IN": "India",
    "IQ": "Iraq",
    "IR": "Iran",
    "IS": "Iceland",
    "IT": "Italy",
    "JM": "Jamaica",
    "JO": "Jordan",
    "JP": "Japan",
    "KE": "Kenya",
    "KG": "Kyrgyzstan",
    "KH": "Cambodia",
    "KR": "South Korea",
    "KW": "Kuwait",
    "KZ": "Kazakhstan",
    "LB": "Lebanon",
    "LK": "Sri Lanka",
    "LT": "Lithuania",
    "LU": "Luxembourg",
    "LV": "Latvia",
    "LY": "Libya",
    "MA": "Morocco",
    "MD": "Moldova",
    "MK": "North Macedonia",
    "ML": "Mali",
    "MM": "Myanmar",
    "MN": "Mongolia",
    "MX": "Mexico",
    "MY": "Malaysia",
    "MZ": "Mozambique",
    "NA": "Namibia",
    "NG": "Nigeria",
    "NI": "Nicaragua",
    "NL": "Netherlands",
    "NO": "Norway",
    "NP": "Nepal",
    "NZ": "New Zealand",
    "OM": "Oman",
    "PA": "Panama",
    "PE": "Peru",
    "PH": "Philippines",
    "PK": "Pakistan",
    "PL": "Poland",
    "PR": "Puerto Rico",
    "PT": "Portugal",
    "PY": "Paraguay",
    "QA": "Qatar",
    "RO": "Romania",
    "RS": "Serbia",
    "RU": "Russia",
    "RW": "Rwanda",
    "SA": "Saudi Arabia",
    "SD": "Sudan",
    "SE": "Sweden",
    "SG": "Singapore",
    "SI": "Slovenia",
    "SK": "Slovakia",
    "SN": "Senegal",
    "SO": "Somalia",
    "SV": "El Salvador",
    "SY": "Syria",
    "TH": "Thailand",
    "TN": "Tunisia",
    "TR": "Turkey",
    "TW": "Taiwan",
    "TZ": "Tanzania",
    "UA": "Ukraine",
    "UG": "Uganda",
    "US": "United States",
    "UY": "Uruguay",
    "UZ": "Uzbekistan",
    "VE": "Venezuela",
    "VN": "Vietnam",
    "YE": "Yemen",
    "ZA": "South Africa",
    "ZM": "Zambia",
    "ZW": "Zimbabwe",
}

# ─────────────────────────────────────────────
# CITY → STATE LOOKUP  (US only, high-confidence)
# Only includes cities where the name unambiguously maps to one state.
# Ambiguous city names (e.g. "Springfield") are intentionally omitted.
# ─────────────────────────────────────────────
US_CITY_STATE = {
    # Alabama
    "Birmingham": "Alabama", "Montgomery": "Alabama", "Huntsville": "Alabama",
    "Mobile": "Alabama", "Tuscaloosa": "Alabama",
    # Alaska
    "Anchorage": "Alaska", "Fairbanks": "Alaska", "Juneau": "Alaska",
    # Arizona
    "Phoenix": "Arizona", "Tucson": "Arizona", "Mesa": "Arizona",
    "Chandler": "Arizona", "Scottsdale": "Arizona", "Gilbert": "Arizona",
    "Tempe": "Arizona", "Peoria": "Arizona", "Surprise": "Arizona",
    "Goodyear": "Arizona", "Avondale": "Arizona", "Flagstaff": "Arizona",
    # Arkansas
    "Little Rock": "Arkansas", "Fort Smith": "Arkansas", "Fayetteville": "Arkansas",
    # California
    "Los Angeles": "California", "San Diego": "California", "San Jose": "California",
    "San Francisco": "California", "Fresno": "California", "Sacramento": "California",
    "Long Beach": "California", "Oakland": "California", "Bakersfield": "California",
    "Anaheim": "California", "Santa Ana": "California", "Riverside": "California",
    "Stockton": "California", "Chula Vista": "California", "Irvine": "California",
    "Fremont": "California", "San Bernardino": "California", "Modesto": "California",
    "Fontana": "California", "Moreno Valley": "California", "Glendale": "California",
    "Huntington Beach": "California", "Santa Clarita": "California",
    "Garden Grove": "California", "Oceanside": "California", "Rancho Cucamonga": "California",
    "Santa Rosa": "California", "Ontario": "California", "Elk Grove": "California",
    "Corona": "California", "Hayward": "California", "Lancaster": "California",
    "Salinas": "California", "Palmdale": "California", "Sunnyvale": "California",
    "Pomona": "California", "Escondido": "California", "Torrance": "California",
    "Pasadena": "California", "Orange": "California", "Fullerton": "California",
    "Roseville": "California", "Visalia": "California", "Concord": "California",
    "Thousand Oaks": "California", "Simi Valley": "California",
    "Santa Clara": "California", "Vallejo": "California", "Berkeley": "California",
    "El Monte": "California", "Downey": "California", "Costa Mesa": "California",
    "Inglewood": "California", "Burbank": "California", "West Covina": "California",
    "Norwalk": "California", "Ventura": "California", "San Mateo": "California",
    "Murrieta": "California", "Temecula": "California", "Antioch": "California",
    "Richmond": "California", "Santa Barbara": "California",
    "Daly City": "California", "San Leandro": "California",
    "Redding": "California", "West Sacramento": "California",
    "San Marcos": "California", "Jurupa Valley": "California",
    "Vista": "California", "Chico": "California", "Alhambra": "California",
    "Menifee": "California", "Compton": "California", "Clovis": "California",
    "El Cajon": "California", "San Buenaventura": "California",
    "Clearfield": "California", "Fairfield": "California",
    "Redwood City": "California", "Carlsbad": "California",
    "Oxnard": "California", "Rialto": "California",
    "Mountain View": "California", "Palo Alto": "California",
    "Santa Monica": "California", "Beverly Hills": "California",
    "Cupertino": "California", "Redondo Beach": "California",
    "San Ramon": "California", "Pleasanton": "California",
    "Livermore": "California", "Folsom": "California",
    "Thousand Oaks": "California", "Napa": "California",
    "Petaluma": "California", "San Luis Obispo": "California",
    # Colorado
    "Denver": "Colorado", "Colorado Springs": "Colorado", "Aurora": "Colorado",
    "Fort Collins": "Colorado", "Lakewood": "Colorado", "Thornton": "Colorado",
    "Arvada": "Colorado", "Westminster": "Colorado", "Pueblo": "Colorado",
    "Centennial": "Colorado", "Boulder": "Colorado", "Highlands Ranch": "Colorado",
    "Greeley": "Colorado", "Longmont": "Colorado", "Loveland": "Colorado",
    "Broomfield": "Colorado", "Castle Rock": "Colorado", "Parker": "Colorado",
    # Connecticut
    "Bridgeport": "Connecticut", "New Haven": "Connecticut", "Stamford": "Connecticut",
    "Hartford": "Connecticut", "Waterbury": "Connecticut", "Norwalk": "Connecticut",
    "Danbury": "Connecticut", "Greenwich": "Connecticut",
    # Delaware
    "Wilmington": "Delaware", "Dover": "Delaware", "Newark": "Delaware",
    # Florida
    "Jacksonville": "Florida", "Miami": "Florida", "Tampa": "Florida",
    "Orlando": "Florida", "St. Petersburg": "Florida", "Hialeah": "Florida",
    "Tallahassee": "Florida", "Fort Lauderdale": "Florida", "Port St. Lucie": "Florida",
    "Cape Coral": "Florida", "Pembroke Pines": "Florida", "Hollywood": "Florida",
    "Miramar": "Florida", "Gainesville": "Florida", "Coral Springs": "Florida",
    "Clearwater": "Florida", "Miami Gardens": "Florida", "Palm Bay": "Florida",
    "West Palm Beach": "Florida", "Pompano Beach": "Florida", "Lakeland": "Florida",
    "Davie": "Florida", "Miami Beach": "Florida", "Sunrise": "Florida",
    "Boca Raton": "Florida", "Deltona": "Florida", "Plantation": "Florida",
    "Deerfield Beach": "Florida", "Boynton Beach": "Florida",
    "Lauderhill": "Florida", "Weston": "Florida", "Daytona Beach": "Florida",
    "Fort Myers": "Florida", "Homestead": "Florida", "Kissimmee": "Florida",
    "Clearwater": "Florida", "Naples": "Florida", "Sarasota": "Florida",
    "Pensacola": "Florida", "St. Augustine": "Florida",
    # Georgia
    "Atlanta": "Georgia", "Augusta": "Georgia", "Columbus": "Georgia",
    "Savannah": "Georgia", "Athens": "Georgia", "Sandy Springs": "Georgia",
    "Roswell": "Georgia", "Macon": "Georgia", "Johns Creek": "Georgia",
    "Albany": "Georgia", "Warner Robins": "Georgia", "Alpharetta": "Georgia",
    "Marietta": "Georgia", "Smyrna": "Georgia", "Dunwoody": "Georgia",
    # Hawaii
    "Honolulu": "Hawaii", "Pearl City": "Hawaii", "Hilo": "Hawaii", "Kailua": "Hawaii",
    # Idaho
    "Boise": "Idaho", "Nampa": "Idaho", "Meridian": "Idaho", "Idaho Falls": "Idaho",
    "Pocatello": "Idaho", "Coeur d'Alene": "Idaho",
    # Illinois
    "Chicago": "Illinois", "Aurora": "Illinois", "Joliet": "Illinois",
    "Naperville": "Illinois", "Rockford": "Illinois", "Elgin": "Illinois",
    "Peoria": "Illinois", "Champaign": "Illinois", "Waukegan": "Illinois",
    "Cicero": "Illinois", "Bloomington": "Illinois", "Evanston": "Illinois",
    "Decatur": "Illinois", "Schaumburg": "Illinois", "Bolingbrook": "Illinois",
    # Indiana
    "Indianapolis": "Indiana", "Fort Wayne": "Indiana", "Evansville": "Indiana",
    "South Bend": "Indiana", "Carmel": "Indiana", "Fishers": "Indiana",
    "Hammond": "Indiana", "Gary": "Indiana", "Muncie": "Indiana",
    "Terre Haute": "Indiana", "Lafayette": "Indiana",
    # Iowa
    "Des Moines": "Iowa", "Cedar Rapids": "Iowa", "Davenport": "Iowa",
    "Sioux City": "Iowa", "Iowa City": "Iowa", "Waterloo": "Iowa",
    # Kansas
    "Wichita": "Kansas", "Overland Park": "Kansas", "Kansas City": "Kansas",
    "Topeka": "Kansas", "Olathe": "Kansas", "Lawrence": "Kansas",
    "Shawnee": "Kansas", "Lenexa": "Kansas",
    # Kentucky
    "Louisville": "Kentucky", "Lexington": "Kentucky", "Bowling Green": "Kentucky",
    "Owensboro": "Kentucky", "Covington": "Kentucky",
    # Louisiana
    "New Orleans": "Louisiana", "Baton Rouge": "Louisiana", "Shreveport": "Louisiana",
    "Metairie": "Louisiana", "Lafayette": "Louisiana",
    # Maine
    "Portland": "Maine", "Lewiston": "Maine", "Bangor": "Maine",
    # Maryland
    "Baltimore": "Maryland", "Frederick": "Maryland", "Rockville": "Maryland",
    "Gaithersburg": "Maryland", "Bowie": "Maryland", "Annapolis": "Maryland",
    "College Park": "Maryland", "Bethesda": "Maryland",
    # Massachusetts
    "Boston": "Massachusetts", "Worcester": "Massachusetts", "Springfield": "Massachusetts",
    "Lowell": "Massachusetts", "Cambridge": "Massachusetts", "New Bedford": "Massachusetts",
    "Brockton": "Massachusetts", "Quincy": "Massachusetts", "Lynn": "Massachusetts",
    "Fall River": "Massachusetts", "Newton": "Massachusetts", "Somerville": "Massachusetts",
    "Lawrence": "Massachusetts", "Waltham": "Massachusetts", "Haverhill": "Massachusetts",
    # Michigan
    "Detroit": "Michigan", "Grand Rapids": "Michigan", "Warren": "Michigan",
    "Sterling Heights": "Michigan", "Ann Arbor": "Michigan", "Lansing": "Michigan",
    "Flint": "Michigan", "Dearborn": "Michigan", "Livonia": "Michigan",
    "Westland": "Michigan", "Troy": "Michigan", "Farmington Hills": "Michigan",
    "Kalamazoo": "Michigan", "Wyoming": "Michigan", "Southfield": "Michigan",
    "Rochester Hills": "Michigan", "Taylor": "Michigan", "Pontiac": "Michigan",
    "Novi": "Michigan", "Dearborn Heights": "Michigan",
    # Minnesota
    "Minneapolis": "Minnesota", "Saint Paul": "Minnesota", "St. Paul": "Minnesota",
    "Rochester": "Minnesota", "Duluth": "Minnesota", "Bloomington": "Minnesota",
    "Brooklyn Park": "Minnesota", "Plymouth": "Minnesota", "St. Cloud": "Minnesota",
    "Eagan": "Minnesota", "Woodbury": "Minnesota", "Burnsville": "Minnesota",
    "Eden Prairie": "Minnesota", "Coon Rapids": "Minnesota",
    # Mississippi
    "Jackson": "Mississippi", "Gulfport": "Mississippi", "Southaven": "Mississippi",
    "Hattiesburg": "Mississippi", "Biloxi": "Mississippi",
    # Missouri
    "Kansas City": "Missouri", "St. Louis": "Missouri", "Springfield": "Missouri",
    "Columbia": "Missouri", "Independence": "Missouri", "Lee's Summit": "Missouri",
    "O'Fallon": "Missouri", "St. Joseph": "Missouri", "St. Charles": "Missouri",
    # Montana
    "Billings": "Montana", "Missoula": "Montana", "Great Falls": "Montana",
    # Nebraska
    "Omaha": "Nebraska", "Lincoln": "Nebraska", "Bellevue": "Nebraska",
    # Nevada
    "Las Vegas": "Nevada", "Henderson": "Nevada", "Reno": "Nevada",
    "North Las Vegas": "Nevada", "Sparks": "Nevada", "Carson City": "Nevada",
    # New Hampshire
    "Manchester": "New Hampshire", "Nashua": "New Hampshire", "Concord": "New Hampshire",
    # New Jersey
    "Newark": "New Jersey", "Jersey City": "New Jersey", "Paterson": "New Jersey",
    "Elizabeth": "New Jersey", "Trenton": "New Jersey", "Clifton": "New Jersey",
    "Camden": "New Jersey", "Passaic": "New Jersey", "Hackensack": "New Jersey",
    "Vineland": "New Jersey", "Bayonne": "New Jersey", "East Orange": "New Jersey",
    "Union City": "New Jersey", "New Brunswick": "New Jersey",
    "Iselin": "New Jersey", "Edison": "New Jersey", "Parsippany": "New Jersey",
    "Princeton": "New Jersey", "Hoboken": "New Jersey", "Cherry Hill": "New Jersey",
    "Morristown": "New Jersey", "Fort Lee": "New Jersey",
    # New Mexico
    "Albuquerque": "New Mexico", "Las Cruces": "New Mexico", "Rio Rancho": "New Mexico",
    "Santa Fe": "New Mexico",
    # New York
    "New York": "New York", "New York City": "New York", "Buffalo": "New York",
    "Rochester": "New York", "Yonkers": "New York", "Syracuse": "New York",
    "Albany": "New York", "New Rochelle": "New York", "Mount Vernon": "New York",
    "Schenectady": "New York", "Utica": "New York", "White Plains": "New York",
    "Hempstead": "New York", "Troy": "New York", "Niagara Falls": "New York",
    "Binghamton": "New York", "Manhattan": "New York", "Brooklyn": "New York",
    "Queens": "New York", "Staten Island": "New York", "Bronx": "New York",
    # North Carolina
    "Charlotte": "North Carolina", "Raleigh": "North Carolina",
    "Greensboro": "North Carolina", "Durham": "North Carolina",
    "Winston-Salem": "North Carolina", "Fayetteville": "North Carolina",
    "Cary": "North Carolina", "Wilmington": "North Carolina",
    "High Point": "North Carolina", "Concord": "North Carolina",
    "Asheville": "North Carolina", "Chapel Hill": "North Carolina",
    # North Dakota
    "Fargo": "North Dakota", "Bismarck": "North Dakota", "Grand Forks": "North Dakota",
    # Ohio
    "Columbus": "Ohio", "Cleveland": "Ohio", "Cincinnati": "Ohio",
    "Toledo": "Ohio", "Akron": "Ohio", "Dayton": "Ohio",
    "Parma": "Ohio", "Canton": "Ohio", "Youngstown": "Ohio",
    "Lorain": "Ohio", "Hamilton": "Ohio", "Dublin": "Ohio",
    "Kettering": "Ohio", "Elyria": "Ohio",
    # Oklahoma
    "Oklahoma City": "Oklahoma", "Tulsa": "Oklahoma", "Norman": "Oklahoma",
    "Broken Arrow": "Oklahoma", "Lawton": "Oklahoma", "Edmond": "Oklahoma",
    # Oregon
    "Portland": "Oregon", "Eugene": "Oregon", "Salem": "Oregon",
    "Gresham": "Oregon", "Hillsboro": "Oregon", "Beaverton": "Oregon",
    "Bend": "Oregon", "Medford": "Oregon", "Springfield": "Oregon",
    # Pennsylvania
    "Philadelphia": "Pennsylvania", "Pittsburgh": "Pennsylvania", "Allentown": "Pennsylvania",
    "Erie": "Pennsylvania", "Reading": "Pennsylvania", "Scranton": "Pennsylvania",
    "Bethlehem": "Pennsylvania", "Lancaster": "Pennsylvania", "Harrisburg": "Pennsylvania",
    "Altoona": "Pennsylvania", "York": "Pennsylvania", "Wilkes-Barre": "Pennsylvania",
    "King of Prussia": "Pennsylvania", "Wayne": "Pennsylvania",
    # Rhode Island
    "Providence": "Rhode Island", "Warwick": "Rhode Island", "Cranston": "Rhode Island",
    # South Carolina
    "Columbia": "South Carolina", "Charleston": "South Carolina",
    "North Charleston": "South Carolina", "Mount Pleasant": "South Carolina",
    "Rock Hill": "South Carolina", "Greenville": "South Carolina",
    # South Dakota
    "Sioux Falls": "South Dakota", "Rapid City": "South Dakota",
    # Tennessee
    "Nashville": "Tennessee", "Memphis": "Tennessee", "Knoxville": "Tennessee",
    "Chattanooga": "Tennessee", "Clarksville": "Tennessee", "Murfreesboro": "Tennessee",
    "Franklin": "Tennessee", "Jackson": "Tennessee", "Johnson City": "Tennessee",
    "Bartlett": "Tennessee", "Hendersonville": "Tennessee",
    # Texas
    "Houston": "Texas", "San Antonio": "Texas", "Dallas": "Texas",
    "Austin": "Texas", "Fort Worth": "Texas", "El Paso": "Texas",
    "Arlington": "Texas", "Corpus Christi": "Texas", "Plano": "Texas",
    "Laredo": "Texas", "Lubbock": "Texas", "Garland": "Texas",
    "Irving": "Texas", "Amarillo": "Texas", "Grand Prairie": "Texas",
    "Brownsville": "Texas", "McKinney": "Texas", "Frisco": "Texas",
    "Pasadena": "Texas", "Mesquite": "Texas", "Killeen": "Texas",
    "McAllen": "Texas", "Midland": "Texas", "Denton": "Texas",
    "Waco": "Texas", "Carrollton": "Texas", "Beaumont": "Texas",
    "Abilene": "Texas", "Odessa": "Texas", "Round Rock": "Texas",
    "Pearland": "Texas", "Richardson": "Texas", "Lewisville": "Texas",
    "Tyler": "Texas", "Allen": "Texas", "College Station": "Texas",
    "Wichita Falls": "Texas", "Sugar Land": "Texas", "Edinburg": "Texas",
    "The Woodlands": "Texas", "Cedar Park": "Texas", "Georgetown": "Texas",
    "Flower Mound": "Texas", "League City": "Texas", "Katy": "Texas",
    # Utah
    "Salt Lake City": "Utah", "West Valley City": "Utah", "Provo": "Utah",
    "West Jordan": "Utah", "Orem": "Utah", "Sandy": "Utah",
    "Ogden": "Utah", "St. George": "Utah", "Layton": "Utah",
    "Taylorsville": "Utah", "South Jordan": "Utah",
    # Vermont
    "Burlington": "Vermont", "Essex": "Vermont", "South Burlington": "Vermont",
    # Virginia
    "Virginia Beach": "Virginia", "Norfolk": "Virginia", "Chesapeake": "Virginia",
    "Richmond": "Virginia", "Newport News": "Virginia", "Alexandria": "Virginia",
    "Hampton": "Virginia", "Roanoke": "Virginia", "Portsmouth": "Virginia",
    "Suffolk": "Virginia", "Lynchburg": "Virginia", "Harrisonburg": "Virginia",
    "Arlington": "Virginia", "McLean": "Virginia", "Tysons": "Virginia",
    "Herndon": "Virginia", "Reston": "Virginia", "Vienna": "Virginia",
    # Washington
    "Seattle": "Washington", "Spokane": "Washington", "Tacoma": "Washington",
    "Vancouver": "Washington", "Bellevue": "Washington", "Kent": "Washington",
    "Everett": "Washington", "Renton": "Washington", "Spokane Valley": "Washington",
    "Kirkland": "Washington", "Bellingham": "Washington", "Kennewick": "Washington",
    "Federal Way": "Washington", "Yakima": "Washington", "Redmond": "Washington",
    "Marysville": "Washington", "Shoreline": "Washington", "Richland": "Washington",
    "Sammamish": "Washington", "Lacey": "Washington", "Bothell": "Washington",
    "Auburn": "Washington",
    # West Virginia
    "Charleston": "West Virginia", "Huntington": "West Virginia",
    "Morgantown": "West Virginia",
    # Wisconsin
    "Milwaukee": "Wisconsin", "Madison": "Wisconsin", "Green Bay": "Wisconsin",
    "Kenosha": "Wisconsin", "Racine": "Wisconsin", "Appleton": "Wisconsin",
    "Waukesha": "Wisconsin", "Oshkosh": "Wisconsin",
    # Wyoming
    "Cheyenne": "Wyoming", "Casper": "Wyoming",
    # DC
    "Washington": "District of Columbia", "Washington D.C.": "District of Columbia",
    "Washington DC": "District of Columbia",
}

# ─────────────────────────────────────────────
# CANADA: CITY → PROVINCE
# ─────────────────────────────────────────────
CA_CITY_PROVINCE = {
    # Alberta
    "Calgary": "Alberta", "Edmonton": "Alberta", "Red Deer": "Alberta",
    "Lethbridge": "Alberta", "St. Albert": "Alberta", "Medicine Hat": "Alberta",
    "Grande Prairie": "Alberta", "Airdrie": "Alberta", "Spruce Grove": "Alberta",
    # British Columbia
    "Vancouver": "British Columbia", "Surrey": "British Columbia",
    "Burnaby": "British Columbia", "Richmond": "British Columbia",
    "Kelowna": "British Columbia", "Abbotsford": "British Columbia",
    "Coquitlam": "British Columbia", "Langley": "British Columbia",
    "Saanich": "British Columbia", "Delta": "British Columbia",
    "Victoria": "British Columbia", "Nanaimo": "British Columbia",
    "Kamloops": "British Columbia", "Chilliwack": "British Columbia",
    "Port Coquitlam": "British Columbia", "North Vancouver": "British Columbia",
    "Prince George": "British Columbia",
    # Manitoba
    "Winnipeg": "Manitoba", "Brandon": "Manitoba",
    # New Brunswick
    "Moncton": "New Brunswick", "Fredericton": "New Brunswick",
    "Saint John": "New Brunswick",
    # Newfoundland and Labrador
    "St. John's": "Newfoundland and Labrador",
    # Nova Scotia
    "Halifax": "Nova Scotia", "Sydney": "Nova Scotia",
    # Ontario
    "Toronto": "Ontario", "Ottawa": "Ontario", "Mississauga": "Ontario",
    "Brampton": "Ontario", "Hamilton": "Ontario", "London": "Ontario",
    "Markham": "Ontario", "Vaughan": "Ontario", "Kitchener": "Ontario",
    "Windsor": "Ontario", "Oakville": "Ontario", "Richmond Hill": "Ontario",
    "Burlington": "Ontario", "Oshawa": "Ontario", "Barrie": "Ontario",
    "Waterloo": "Ontario", "St. Catharines": "Ontario", "Cambridge": "Ontario",
    "Kingston": "Ontario", "Guelph": "Ontario", "Thunder Bay": "Ontario",
    "Ajax": "Ontario", "Whitby": "Ontario", "Pickering": "Ontario",
    "Newmarket": "Ontario", "Sudbury": "Ontario", "Aurora": "Ontario",
    "Mississauga": "Ontario", "Etobicoke": "Ontario", "Scarborough": "Ontario",
    "North York": "Ontario", "Thornhill": "Ontario",
    # Prince Edward Island
    "Charlottetown": "Prince Edward Island",
    # Quebec
    "Montreal": "Quebec", "Quebec City": "Quebec", "Laval": "Quebec",
    "Gatineau": "Quebec", "Longueuil": "Quebec", "Sherbrooke": "Quebec",
    "Saguenay": "Quebec", "Levis": "Quebec", "Trois-Rivieres": "Quebec",
    "Terrebonne": "Quebec",
    # Saskatchewan
    "Saskatoon": "Saskatchewan", "Regina": "Saskatchewan",
}

# ─────────────────────────────────────────────
# COMPETITOR FILTER
# Records where Company Name (col I, index 8) contains any of these strings
# (case-insensitive) will be removed from the output entirely.
# ─────────────────────────────────────────────
COMPETITOR_NAMES = {"omnissa", "citrix", "hydra"}

# Cities that appear in multiple US states — exclude from lookup to stay ≥98% certain
AMBIGUOUS_US_CITIES = {
    "Springfield", "Portland", "Columbus", "Jackson", "Madison", "Jefferson",
    "Franklin", "Georgetown", "Clinton", "Manchester", "Lexington", "Richmond",
    "Burlington", "Lawrence", "Greenville", "Salem", "Aurora", "Bristol",
    "Chester", "Florence", "Milford", "Monroe", "Troy", "Camden", "Concord",
    "Dover", "Hudson", "Lincoln", "Milford", "Riverside", "Washington",
    "Arlington", "Clayton", "Winchester", "Fairfield", "Oakland",
}

# ─────────────────────────────────────────────
# INDUSTRY NORMALIZATION
# ─────────────────────────────────────────────
def load_industry_map():
    """Load industry mapping from the bundled references/industry_mapping.json"""
    script_dir = Path(__file__).parent
    mapping_file = script_dir.parent / "references" / "industry_mapping.json"
    with open(mapping_file, "r") as f:
        raw = json.load(f)
    # Invert: "source value" → "Salesforce picklist value"
    inverted = {}
    for sf_value, source_list in raw.items():
        for src in source_list:
            inverted[src.lower().strip()] = sf_value
    return inverted

def normalize_industry(value, industry_map):
    """Map a Limelight industry value to a Salesforce picklist value."""
    if not value:
        return value
    key = str(value).lower().strip()
    # Strip zero-width spaces and other invisible characters
    key = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', key)
    if key in industry_map:
        return industry_map[key]
    # Try partial match as fallback
    for src_key, sf_val in industry_map.items():
        if src_key in key or key in src_key:
            return sf_val
    # Return original if no match found (better than losing data)
    return value

# ─────────────────────────────────────────────
# GEO NORMALIZATION
# ─────────────────────────────────────────────
def expand_country(code):
    """Expand a 2-char ISO country code to full country name."""
    if not code:
        return code
    code_upper = str(code).strip().upper()
    return COUNTRY_MAP.get(code_upper, code)  # Return original if not found

def lookup_state(city, country):
    """
    Look up state/province for a city. Returns the state name if ≥98% certain,
    otherwise returns empty string.
    """
    if not city or not country:
        return ""
    city_clean = str(city).strip()
    country_clean = str(country).strip()

    if country_clean == "United States":
        # Skip known ambiguous city names
        if city_clean in AMBIGUOUS_US_CITIES:
            return ""
        return US_CITY_STATE.get(city_clean, "")

    elif country_clean == "Canada":
        return CA_CITY_PROVINCE.get(city_clean, "")

    return ""

# ─────────────────────────────────────────────
# CAMPAIGN ID LOGIC
# ─────────────────────────────────────────────
def load_campaign_mapping():
    """Load campaign ID rules from references/campaign_mapping.json."""
    script_dir = Path(__file__).parent
    mapping_file = script_dir.parent / "references" / "campaign_mapping.json"
    with open(mapping_file, "r") as f:
        data = json.load(f)
    # Remove the _notes meta key
    return {k: v for k, v in data.items() if not k.startswith("_")}

def detect_file_type(filename, campaign_mapping):
    """
    Match a filename against the known Limelight file type patterns.
    Strips extension and trailing date digits before matching.
    Normalizes hyphens, underscores, spaces and case so
    'LImelight-CompMentions_20260424' or 'Limelight CompMentions 20260424'
    all match the key 'Limelight_CompMentions'.
    Returns the matching key from campaign_mapping, or None if unrecognized.
    """
    def normalize(s):
        return re.sub(r'[-_ ]', '_', s).lower()  # also handle spaces

    stem = Path(filename).stem                          # drop .xlsx
    stem = re.sub(r'[_\- ]?\d{6,8}$', '', stem)        # strip trailing date like _20260424
    norm_stem = normalize(stem)

    # Exact match first (fastest, most precise)
    for key in campaign_mapping:
        if norm_stem == normalize(key):
            return key

    # Prefix match: handle filenames with extra suffixes
    # e.g. "Limelight_Nerdio_Staff" → matches "Limelight_Nerdio"
    # Longest key wins to avoid "Limelight_Nerdio" beating "Limelight_NerdioPage".
    best_key, best_len = None, 0
    for key in campaign_mapping:
        norm_key = normalize(key)
        if norm_stem.startswith(norm_key) and len(norm_key) > best_len:
            best_key, best_len = key, len(norm_key)
    if best_key:
        return best_key

    # Substring (contains) match: handles "LLT-Export-getnerdiocom_<phrase>.csv" style names
    # where the descriptive phrase appears anywhere in the filename.
    # Longest matching key wins so more-specific keys beat shorter ones.
    best_key, best_len = None, 0
    for key in campaign_mapping:
        norm_key = normalize(key)
        if norm_key in norm_stem and len(norm_key) > best_len:
            best_key, best_len = key, len(norm_key)
    return best_key

def get_campaign_id(file_type, author, campaign_mapping):
    """
    Return the Salesforce Campaign ID for a given file type and post author.
    For by_author file types, matches author name case-insensitively as a substring.
    Falls back to _default if no specific author match is found.
    """
    if file_type is None:
        return ""
    rules = campaign_mapping.get(file_type, {})

    if "campaign_id" in rules:
        # Flat mapping — same ID regardless of author
        return rules["campaign_id"]

    if "by_author" in rules:
        author_str = str(author).strip() if author else ""
        author_lower = author_str.lower()
        by_author = rules["by_author"]
        for name, campaign_id in by_author.items():
            if name.startswith("_"):
                continue  # skip _default here
            if name.lower() in author_lower:
                return campaign_id
        # No specific match — use catch-all
        return by_author.get("_default", "")

    return ""

# ─────────────────────────────────────────────
# SINGLE-FILE PROCESSING
# ─────────────────────────────────────────────
def normalize_limelight(input_path, output_path=None, campaign_mapping=None, campaign_type_override=None):
    """
    Normalize one Limelight file. Returns (processed_rows, final_headers, stats, filtered_out).
    Caller is responsible for writing output files (supports multi-file combining).
    """
    if campaign_mapping is None:
        campaign_mapping = load_campaign_mapping()

    filename = Path(input_path).name
    if campaign_type_override is not None:
        file_type = campaign_type_override
        print(f"   File type set manually: {file_type}")
    else:
        file_type = detect_file_type(filename, campaign_mapping)
        if file_type:
            print(f"   File type detected: {file_type}")
        else:
            print(f"   ⚠️  File type not recognized from filename — Salesforce Campaign ID will be blank.")

    print(f"Reading: {input_path}")
    ext = Path(input_path).suffix.lower()
    all_data = []

    # ── Step 1: Read all data into memory ──
    if ext == '.csv':
        with open(input_path, newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                # CSV has 29 columns (0-28); XLSX has 30 (0-29).
                # Col 29 is the Salesforce Campaign ID slot — pad so it can be written.
                while len(row) < 30:
                    row.append("")
                all_data.append(list(row))
    else:
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            all_data.append(list(row))

    # Row index 0 = Row 1 (original headers)
    # Row index 1 = Row 2 (rename guide — will be dropped, or first data row)
    # Row index 2+ = Data rows

    # ── Step 2: Remove competitor records ──
    # Column I (index 8) = companyName. Drop any row where it contains a competitor name.
    COMPANY_COL = 8
    pre_filter_count = len(all_data) - 1  # rough count before we know data_start
    filtered_out = []
    cleaned_data = [all_data[0]]  # always keep the header row
    for row in all_data[1:]:
        company = str(row[COMPANY_COL]).strip().lower() if row[COMPANY_COL] else ""
        if any(competitor in company for competitor in COMPETITOR_NAMES):
            filtered_out.append(row[COMPANY_COL])
        else:
            cleaned_data.append(row)
    if filtered_out:
        print(f"   ⚠️  Removed {len(filtered_out)} competitor record(s): {filtered_out}")
    all_data = cleaned_data

    # ── Step 4: Determine final column set ──
    # Keep GREEN and BLUE cols; drop RED cols
    keep_cols = sorted(
        list(GREEN_RENAME.keys()) + list(BLUE_COLS)
    )

    # Build the new header row
    # - GREEN cols: use the rename value
    # - BLUE cols: use original header from Row 1
    new_headers = []
    for col_idx in keep_cols:
        if col_idx in GREEN_RENAME:
            new_headers.append(GREEN_RENAME[col_idx])
        else:
            # BLUE: use original header from Row 1
            val = all_data[0][col_idx] if col_idx < len(all_data[0]) else ""
            new_headers.append(val)

    # ── Step 3: Find positions of key columns in keep_cols ──
    keep_cols_list = keep_cols  # already sorted list
    try:
        industry_pos = keep_cols_list.index(13)   # N = companyIndustry
    except ValueError:
        industry_pos = None
    try:
        city_pos = keep_cols_list.index(14)        # O = companyCity
    except ValueError:
        city_pos = None
    try:
        country_pos = keep_cols_list.index(15)     # P = companyCountry
    except ValueError:
        country_pos = None

    # ── Step 4: Load industry mapping ──
    industry_map = load_industry_map()

    # ── Step 5: Detect and skip Row 2 if it's a legacy "friendly name" guide row ──
    # Older Limelight exports had a Row 2 with human-readable column labels.
    # Newer exports go straight to data in Row 2. We detect this by checking
    # whether Row 2 contains known guide-row values (e.g. "Industry", "City").
    guide_row_signals = {"industry", "city", "first name", "company name", "country/region"}
    data_start = 1  # default: data starts at Row 2 (index 1)
    if len(all_data) > 1:
        row2_values = {str(v).strip().lower() for v in all_data[1] if v}
        if len(row2_values & guide_row_signals) >= 2:
            data_start = 2  # legacy file — skip the guide row
            print("   (Detected legacy Row 2 guide row — skipping it)")

    # ── Step 6: Process data rows ──
    processed_rows = []
    stats = {"total": 0, "industry_mapped": 0, "industry_unchanged": 0,
             "state_filled": 0, "state_blank": 0, "country_expanded": 0}

    # Raw column index for engmt1PostAuthor (Y = index 24) — used for campaign ID lookup
    AUTHOR_RAW_COL = 24
    # Raw column index for Salesforce Campaign ID (AD = index 29)
    CAMPAIGN_RAW_COL = 29

    for row_data in all_data[data_start:]:
        stats["total"] += 1

        # ── Compute stable engagement key from raw columns B, C, X ──────────
        # B (1)  = First Name
        # C (2)  = Last Name
        # X (23) = engmt1PostPostedAt (post date/time)
        # Using contact name + post timestamp so the same engagement is deduped
        # across multiple campaign files (e.g. same person appears in both
        # "our-company-page-post-engagers" and "engagers-on-posts-mentioning-our-company").
        _first_name = str(row_data[1]  or "").strip().lower() if 1  < len(row_data) else ""
        _last_name  = str(row_data[2]  or "").strip().lower() if 2  < len(row_data) else ""
        _post_dt    = str(row_data[23] or "").strip()          if 23 < len(row_data) else ""
        _key = f"{_first_name}|{_last_name}|{_post_dt}"
        engagement_key = str(uuid.uuid5(_ENGAGEMENT_NS, _key)) if _key.strip("|") else ""

        # ── Assign Salesforce Campaign ID before column extraction ──
        # Read the post author from the raw row, then stamp the campaign ID
        raw_author = row_data[AUTHOR_RAW_COL] if AUTHOR_RAW_COL < len(row_data) else ""
        campaign_id = get_campaign_id(file_type, raw_author, campaign_mapping)
        if CAMPAIGN_RAW_COL < len(row_data):
            row_data[CAMPAIGN_RAW_COL] = campaign_id

        # Extract only the columns we're keeping
        new_row = []
        for col_idx in keep_cols_list:
            val = row_data[col_idx] if col_idx < len(row_data) else None
            new_row.append(val)

        # ── Normalize Industry ──
        if industry_pos is not None:
            original_industry = new_row[industry_pos]
            normalized = normalize_industry(original_industry, industry_map)
            new_row[industry_pos] = normalized
            if normalized != original_industry:
                stats["industry_mapped"] += 1
            else:
                stats["industry_unchanged"] += 1

        # ── Expand Country Code ──
        if country_pos is not None:
            raw_country = new_row[country_pos]
            expanded = expand_country(raw_country)
            new_row[country_pos] = expanded
            if expanded != raw_country:
                stats["country_expanded"] += 1

        # ── Insert State/Region for US and Canada ──
        state_val = ""
        if city_pos is not None and country_pos is not None:
            city = new_row[city_pos]
            country = new_row[country_pos]
            state_val = lookup_state(city, country)
            if state_val:
                stats["state_filled"] += 1
            elif country in ("United States", "Canada"):
                stats["state_blank"] += 1

        # ── Fill engmt1Type from engmt1 when blank ──────────────────────────
        # Col Z (engmt1, pos 13): "reaction" | "repost" | "comment"
        # Col AA (engmt1Type, pos 14): "LIKE" | "SHARE" | blank for repost/comment
        # Must run BEFORE the state insertion that shifts these indices by 1.
        _engmt1      = str(new_row[13] or "").strip().lower()
        _engmt1_type = str(new_row[14] or "").strip()
        if not _engmt1_type:
            if _engmt1 == "repost":
                new_row[14] = "REPOST"
            elif _engmt1 == "comment":
                new_row[14] = "COMMENT"

        # Insert State/Region just before Country/Region
        insert_pos = country_pos
        new_row.insert(insert_pos, state_val)
        new_row.append(filename)        # index 18: Import Batch
        new_row.append(engagement_key)  # index 19: Engagement Key (stable dedup ID)
        processed_rows.append(new_row)

    # ── Step 7: Build the final header row with State/Region inserted ──
    # Insert "State/Region" before "Country/Region" in headers
    final_headers = new_headers.copy()
    final_headers.insert(country_pos, "State/Region")

    # After insertion, the column positions in every processed row are:
    #  0  First Name
    #  1  Last Name
    #  2  Job Title
    #  3  LinkedIn URL-Contact        ← dedup key for contacts; VLOOKUP key for engagements
    #  4  Company Name
    #  5  Employees- Do Not Use
    #  6  Website URL
    #  7  Industry
    #  8  City
    #  9  State/Region
    # 10  Country/Region
    # 11  Event Task Notes            ← post URL
    # 12  engmt1PostPostedAt          ← timestamp
    # 13  engmt1PostAuthor
    # 14  engmt1                      ← reaction/comment
    # 15  engmt1Type                  ← LIKE / COMMENT etc.
    # 16  engmt1Text                  ← content
    # 17  Salesforce Campaign ID

    CONTACT_COLS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 17]
    ENGAGE_COLS  = [3, 11, 12, 13, 14, 15, 16, 17]  # used in write step

    stats["filtered_out"] = filtered_out
    stats["file_type"]    = file_type

    return processed_rows, final_headers, stats, CONTACT_COLS, ENGAGE_COLS


# ─────────────────────────────────────────────
# MULTI-FILE ORCHESTRATION + OUTPUT WRITING
# ─────────────────────────────────────────────
def write_outputs(all_processed_rows, final_headers, output_dir, label="SocialEngagement"):
    """
    Given processed rows (from one or more files, already combined), write:
      • _contacts_staging.xlsx   — deduplicated contacts, Email column empty (filled by ZoomInfo step)
      • _engagements_staging.xlsx — all engagement rows, Email placeholder empty
      • _to_enrich.json          — contact identity data for ZoomInfo enrichment calls
    Returns paths to staging files + the to_enrich list for reporting.
    """
    from datetime import datetime
    date_str = datetime.now().strftime("%Y%m%d")
    contacts_path    = str(Path(output_dir) / f"{label}_{date_str}_contacts_staging.xlsx")
    engagements_path = str(Path(output_dir) / f"{label}_{date_str}_engagements_staging.xlsx")
    to_enrich_path   = str(Path(output_dir) / f"{label}_{date_str}_to_enrich.json")

    CONTACT_COLS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 17]
    CONTACT_HDRS = [final_headers[i] for i in CONTACT_COLS]

    ENGAGE_HDRS = (
        [final_headers[3]]                                          # LinkedIn URL-Contact (VLOOKUP key)
        + ["Email"]                                                 # placeholder — fill via VLOOKUP after RingLead
        + [final_headers[0], final_headers[1], final_headers[4]]   # First Name, Last Name, Company
        + [final_headers[i] for i in [11, 12, 13, 14, 15, 16, 17]]
        + ["Import Batch"]                                          # source filename
        + ["Engagement ID"]                                          # Limelight record ID (dedup key)
    )

    # Contacts — deduplicated across ALL input files
    seen = set()
    contact_rows = []
    for row in all_processed_rows:
        linkedin_url = str(row[3]).strip() if row[3] else ""
        dedup_key = linkedin_url or f"{row[0]}|{row[1]}|{row[4]}"
        if dedup_key not in seen:
            seen.add(dedup_key)
            contact_rows.append([row[i] for i in CONTACT_COLS])

    contacts_wb = openpyxl.Workbook()
    contacts_ws = contacts_wb.active
    contacts_ws.title = "Contacts"
    contacts_ws.append(CONTACT_HDRS)
    for row in contact_rows:
        contacts_ws.append(row)
    contacts_wb.save(contacts_path)

    # Engagements — every row from every file, no dedup
    engage_wb = openpyxl.Workbook()
    engage_ws = engage_wb.active
    engage_ws.title = "Engagements"
    engage_ws.append(ENGAGE_HDRS)
    for row in all_processed_rows:
        engage_row = (
            [row[3]]                              # LinkedIn URL-Contact
            + [""]                                # Email placeholder — filled by finalize.py
            + [row[0], row[1], row[4]]            # First Name, Last Name, Company Name
            + [row[i] for i in [11, 12, 13, 14, 15, 16, 17]]
            + [row[18] if len(row) > 18 else ""]  # Import Batch (source filename)
            + [row[19] if len(row) > 19 else ""]  # Limelight ID (dedup key)
        )
        engage_ws.append(engage_row)
    engage_wb.save(engagements_path)

    # to_enrich.json — contact identity data for ZoomInfo enrichment
    # One entry per unique contact (same dedup set as contacts file)
    to_enrich = []
    for row in contact_rows:
        # contact_rows already uses CONTACT_COLS = [0,1,2,3,4,5,6,7,8,9,10,17]
        # so indices here: 0=FirstName, 1=LastName, 2=JobTitle, 3=LinkedIn, 4=Company
        to_enrich.append({
            "linkedin_url":  row[3] or "",
            "first_name":    row[0] or "",
            "last_name":     row[1] or "",
            "company_name":  row[4] or "",
        })
    with open(to_enrich_path, "w") as f:
        json.dump(to_enrich, f, indent=2)

    return contacts_path, engagements_path, to_enrich_path, contact_rows, CONTACT_HDRS, ENGAGE_HDRS


def run(input_files, output_dir):
    """
    Main entry point. Accepts one or more input file paths and an output directory.
    Processes each file, combines results, and writes the two output files.
    """
    campaign_mapping = load_campaign_mapping()
    all_rows    = []
    all_stats   = []
    final_hdrs  = None

    for input_path in input_files:
        print(f"\n{'─'*55}")
        processed_rows, final_headers, stats, _, _ = normalize_limelight(
            input_path, campaign_mapping=campaign_mapping
        )
        all_rows.extend(processed_rows)
        all_stats.append((Path(input_path).name, stats))
        if final_hdrs is None:
            final_hdrs = final_headers  # same across all files

    contacts_path, engagements_path, to_enrich_path, contact_rows, contact_hdrs, engage_hdrs = \
        write_outputs(all_rows, final_hdrs, output_dir)

    # ── Summary report ──
    print(f"\n{'═'*55}")
    print(f"✅ All files processed!\n")
    total_rows = sum(s["total"] for _, s in all_stats)
    total_filtered = sum(len(s["filtered_out"]) for _, s in all_stats)

    for fname, s in all_stats:
        print(f"  {fname}")
        ft = s.get("file_type") or "unrecognized"
        print(f"    File type:         {ft}")
        print(f"    Rows processed:    {s['total']}")
        if s["filtered_out"]:
            print(f"    Competitors removed: {len(s['filtered_out'])} ({', '.join(str(x) for x in s['filtered_out'])})")
        print(f"    Industries mapped: {s['industry_mapped']}")
        print(f"    States filled:     {s['state_filled']}")

    dupes = total_rows - len(contact_rows)
    print(f"\n  Combined total rows:    {total_rows}")
    print(f"  Competitor rows removed:{total_filtered}")
    print(f"\n📋 Contacts staging  ({len(contact_rows)} unique people, {dupes} dupes removed):")
    print(f"   {contacts_path}")
    print(f"\n📣 Engagements staging ({total_rows} rows):")
    print(f"   {engagements_path}")
    print(f"\n🔍 ZoomInfo enrichment list ({len(contact_rows)} contacts to enrich):")
    print(f"   {to_enrich_path}")
    print(f"\n→ Next: enrich contacts via ZoomInfo, then run finalize.py")

    return contacts_path, engagements_path, to_enrich_path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(
        description="Normalize Limelight export file(s) into Contacts and Engagements outputs."
    )
    parser.add_argument("files", nargs="+", help="One or more Limelight .xlsx input files")
    parser.add_argument(
        "--output-dir", "-o",
        default=None,
        help="Directory for output files (default: same folder as the first input file)"
    )
    args = parser.parse_args()

    missing = [f for f in args.files if not os.path.exists(f)]
    if missing:
        print(f"Error: file(s) not found: {', '.join(missing)}")
        sys.exit(1)

    out_dir = args.output_dir or str(Path(args.files[0]).parent)
    run(args.files, out_dir)
