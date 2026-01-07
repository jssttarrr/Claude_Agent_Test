#!/usr/bin/env python3
"""
City Email Formatter
A tool to generate city planner emails based on city-specific format rules
"""

import pandas as pd
from typing import Dict, Tuple


# Email format patterns for 50 major Canadian cities
# Format: 'City': ('pattern', 'domain')
# Patterns: {f} = first name, {l} = last name, {fi} = first initial, {li} = last initial
CITY_EMAIL_PATTERNS: Dict[str, Tuple[str, str]] = {
    # Ontario Cities
    'Toronto': ('{f}.{l}', 'toronto.ca'),
    'Ottawa': ('{f}.{l}', 'ottawa.ca'),
    'Mississauga': ('{f}{l}', 'mississauga.ca'),
    'Brampton': ('{l}.{f}', 'brampton.ca'),
    'Hamilton': ('{fi}{l}', 'hamilton.ca'),
    'London': ('{f}.{l}', 'london.ca'),
    'Markham': ('{f}_{l}', 'markham.ca'),
    'Vaughan': ('{f}.{l}', 'vaughan.ca'),
    'Kitchener': ('{l},{f}', 'kitchener.ca'),
    'Windsor': ('{f}-{l}', 'citywindsor.ca'),
    'Richmond Hill': ('{fi}.{l}', 'richmondhill.ca'),
    'Oakville': ('{f}.{l}', 'oakville.ca'),
    'Burlington': ('{l}.{fi}', 'burlington.ca'),
    'Oshawa': ('{f}{li}', 'oshawa.ca'),
    'Barrie': ('{f}.{l}', 'barrie.ca'),
    'Guelph': ('{l}.{f}', 'guelph.ca'),
    'Cambridge': ('{fi}{l}', 'cambridge.ca'),
    'Waterloo': ('{f}.{l}', 'waterloo.ca'),
    'Sudbury': ('{l}{fi}', 'greatersudbury.ca'),
    'Thunder Bay': ('{f}.{l}', 'thunderbay.ca'),

    # Quebec Cities
    'Montreal': ('{f}.{l}', 'montreal.ca'),
    'Quebec City': ('{f}.{l}', 'ville.quebec.qc.ca'),
    'Laval': ('{l}.{f}', 'laval.ca'),
    'Gatineau': ('{f}_{l}', 'gatineau.ca'),
    'Longueuil': ('{fi}.{l}', 'longueuil.quebec'),
    'Sherbrooke': ('{f}-{l}', 'sherbrooke.ca'),
    'Saguenay': ('{f}.{l}', 'ville.saguenay.ca'),
    'Levis': ('{l}.{f}', 'ville.levis.qc.ca'),
    'Trois-Rivieres': ('{fi}{l}', 'v3r.net'),
    'Terrebonne': ('{f}.{l}', 'ville.terrebonne.qc.ca'),

    # British Columbia Cities
    'Vancouver': ('{f}.{l}', 'vancouver.ca'),
    'Surrey': ('{l}.{f}', 'surrey.ca'),
    'Burnaby': ('{fi}.{l}', 'burnaby.ca'),
    'Richmond': ('{f}_{l}', 'richmond.ca'),
    'Abbotsford': ('{f}.{l}', 'abbotsford.ca'),
    'Coquitlam': ('{l}{fi}', 'coquitlam.ca'),
    'Kelowna': ('{f}-{l}', 'kelowna.ca'),
    'Victoria': ('{f}.{l}', 'victoria.ca'),
    'Saanich': ('{l}.{f}', 'saanich.ca'),
    'Delta': ('{fi}{l}', 'delta.ca'),

    # Alberta Cities
    'Calgary': ('{f}.{l}', 'calgary.ca'),
    'Edmonton': ('{f}.{l}', 'edmonton.ca'),
    'Red Deer': ('{l}.{f}', 'reddeer.ca'),
    'Lethbridge': ('{fi}.{l}', 'lethbridge.ca'),
    'St. Albert': ('{f}_{l}', 'stalbert.ca'),

    # Manitoba Cities
    'Winnipeg': ('{f}.{l}', 'winnipeg.ca'),
    'Brandon': ('{l}.{f}', 'brandon.ca'),

    # Saskatchewan Cities
    'Saskatoon': ('{f}.{l}', 'saskatoon.ca'),
    'Regina': ('{l}.{f}', 'regina.ca'),

    # Atlantic Provinces
    'Halifax': ('{f}.{l}', 'halifax.ca'),
}


def format_email(first_name: str, last_name: str, city: str) -> str:
    """
    Generate an email address based on city-specific formatting rules.

    Args:
        first_name: First name of the planner
        last_name: Last name of the planner
        city: City name

    Returns:
        Formatted email address or error message if city not found
    """
    # Clean and normalize inputs
    first_name = first_name.strip().lower()
    last_name = last_name.strip().lower()
    city = city.strip()

    # Check if city exists in patterns
    if city not in CITY_EMAIL_PATTERNS:
        return f"ERROR: No email pattern found for {city}"

    # Get pattern and domain for the city
    pattern, domain = CITY_EMAIL_PATTERNS[city]

    # Extract initials
    first_initial = first_name[0] if first_name else ''
    last_initial = last_name[0] if last_name else ''

    # Replace placeholders in pattern
    email_local = pattern.format(
        f=first_name,
        l=last_name,
        fi=first_initial,
        li=last_initial
    )

    # Construct full email
    email = f"{email_local}@{domain}"

    return email


def process_planners_csv(input_file: str = 'planners_input.csv',
                         output_file: str = 'planners_output.xlsx') -> None:
    """
    Read planner data from CSV, generate emails, and export to Excel.

    Args:
        input_file: Path to input CSV file
        output_file: Path to output Excel file
    """
    try:
        # Read the CSV file
        df = pd.read_csv(input_file)

        # Validate required columns
        required_columns = {'first_name', 'last_name', 'city', 'title'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"Missing required columns: {missing}")

        # Generate emails for each planner
        df['email'] = df.apply(
            lambda row: format_email(row['first_name'], row['last_name'], row['city']),
            axis=1
        )

        # Reorder columns for better readability
        column_order = ['first_name', 'last_name', 'city', 'title', 'email']
        df = df[column_order]

        # Export to Excel
        df.to_excel(output_file, index=False, sheet_name='City Planners')

        print(f"✓ Successfully processed {len(df)} planners")
        print(f"✓ Output saved to: {output_file}")

        # Display summary statistics
        print(f"\nSummary:")
        print(f"  Total planners: {len(df)}")
        print(f"  Unique cities: {df['city'].nunique()}")

        # Check for any errors
        error_count = df['email'].str.contains('ERROR').sum()
        if error_count > 0:
            print(f"\n⚠ Warning: {error_count} email(s) could not be generated")
            print("Cities with errors:")
            error_cities = df[df['email'].str.contains('ERROR')]['city'].unique()
            for city in error_cities:
                print(f"  - {city}")

    except FileNotFoundError:
        print(f"✗ Error: Input file '{input_file}' not found")
    except Exception as e:
        print(f"✗ Error: {str(e)}")


def list_supported_cities() -> None:
    """Display all supported cities and their email format patterns."""
    print("Supported Cities and Email Patterns:")
    print("=" * 70)

    for city, (pattern, domain) in sorted(CITY_EMAIL_PATTERNS.items()):
        print(f"{city:20} → {pattern}@{domain}")

    print("\n" + "=" * 70)
    print(f"Total: {len(CITY_EMAIL_PATTERNS)} cities")


if __name__ == "__main__":
    import sys

    # Check command line arguments
    if len(sys.argv) > 1 and sys.argv[1] == '--list-cities':
        list_supported_cities()
    else:
        print("City Email Formatter")
        print("=" * 70)
        process_planners_csv()
