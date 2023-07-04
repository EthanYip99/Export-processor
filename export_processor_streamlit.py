import streamlit as st
import csv
from collections import defaultdict
import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Install openpyxl
install("git+https://github.com/openpyxl/openpyxl.git")

import pandas as pd

def convert_region_value(value):
    scotland_regions = ['East of Scotland', 'West of Scotland', 'Tayside', 'Aberdeen', 'Highlands and Islands', 'South of Scotland']
    england_regions = ['North West', 'South East', 'East Midlands', 'South West', 'North East', 'West Midlands']

    if value in scotland_regions:
        return 'Scotland'
    elif value in england_regions:
        return value + ' (England)'
    else:
        return value
    pass


def count_universities(df):
    # Get list of columns that contain "Academic Institution Name"
    institution_cols = [col for col in df.columns if "Academic Institution Name" in col]

    # Concatenate all these columns into a single Series
    institution_series = pd.concat([df[col] for col in institution_cols])

    # Drop all rows that are NaN (since they don't represent a university)
    institution_series = institution_series.dropna()

    # Count the frequency of each university
    institution_counts = institution_series.value_counts()

    # Convert the Series to a DataFrame
    institution_counts_df = institution_counts.reset_index()
    institution_counts_df.columns = ['University', 'Count']
    return institution_counts_df
    pass

def count_accelerators(df):
    # Get list of columns that contain "Accelerator Name"
    accelerator_cols = [col for col in df.columns if "Accelerator Name" in col]

    # Concatenate all these columns into a single Series
    accelerator_series = pd.concat([df[col] for col in accelerator_cols])

    # Drop all rows that are NaN (since they don't represent an accelerator)
    accelerator_series = accelerator_series.dropna()

    # Count the frequency of each university
    accelerator_counts = accelerator_series.value_counts()

    # Convert the Series to a DataFrame
    accelerator_counts_df = accelerator_counts.reset_index()
    accelerator_counts_df.columns = ['Accelerator', 'Count']
    return  accelerator_counts_df
    pass

def count_age_brackets(df):
    age_columns = [col for col in df.columns if "Age (if known)" in col]

    age_ranges = ['20-29', '30-39', '40-49', '50-59', '60-69', '70+']
    age_counts_dict = defaultdict(int)

    for col in age_columns:
        age_series = df[col].dropna()

        # Convert age strings to numeric values
        age_series = pd.to_numeric(age_series, errors='coerce')

        # Exclude age values that are NaN or negative
        age_series = age_series[(age_series >= 0)]

        # Calculate counts for each age range
        age_counts_df = pd.cut(age_series, bins=[19, 29, 39, 49, 59, 69, float('inf')], labels=age_ranges).value_counts().to_frame()
        age_counts_df.columns = ['Count']

        for age_range, count in age_counts_df['Count'].items():
            age_counts_dict[age_range] += count

    age_brackets_df = pd.DataFrame(list(age_counts_dict.items()), columns=['Age Bracket', 'Count'])
    return age_brackets_df
    pass

def count_director_nationalities(df, country_list):
    director_cols = [col for col in df.columns if "Current Directors" in col and "Nationalities" in col]

    nationality_counts = {country: 0 for country in country_list}

    for col in director_cols:
        nationality_series = df[col].dropna()

        for nationality in nationality_series:
            for country in country_list:
                if country in nationality:
                    nationality_counts[country] += 1

    nationality_counts_df = pd.DataFrame({"Nationality": list(nationality_counts.keys()), "Count": list(nationality_counts.values())})
    return nationality_counts_df
    pass

country_list = ['Åland Islands', 'Albania', 'Afghanistan', 'Algeria', 'Azerbaijan', 'American Samoa', 'Andorra', 'Angola', 'Anguilla', 'Antarctica', 'Antigua and Barbuda', 'Argentina', 'Armenia', 'Aruba', 'Australia', 'Austria', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium', 'Belize', 'Benin', 'Bermuda', 'Bhutan', 'Bolivia, Plurinational State Of', 'Bosnia and Herzegovina', 'Botswana', 'Bouvet Island', 'Brazil', 'British Indian Ocean Territory', 'Brunei Darussalam', 'Bulgaria', 'Burkina Faso', 'Burundi', 'Cambodia', 'Cameroon', 'Canada', 'Cape Verde', 'Cayman Islands', 'Central African Republic', 'Chad', 'Chile', 'China', 'Christmas Island', 'Cocos (Keeling) Islands', 'Colombia', 'Comoros', 'Congo', 'Congo, The Democratic Republic Of The', 'Cook Islands', 'Costa Rica', "Côte D'Ivoire", 'Croatia', 'Cuba', 'Curaçao', 'Cyprus', 'Czech Republic', 'Denmark', 'Djibouti', 'Dominica', 'Dominican Republic', 'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Falkland Islands (Malvinas)', 'Faroe Islands', 'Fiji', 'Finland', 'French Guiana', 'French Polynesia', 'French Southern Territories', 'Gabon', 'Gambia', 'Georgia', 'Ghana', 'Gibraltar', 'Greenland', 'Grenada', 'Guadeloupe', 'Guam', 'Guatemala', 'Guinea', 'Guinea-Bissau', 'Guyana', 'Haiti', 'Heard Island and Mcdonald Islands', 'Holy See (Vatican City State)', 'Honduras', 'Hong Kong', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran, Islamic Republic Of', 'Iraq', 'Ireland', 'Isle Of Man', 'Jamaica', 'Japan', 'Jersey', 'Bonaire, Saint Eustatius and Saba', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', "Korea, Democratic People's Republic Of (North Korea)", 'Korea, Republic Of (South Korea)', 'Kuwait', 'Kyrgyzstan', 'Lao People\'S Democratic Republic', 'Latvia', 'Lebanon', 'Lesotho', 'Liberia', 'Liechtenstein', 'Lithuania', 'Luxembourg', 'Macao', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali', 'Malta', 'Marshall Islands', 'Martinique', 'Mauritania', 'Mauritius', 'Mayotte', 'Mexico', 'Micronesia, Federated States Of', 'Moldova, Republic Of', 'Monaco', 'Mongolia', 'Montenegro', 'Montserrat', 'Morocco', 'Mozambique', 'Myanmar', 'Namibia', 'Nauru', 'Nepal', 'New Caledonia', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Niue', 'Norfolk Island', 'Northern Mariana Islands', 'Norway', 'Oman', 'Pakistan', 'Palau', 'Palestinian Territory, Occupied', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Pitcairn', 'Poland', 'Portugal', 'Puerto Rico', 'Qatar', 'Réunion', 'Romania', 'Rwanda', 'Saint Barthélemy', 'Saint Helena, Ascension and Tristan Da Cunha', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Martin (French Part)', 'Saint Pierre and Miquelon', 'Saint Vincent and The Grenadines', 'Samoa', 'San Marino', 'Sao Tome and Principe', 'Saudi Arabia', 'Senegal', 'Serbia', 'Seychelles', 'Sierra Leone', 'Singapore', 'Sint Maarten (Dutch Part)', 'Slovakia', 'Slovenia', 'Solomon Islands', 'Somalia', 'South Georgia and The South Sandwich Islands', 'South Sudan', 'North Macedonia, Republic of', 'Libya, State of', 'Sri Lanka', 'Sudan', 'Suriname', 'Svalbard and Jan Mayen', 'Sweden', 'Switzerland', 'Syrian Arab Republic', 'Taiwan', 'Tajikistan', 'Tanzania, United Republic Of', 'Thailand', 'Timor-Leste', 'Togo', 'Tokelau', 'Tonga', 'Trinidad and Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Turks and Caicos Islands', 'Tuvalu', 'Uganda', 'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States Minor Outlying Islands', 'Uruguay', 'Uzbekistan', 'Vanuatu', 'Viet Nam', 'Virgin Islands, British', 'Virgin Islands, U.S.', 'Wallis and Futuna', 'Western Sahara', 'Yemen', 'Zambia', 'Zimbabwe', 'Eswatini', 'Venezuela, Bolivarian Republic Of', 'France', 'Germany', 'Greece', 'Guernsey', 'Israel', 'Italy', 'Netherlands', 'Russian Federation', 'South Africa', 'Spain', 'United States']


def count_fund_managers(df):
    # Get list of columns that contain "Fund Manager"
    fund_manager_cols = [col for col in df.columns if "Fundraising investors" in col and "Fund manager" in col]

    # Concatenate all these columns into a single Series
    fund_manager_series = pd.concat([df[col] for col in fund_manager_cols])

    # Drop all rows that are NaN (since they don't represent a university)
    fund_manager_series = fund_manager_series.dropna()

    # Count the frequency of each university
    fund_manager_counts = fund_manager_series.value_counts()

    # Convert the Series to a DataFrame
    fund_manager_counts_df = fund_manager_counts.reset_index()
    fund_manager_counts_df.columns = ['Fund manager', 'Count of deal participations']
    return fund_manager_counts_df
    pass


def count_fund_country(df):
   # Get list of columns that contain "Head office country"
    fund_country_cols = [col for col in df.columns if "Fundraising investors" in col and "Head office country" in col]
    
    # Concatenate all these columns into a single Series
    fund_country_series = pd.concat([df[col] for col in fund_country_cols])

    # Drop all rows that are NaN (since they don't represent a university)
    fund_country_series = fund_country_series.dropna()

    # Count the frequency of each university
    fund_country_counts = fund_country_series.value_counts()

    # Convert the Series to a DataFrame
    fund_country_counts_df = fund_country_counts.reset_index()
    fund_country_counts_df.columns = ['Fund - Head office country', 'Count of deal participations']
    return fund_country_counts_df
    pass


def count_fund_type(df):
  # Get list of columns that contain "fund type"
    fund_type_cols = [col for col in df.columns if "Fundraising investors" in col and "Fund type" in col]
    
    # Concatenate all these columns into a single Series
    fund_type_series = pd.concat([df[col] for col in fund_type_cols])

    # Drop all rows that are NaN (since they don't represent a university)
    fund_type_series = fund_type_series.dropna()

    # Count the frequency of each university
    fund_type_counts = fund_type_series.value_counts()

    # Convert the Series to a DataFrame
    fund_type_counts_df = fund_type_counts.reset_index()
    fund_type_counts_df.columns = ['Fund type', 'Count of deal participations']
    return fund_type_counts_df
    pass


def process_company_export(df):
    university_counts_df = count_universities(df)
    accelerator_counts_df = count_accelerators(df)
    age_brackets_df = count_age_brackets(df)
    nationality_counts_df = count_director_nationalities(df, country_list)

    # Create a Pandas Excel writer
    writer = pd.ExcelWriter('processed_company_export.xlsx')
   
    # Write the original DataFrame to the first sheet
    df.to_excel(writer, sheet_name='Raw data', index=False)

    # Write the institution counts DataFrame to a separate sheet
    university_counts_df.to_excel(writer, sheet_name='Institution Counts', index=False)

    # Write the accelerator counts DataFrame to a separate sheet
    accelerator_counts_df.to_excel(writer, sheet_name='Accelerator Counts', index=False)

    # Write the director age bracket counts DataFrame to a separate sheet
    age_brackets_df.to_excel(writer, sheet_name='Director ages', index=False)

    # Write the director nationality counts DataFrame to a separate sheet
    nationality_counts_df.to_excel(writer, sheet_name='Director nationalities', index=False)
    
    writer.save()
    writer.close()

def process_fundraising_export(df):
    fund_manager_counts_df = count_fund_managers(df)
    fund_country_counts_df = count_fund_country(df)
    fund_type_counts_df = count_fund_type(df)

    # Create a Pandas Excel writer
    writer = pd.ExcelWriter('processed_fundraising_export.xlsx')

    # Write the original DataFrame to the first sheet
    df.to_excel(writer, sheet_name='Raw data', index=False)

    # Write the fund manager counts DataFrame to a separate sheet
    fund_manager_counts_df.to_excel(writer, sheet_name='Fund manager counts', index=False)

    # Write the fund head office country counts DataFrame to a separate sheet
    fund_country_counts_df.to_excel(writer, sheet_name='Deal counts by country', index=False)

    # Write the fund type counts DataFrame to a separate sheet
    fund_type_counts_df.to_excel(writer, sheet_name='Deal counts by fund type', index=False)
    writer.save()
    writer.close()


def main():
    st.title("Data Export Processing")
    export_type = st.selectbox("Select Export Type", ["Company Data", "Fundraising Data"])

    file = st.file_uploader("Upload a CSV file", type=["csv"])

    if file is not None:
        try:
            df = pd.read_csv(file)
            if export_type == "Company Data":
                process_company_export(df)
                st.success("Company export processed successfully. Click below to download the Excel file.")
                st.download_button("Download Processed File", "processed_company_export.xlsx")
            elif export_type == "Fundraising Data":
                process_fundraising_export(df)
                st.success("Fundraising export processed successfully. Click below to download the Excel file.")
                st.download_button("Download Processed File", "processed_fundraising_export.xlsx")
        except Exception as e:
            st.error(f"Error occurred during processing: {str(e)}")


if __name__ == '__main__':
    main()
