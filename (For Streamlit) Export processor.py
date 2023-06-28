import streamlit as st
import csv
import pandas as pd
from collections import defaultdict
import openpyxl

def convert_region_value(value):
    scotland_regions = ['East of Scotland', 'West of Scotland', 'Tayside', 'Aberdeen', 'Highlands and Islands', 'South of Scotland']
    england_regions = ['North West', 'South East', 'East Midlands', 'South West', 'North East', 'West Midlands']

    if value in scotland_regions:
        return 'Scotland'
    elif value in england_regions:
        return value + ' (England)'
    else:
        return value

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

def count_accelerators(df):
    # Get list of columns that contain "Academic Institution Name"
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

    return accelerator_counts_df

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

country_list = ['Åland Islands', 'Albania', 'Afghanistan', 'Algeria', 'Azerbaijan', 'American Samoa', 'Andorra', 'Angola', 'Anguilla', 'Antarctica', 'Antigua and Barbuda', 'Argentina', 'Armenia', 'Aruba', 'Australia', 'Austria', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium', 'Belize', 'Benin', 'Bermuda', 'Bhutan', 'Bolivia, Plurinational State Of', 'Bosnia and Herzegovina', 'Botswana', 'Bouvet Island', 'Brazil', 'British Indian Ocean Territory', 'Brunei Darussalam', 'Bulgaria', 'Burkina Faso', 'Burundi', 'Cambodia', 'Cameroon', 'Canada', 'Cape Verde', 'Cayman Islands', 'Central African Republic', 'Chad', 'Chile', 'China', 'Christmas Island', 'Cocos (Keeling) Islands', 'Colombia', 'Comoros', 'Congo', 'Congo, The Democratic Republic Of The', 'Cook Islands', 'Costa Rica', "Côte D'Ivoire", 'Croatia', 'Cuba', 'Curaçao', 'Cyprus', 'Czech Republic', 'Denmark', 'Djibouti', 'Dominica', 'Dominican Republic', 'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Falkland Islands (Malvinas)', 'Faroe Islands', 'Fiji', 'Finland', 'French Guiana', 'French Polynesia', 'French Southern Territories', 'Gabon', 'Gambia', 'Georgia', 'Ghana', 'Gibraltar', 'Greenland', 'Grenada', 'Guadeloupe', 'Guam', 'Guatemala', 'Guinea', 'Guinea-Bissau', 'Guyana', 'Haiti', 'Heard Island and Mcdonald Islands', 'Holy See (Vatican City State)', 'Honduras', 'Hong Kong', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran, Islamic Republic Of', 'Iraq', 'Ireland', 'Isle Of Man', 'Jamaica', 'Japan', 'Jersey', 'Bonaire, Saint Eustatius and Saba', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', "Korea, Democratic People's Republic Of (North Korea)", 'Korea, Republic Of (South Korea)', 'Kuwait', 'Kyrgyzstan', 'Lao People\'S Democratic Republic', 'Latvia', 'Lebanon', 'Lesotho', 'Liberia', 'Liechtenstein', 'Lithuania', 'Luxembourg', 'Macao', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali', 'Malta', 'Marshall Islands', 'Martinique', 'Mauritania', 'Mauritius', 'Mayotte', 'Mexico', 'Micronesia, Federated States Of', 'Moldova, Republic Of', 'Monaco', 'Mongolia', 'Montenegro', 'Montserrat', 'Morocco', 'Mozambique', 'Myanmar', 'Namibia', 'Nauru', 'Nepal', 'New Caledonia', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Niue', 'Norfolk Island', 'Northern Mariana Islands', 'Norway', 'Oman', 'Pakistan', 'Palau', 'Palestinian Territory, Occupied', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Pitcairn', 'Poland', 'Portugal', 'Puerto Rico', 'Qatar', 'Réunion', 'Romania', 'Rwanda', 'Saint Barthélemy', 'Saint Helena, Ascension and Tristan Da Cunha', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Martin (French Part)', 'Saint Pierre and Miquelon', 'Saint Vincent and The Grenadines', 'Samoa', 'San Marino', 'Sao Tome and Principe', 'Saudi Arabia', 'Senegal', 'Serbia', 'Seychelles', 'Sierra Leone', 'Singapore', 'Sint Maarten (Dutch Part)', 'Slovakia', 'Slovenia', 'Solomon Islands', 'Somalia', 'South Georgia and The South Sandwich Islands', 'South Sudan', 'North Macedonia, Republic of', 'Libya, State of', 'Sri Lanka', 'Sudan', 'Suriname', 'Svalbard and Jan Mayen', 'Sweden', 'Switzerland', 'Syrian Arab Republic', 'Taiwan', 'Tajikistan', 'Tanzania, United Republic Of', 'Thailand', 'Timor-Leste', 'Togo', 'Tokelau', 'Tonga', 'Trinidad and Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Turks and Caicos Islands', 'Tuvalu', 'Uganda', 'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States Minor Outlying Islands', 'Uruguay', 'Uzbekistan', 'Vanuatu', 'Viet Nam', 'Virgin Islands, British', 'Virgin Islands, U.S.', 'Wallis and Futuna', 'Western Sahara', 'Yemen', 'Zambia', 'Zimbabwe', 'Eswatini', 'Venezuela, Bolivarian Republic Of', 'France', 'Germany', 'Greece', 'Guernsey', 'Israel', 'Italy', 'Netherlands', 'Russian Federation', 'South Africa', 'Spain', 'United States']

# Assuming 'input.csv' is the CSV file you want to convert
df = pd.read_csv('input.csv')  # Replace 'input.csv' with the actual file path or provide your DataFrame

def main():
    st.title("Raw export processor")
    st.write("Upload a CSV file and get the processed Excel file as output.")

    # File upload
    uploaded_file = st.file_uploader("Upload a CSV file", type="csv")

    if uploaded_file is not None:
        # Read the uploaded file into a DataFrame
        df = pd.read_csv(uploaded_file)


# Apply convert_region_value function to the 'Data' tab of the DataFrame
df['Head Office Address - Region'] = df['Head Office Address - Region'].apply(convert_region_value)

university_counts_df = count_universities(df)
accelerator_counts_df = count_accelerators(df)
age_brackets_df = count_age_brackets(df)
nationality_counts_df = count_director_nationalities(df, country_list)

# Create a Pandas Excel writer
writer = pd.ExcelWriter('processed_export.xlsx')

# Write the original DataFrame to the first sheet
df.to_excel(writer, sheet_name='Data', index=False)

# Write the institution counts DataFrame to a separate sheet
university_counts_df.to_excel(writer, sheet_name='Institution Counts', index=False)

# Write the accelerator counts DataFrame to a separate sheet
accelerator_counts_df.to_excel(writer, sheet_name='Accelerator Counts', index=False)

# Write the director age bracket counts DataFrame to a separate sheet
age_brackets_df.to_excel(writer, sheet_name='Director ages', index=False)

# Write the director nationality counts DataFrame to a separate sheet
nationality_counts_df.to_excel(writer, sheet_name='Director nationalities', index=False)

# Save the Excel file
writer.save()

 # Provide the download link to the processed Excel file
st.success("Excel file processed successfully!")
st.download_button("Download processed Excel file", "processed_export.xlsx")

if __name__ == '__main__':
    main()