import spacy 
import pandas as pd


try:
    nlp = spacy.load('en_core_web_sm')
except Exception as e:
    print(f"Error loading spaCy model: {e}")

# Load Excel file into pandas DataFrame
df = pd.read_excel('Study_Fall_2024.xlsx')

# Display the first few rows of the Excel data
print("Excel data:")
print(df.head())

# Function to extract entities
def extract_event_details(text):
    doc = nlp(text)
    time = []
    date = []
    location = []
    description = []
    
    for ent in doc.ents:
        if ent.label_ == 'TIME':
            time.append(ent.text)
        elif ent.label_ == 'DATE':
            date.append(ent.text)
        elif ent.label_ == 'GPE' or ent.label_ == 'LOC':
            location.append(ent.text)
        elif ent.label_ == 'EVENT' or ent.label_ == 'ORG' or ent.label_ == 'PERSON':
            description.append(ent.text)

    return {
        'Time': time,
        'Date': date,
        'Location': location,
        'Description': description
    }

# Process each event description from the Excel sheet
events = []

for index, row in df.iterrows():
    if 'Event Description' in df.columns:
        event_description = row['Event Description']
        event_details = extract_event_details(event_description)
        events.append(event_details)
    else:
        print("No 'Event Description' column found.")

# Convert the results into a DataFrame for easy viewing
events_df = pd.DataFrame(events)
print("\nExtracted Event Details:")
print(events_df)

# Save the extracted details to a new Excel file (optional)
events_df.to_excel('extracted_schedule.xlsx', index=False)
print("\nExtracted schedule saved to 'extracted_schedule.xlsx'.")
