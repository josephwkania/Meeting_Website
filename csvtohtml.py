'''
modify the parameter csv_file_path with actual file path.
the registered participants is saved as registered_participants.html
'''

import csv
import os
from datetime import datetime

def process_attendees_csv(csv_file_path):
    """
    Process the attendees CSV file and generate an HTML file with in-person and remote attendees.
    Handles duplicates by using the last entry for each person.
    """
    # Dictionary to store attendees data (will handle duplicates by overwriting)
    attendees_dict = {}
    
    # Read the CSV file
    with open(csv_file_path, 'r', encoding='latin-1') as csvfile:
        reader = csv.DictReader(csvfile)
        
        # Check the first row to determine column names
        # This helps debug the actual column names in the CSV
        first_row = None
        csvfile.seek(0)  # Reset file pointer to beginning
        for row in reader:
            first_row = row
            break
        
        if first_row:
            print("Available column names in CSV:")
            for key in first_row.keys():
                print(f"  - '{key}'")
        
        # Find the institution column - try different possible names
        institution_column = None
        possible_institution_columns = [
            'Institute/Affiliation', 
            'Institution/Affiliation', 
            'Institution', 
            'Affiliation',
            'Institute/Affiliation:',
            'Institute/Affiliation:'
        ]
        
        for col in possible_institution_columns:
            if col in first_row:
                institution_column = col
                print(f"Using '{institution_column}' as the institution column")
                break
        
        if not institution_column:
            print("Warning: Couldn't find institution column. Using empty values.")
        
        # Find the attendance mode column
        attendance_column = None
        possible_attendance_columns = [
            'Attendance Mode', 
            'Attendance', 
            'Mode',
            'Attendance Mode:'
        ]
        
        for col in possible_attendance_columns:
            if col in first_row:
                attendance_column = col
                print(f"Using '{attendance_column}' as the attendance mode column")
                break
        
        if not attendance_column:
            print("Warning: Couldn't find attendance mode column. Assuming all attendees are in-person.")
        
        # Reset file pointer and re-read
        csvfile.seek(0)
        next(reader)  # Skip header row
        
        # Process the rows
        for row in reader:
            # Skip rows with empty first or last name
            if not row.get('First Name:') or not row.get('Last Name:'):
                continue
                
            # Create a unique key using first and last name
            name_key = f"{row['First Name:'].strip()} {row['Last Name:'].strip()}"
            
            # Store the entire row data (overwrites if duplicate)
            attendees_dict[name_key] = row
    
    # Separate attendees into in-person and remote lists
    in_person_attendees = []
    remote_attendees = []
    
    for name, data in attendees_dict.items():
        # Get the attendance mode
        attendance_mode = ''
        if attendance_column:
            attendance_mode = data.get(attendance_column, '').strip().lower()
        
        # Get institution
        institution = ''
        if institution_column:
            institution = data.get(institution_column, '').strip()
        
        # Create a tuple with name and affiliation
        attendee_info = (name, institution)
        
        # Add to appropriate list based on attendance mode
        if attendance_mode == 'remote' or attendance_mode == 'virtual' or attendance_mode == 'online':
            remote_attendees.append(attendee_info)
        else:  # Default to in-person for any other values
            in_person_attendees.append(attendee_info)
    
    # Sort both lists alphabetically
    in_person_attendees.sort(key=lambda x: x[0].lower())
    remote_attendees.sort(key=lambda x: x[0].lower())
    
    # Generate HTML file
    generate_html(in_person_attendees, remote_attendees)
    
    # Return counts for verification
    return {
        'total': len(attendees_dict),
        'in_person': len(in_person_attendees),
        'remote': len(remote_attendees)
    }

def generate_html(in_person_attendees, remote_attendees):
    """
    Generate HTML file with two columns: in-person and remote attendees
    """
    # Get current date and time for the "Last Updated" field
    current_datetime = datetime.now().strftime('%H:%M:%S, %d %B %Y (UTC)')
    
    # Start building the HTML content
    html_content = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Registered Participants</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
        }}
        .wrapper {{
            max-width: 1200px;
            margin: 0 auto;
        }}
        h2 {{
            margin-bottom: 20px;
        }}
        .attendees-container {{
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
        }}
        .attendees-column {{
            flex: 1;
            min-width: 300px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            padding: 8px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #f2f2f2;
        }}
        .last-updated {{
            margin-bottom: 20px;
        }}
    </style>
</head>
<body>
    <div class="wrapper style1">
        <section id="attendees" class="my-section">
            <h2>Registered Participants</h2>
            <div class="last-updated">
                <p><b>Last Updated:</b> {current_datetime}</p>
            </div>
            
            <div class="attendees-container">
                <!-- In-Person Attendees Column -->
                <div class="attendees-column">
                    <h3>In-Person Attendees</h3>
                    <table class="table">
                        <tr>
                            <th>Name</th>
                            <th>Institution</th>
                        </tr>
    '''
    
    # Add in-person attendees
    for name, institution in in_person_attendees:
        html_content += f'''
                        <tr>
                            <td>{name}</td>
                            <td>{institution}</td>
                        </tr>
        '''
    
    # Add remote attendees section
    html_content += f'''
                    </table>
                </div>
                
                <!-- Remote Attendees Column -->
                <div class="attendees-column">
                    <h3>Remote Attendees</h3>
                    <table class="table">
                        <tr>
                            <th>Name</th>
                            <th>Institution</th>
                        </tr>
    '''
    
    # Add remote attendees
    for name, institution in remote_attendees:
        html_content += f'''
                        <tr>
                            <td>{name}</td>
                            <td>{institution}</td>
                        </tr>
        '''
    
    # Close the HTML
    html_content += '''
                    </table>
                </div>
            </div>
        </section>
    </div>
</body>
</html>
    '''
    
    # Write the HTML to a file
    with open('registered_participants.html', 'w', encoding='utf-8') as html_file:
        html_file.write(html_content)
    
    print(f"HTML file 'registered_participants.html' generated successfully.")

if __name__ == "__main__":
    # Path to the CSV file
    csv_file_path = "/Users/user/Downloads/Participants List(Sheet1).csv"
    
    if os.path.exists(csv_file_path):
        counts = process_attendees_csv(csv_file_path)
        print(f"Total unique attendees: {counts['total']}")
        print(f"In-person attendees: {counts['in_person']}")
        print(f"Remote attendees: {counts['remote']}")
    else:
        print(f"Error: File '{csv_file_path}' does not exist.")
