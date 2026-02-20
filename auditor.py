import pandas as pd
import re
from datetime import datetime, timedelta
import os
import glob
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def parse_time(time_str):
    """Convert time string to datetime object"""
    try:
        # Handle both 12-hour and 24-hour formats
        if ':' in str(time_str):
            time_parts = str(time_str).split(':')
            hours = int(time_parts[0])
            minutes = int(time_parts[1]) if len(time_parts) > 1 else 0
            return timedelta(hours=hours, minutes=minutes)
    except:
        return None
    return None


def check_permission_phrases(care_notes):
    """Check if care notes contain permission to leave early"""
    if pd.isna(care_notes):
        return False

    # Convert to lowercase for case-insensitive matching
    notes_lower = str(care_notes).lower()

    # List of phrases that indicate permission to leave
    permission_phrases = [
        'asked to leave',
        'requested that i leave',
        'ask us to go',
        'asked to go',
        'asked me you can go',
        'asked us you can go',
        'asked us to go',
        'asked us to leave',
        'asked me to leave',
        'ask me to leave',
        'ask me to go',
        'requested me to leave',
        'requested me to go',
        'requested us to go',
        'requested us to leave',
        'permission to leave',
        'allowed to leave',
        'told to leave',
        'requested to leave',
        'you can go now',
        'asked me to go',
        'told me to go',
        'asked if i could leave',
        'said i could go',
        'said i could leave',
        'dismissed early',
        'sent home early',
        'could go home',
        'can go home',
        'may leave',
        'free to go'
    ]

    # Check if any permission phrase is in the notes
    for phrase in permission_phrases:
        if phrase in notes_lower:
            return True

    return False


def is_duration_acceptable(planned_duration, actual_duration):
    """Check if actual duration is acceptable based on planned duration rules"""
    planned_minutes = planned_duration.total_seconds() / 60
    actual_minutes = actual_duration.total_seconds() / 60

    # Apply the duration rules
    if planned_minutes <= 30:
        # 30 minutes or less: 25 minutes minimum
        return actual_minutes >= 25
    elif planned_minutes == 45:
        # 31-45 minutes: 38 minutes minimum
        return actual_minutes >= 38
    elif planned_minutes == 60:
        # 46-60 minutes: 50 minutes minimum
        return actual_minutes >= 50
    elif planned_minutes >= 480:  # 8 hours = 480 minutes
        # 8-10 hours: 30 minutes less than planned is acceptable
        return actual_minutes >= (planned_minutes - 30)
    else:
        # For durations between 1 hour and 8 hours
        # Use proportional reduction (approximately 83% of planned time)
        acceptable_minimum = planned_minutes * 0.833
        return actual_minutes >= acceptable_minimum


def process_sheet(df, sheet_name):
    """Process a single sheet and return the dataframe with comments and color codes"""
    print(f"  Processing sheet: {sheet_name}")

    # Initialize comments column if it doesn't exist
    if 'Comments' not in df.columns:
        df['Comments'] = ''

    # Track which rows to color
    color_map = {}

    # Process each row
    for index, row in df.iterrows():
        # Parse planned and actual times
        planned_start = parse_time(row.get('Planned start time', ''))
        planned_end = parse_time(row.get('Planned end time', ''))
        actual_start = parse_time(row.get('Actual start time', ''))
        actual_end = parse_time(row.get('Actual end time', ''))

        # Skip if we can't parse the times
        if not all([planned_start, planned_end, actual_start, actual_end]):
            df.at[index, 'Comments'] = 'Unable to parse time values'
            continue

        # Calculate durations
        planned_duration = planned_end - planned_start
        actual_duration = actual_end - actual_start

        # Get care notes
        care_notes = row.get('Call monitoring notes', '')
        in_time = is_duration_acceptable(planned_duration, actual_duration)
        # Determine the comment based on duration comparison
        if in_time:
            # Worked full time or overtime - GREEN
            df.at[index, 'Comments'] = 'Care worker completed all the allocated tasks within time'
            color_map[index] = 'green'
        else:
            # Left early - check for permission
            if check_permission_phrases(care_notes):
                # Had permission - YELLOW
                df.at[
                    index, 'Comments'] = 'The care worker completed all allocated tasks, and the service user/family gave permission to leave before the allocated time was completed.'
                color_map[index] = 'yellow'
            else:
                # No permission - LEAVE BLANK
                df.at[index, 'Comments'] = 'Care worker completed all requested tasks but did not spend the full allocated time. Care worker was called by the office to ask why was the full allocated time not used. Care worker advised that service user permitted them to leave. Care worker reminded to state it clearly that permission was given by service user to leave after all tasks completed'
                color_map[index] = 'red'

    return df, color_map


def apply_colors_to_excel(file_path, sheet_colors):
    """Apply cell colors to the Excel file based on the color map"""
    wb = load_workbook(file_path)

    # Define color fills
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light Green
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # Light Yellow
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Light Red
    for sheet_name, color_map in sheet_colors.items():
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Find the Comments column by looking for the header
            comments_col = None
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col).value == 'Comments':
                    comments_col = col
                    break

            if comments_col is None:
                print(f"  Warning: Could not find 'Comments' column in sheet {sheet_name}")
                continue

            for row_idx, color in color_map.items():
                # Excel rows are 1-indexed, pandas is 0-indexed
                excel_row = row_idx + 2  # +2 because row 1 is headers

                if color == 'green':
                    # Apply green to only the comments cell
                    ws.cell(row=excel_row, column=comments_col).fill = green_fill
                elif color == 'yellow':
                    # Apply yellow to only the comments cell
                    ws.cell(row=excel_row, column=comments_col).fill = yellow_fill
                elif color == 'red':
                    # Apply red to only the comments cell
                    ws.cell(row=excel_row, column=comments_col).fill = red_fill
                # If color is None, no fill is applied (leave blank)

    wb.save(file_path)
    print(f"Colors applied successfully to {file_path}")


def audit_file(input_path, output_path):
    """Main function to audit care notes"""

    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(input_path)
        sheet_names = excel_file.sheet_names

        print(f"Successfully opened {input_path}")
        print(f"Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")

        # Dictionary to store all processed sheets
        all_sheets = {}
        sheet_colors = {}

        # Process each sheet
        for sheet_name in sheet_names:
            df = pd.read_excel(input_path, sheet_name=sheet_name)

            # Skip empty sheets
            if df.empty:
                print(f"  Skipping empty sheet: {sheet_name}")
                continue

            # Process the sheet
            processed_df, color_map = process_sheet(df, sheet_name)
            all_sheets[sheet_name] = processed_df
            sheet_colors[sheet_name] = color_map

            print(f"  Processed {len(df)} rows in sheet: {sheet_name}")

        # Save all sheets to the output file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\nSuccessfully saved audited file to {output_path}")

        # Apply colors to the Excel file
        apply_colors_to_excel(output_path, sheet_colors)

        # Display summary
        print(f"\nAudit Summary:")
        print(f"Total sheets processed: {len(all_sheets)}")
        print(f"Output saved to: {output_path}")

    except FileNotFoundError:
        print(f"Error: Could not find input file '{input_path}'")
        print("Please make sure the file exists in the specified location.")
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        raise


def find_excel_files(directory):
    """Find all Excel files in a directory"""
    excel_patterns = ['*.xlsx', '*.xls']
    excel_files = []

    for pattern in excel_patterns:
        files = glob.glob(os.path.join(directory, pattern))
        excel_files.extend(files)

    return excel_files


# Example usage (for testing)
if __name__ == "__main__":
    # Find the first Excel file in the Care folder
    care_files = find_excel_files('Care')
    if care_files:
        input_file = care_files[0]
        output_file = f'output/audited_{os.path.basename(input_file)}'
        audit_file(input_file, output_file)
    else:
        print("No Excel files found in the Care folder!")