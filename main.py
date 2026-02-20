from auditor import audit_file, find_excel_files
import os


def main():
    # Automatically find Excel files in the Care folder
    care_files = find_excel_files('Care')

    if not care_files:
        print("Error: No Excel files found in the Care folder!")
        print("Please make sure you have .xlsx or .xls files in the 'Care' folder.")
        return

    print("Starting care notes audit...")
    print(f"Found {len(care_files)} Excel file(s) in the Care folder:")

    # Process each Excel file found
    for input_file in care_files:
        # Get just the filename without the path
        filename = os.path.basename(input_file)
        output_file = f'output/audited_{filename}'

        print(f"\nProcessing: {filename}")
        print(f"Input file: {input_file}")
        print(f"Output file: {output_file}")
        print("-" * 50)

        audit_file(input_file, output_file)

        print("-" * 50)

    print("\nAll audits complete!")


if __name__ == "__main__":
    main()