from flask import Flask, request, send_file, jsonify
import pandas as pd
from io import BytesIO
from flask_cors import CORS
from datetime import datetime

app = Flask(__name__)
CORS(app)

@app.route('/update-excel', methods=['POST'])
def update_excel():
    if 'dump' not in request.files or 'master' not in request.files:
        return jsonify({'error': 'Both dump and master files are required.'}), 400

    dump_file = request.files['dump']
    master_file = request.files['master']

    df_dump = pd.read_excel(dump_file)
    df_master = pd.read_excel(master_file)

    # Column mapping
    column_mapping = {
        'Affected User': 'Affected_User',
        'Short description': 'Short_description',
        'Assignment group': 'Assignment_group',
        'Escalated to': 'Escalated_to',
        'Assigned to': 'Assigned_to',
        'SLA due': 'SLA_due',
        'Configuration item': 'Configuration_item',
        'Resolved': 'Resolved_Date',
        'Resolve time': 'Time_Taken',
        'Fault Code': 'Falt_Code',
        'Duration': 'NetWorkDay',
        'Close notes': 'Comment',
        # Don't map Correlation Id here
    }
    df_dump = df_dump.rename(columns=column_mapping)

    # Explicitly create Analysis_Doc column from Correlation ID
    if 'Correlation ID' in df_dump.columns:
        df_dump['Analysis_Doc'] = df_dump['Correlation ID']
    else:
        df_dump['Analysis_Doc'] = ""

    # Ensure all master columns exist in dump
    for col in df_master.columns:
        if col not in df_dump.columns:
            df_dump[col] = ""

    # Convert dates
    df_dump['Created'] = pd.to_datetime(df_dump['Created'], errors='coerce')
    df_dump['Resolved_Date'] = pd.to_datetime(df_dump['Resolved_Date'], errors='coerce')

    now = datetime.now()
    current_month = now.month
    current_year = now.year

    df_master_updated = df_master.copy()

    # Custom logic fields
    df_dump['Falt_Code'] = df_dump.apply(
        lambda x: "Job Failure" if pd.notna(x['Short_description']) and 'WL_' in str(x['Short_description'])
        else (x['Falt_Code'] if pd.notna(x['Falt_Code']) and str(x['Falt_Code']).strip() != "" else ""),
        axis=1
    )

    for _, dump_row in df_dump.iterrows():
        number = str(dump_row['Number'])
        mask = df_master_updated['Number'].astype(str) == number

        if mask.any():
            # Update existing incident
            for col in df_master.columns:
                if col in df_dump.columns:
                    df_master_updated.loc[mask, col] = dump_row[col]

            # Explicitly copy Created and Incident_Assigned_to_us
            df_master_updated.loc[mask, 'Created'] = dump_row['Created']
            df_master_updated.loc[mask, 'Incident_Assigned_to_us'] = dump_row['Created']

            # Copy Correlation ID to Analysis_Doc
            df_master_updated.loc[mask, 'Analysis_Doc'] = dump_row['Analysis_Doc']

            # Close Notes -> Comment
            df_master_updated.loc[mask, 'Comment'] = dump_row['Comment']

            # Parent Incident logic
            if pd.notna(dump_row.get('Parent Incident')) and str(dump_row.get('Parent Incident')).strip():
                df_master_updated.loc[mask, 'State'] = f"{dump_row['State']}/duplicate"
            else:
                df_master_updated.loc[mask, 'State'] = dump_row['State']

            # Closed logic
            if pd.notna(dump_row['Resolved_Date']) and dump_row['Resolved_Date'].month == current_month and dump_row['Resolved_Date'].year == current_year:
                if pd.notna(dump_row.get('Parent Incident')) and str(dump_row.get('Parent Incident')).strip():
                    df_master_updated.loc[mask, 'State'] = "Closed/duplicate"
                else:
                    df_master_updated.loc[mask, 'State'] = "Closed"

        else:
            # Check if Created current month or Correlation ID present
            is_current_month = (
                pd.notna(dump_row['Created']) and
                dump_row['Created'].month == current_month and
                dump_row['Created'].year == current_year
            )
            has_infra = pd.notna(dump_row['Analysis_Doc']) and str(dump_row['Analysis_Doc']).strip() != ""

            if is_current_month or has_infra:
                new_record = {}
                for col in df_master.columns:
                    if col in df_dump.columns:
                        new_record[col] = dump_row[col]
                    else:
                        new_record[col] = ""

                # Explicitly set Created and Incident_Assigned_to_us
                new_record['Created'] = dump_row['Created']
                new_record['Incident_Assigned_to_us'] = dump_row['Created']

                # Correlation Id -> Analysis_Doc
                new_record['Analysis_Doc'] = dump_row['Analysis_Doc']

                # Close Notes -> Comment
                new_record['Comment'] = dump_row['Comment']

                # Parent incident logic
                if pd.notna(dump_row.get('Parent Incident')) and str(dump_row.get('Parent Incident')).strip():
                    new_record['State'] = f"{dump_row['State']}/duplicate"

                df_master_updated = pd.concat([df_master_updated, pd.DataFrame([new_record])], ignore_index=True)

    # Save and send output
    output = BytesIO()
    df_master_updated.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="updated_master.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)
