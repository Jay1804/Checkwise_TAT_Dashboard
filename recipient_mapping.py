"""
Parses the user-uploaded client -> recipient mapping file.

Expected columns (case-insensitive, order doesn't matter):
    client_external_id, company_name, To_address, CC_address

Each row maps one client_external_id to the To/CC addresses its report
should be emailed to. Multiple addresses in To_address/CC_address may be
separated by comma or semicolon. Clients not present in the uploaded file
are simply not emailed - the caller decides how to surface that (see
report_generator.generate_reports_for_clients), no address is ever guessed.
"""
import pandas as pd

REQUIRED_COLUMNS = ['client_external_id', 'company_name', 'to_address', 'cc_address']


class MappingError(RuntimeError):
    """Raised when the uploaded mapping file is missing or malformed."""


def _split_addresses(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    return [addr.strip() for addr in str(value).replace(';', ',').split(',') if addr.strip()]


def parse_mapping_file(uploaded_file):
    """Returns {client_external_id: {'company_name': str, 'to': [...], 'cc': [...]}}."""
    name = uploaded_file.name.lower()
    if name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    elif name.endswith('.xlsx') or name.endswith('.xls'):
        df = pd.read_excel(uploaded_file)
    else:
        raise MappingError('Mapping file must be a .csv or .xlsx file.')

    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise MappingError(
            f"Mapping file is missing required column(s): {', '.join(missing)}. "
            f"Expected columns: client_external_id, company_name, To_address, CC_address."
        )

    df['client_external_id'] = pd.to_numeric(df['client_external_id'], errors='coerce')
    bad_rows = df['client_external_id'].isna().sum()
    if bad_rows:
        raise MappingError(
            f"Mapping file has {bad_rows} row(s) with a blank or non-numeric client_external_id."
        )
    df['client_external_id'] = df['client_external_id'].astype('int64')

    mapping = {}
    skipped_no_to = []
    for _, row in df.iterrows():
        to_addrs = _split_addresses(row['to_address'])
        client_id = int(row['client_external_id'])
        if not to_addrs:
            skipped_no_to.append(client_id)
            continue
        mapping[client_id] = {
            'company_name': row.get('company_name'),
            'to': to_addrs,
            'cc': _split_addresses(row.get('cc_address')),
        }

    return mapping, skipped_no_to
