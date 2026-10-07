"""
SQL queries against the live `checkpoint_live` MySQL database.

These replace the original checkwise_flat_recal S3/parquet data source.
The original hand-off queries this project inherited (see git-less history)
were written in Redshift/Postgres dialect (DATEADD, DATE_PART('dow', ...),
INTERVAL '2 days') against tables that don't exist in this MySQL database
(`checkpoint_live_ec_case_checks`, `..._for_networkdays`) - confirmed by
testing directly against the DB. They also depended on a 3rd holiday type
that doesn't exist in the live `ec_master_holidays` table (only types 1
and 2 - Sunday+festivals and Saturday, respectively).

QUERY_BASE_DATA below fetches one row per case-check with all raw and
decoded columns needed by the checkwise_flat_recal report logic. All of
the actual TAT/ageing/due-date arithmetic (which requires chaining several
computed columns - not supported by MySQL within a single SELECT list) is
done in pandas in tat_logic.py instead, applied twice (once per
`received_date`, once per the reopen/insufficiency-aware `max_date`) to
produce the Flat and Recal datasets respectively.
"""

QUERY_BASE_DATA = """
SELECT
    ecc.case_id AS case_id,
    ecc.case_check_id AS case_check_id,
    ec.client_external_id AS client_external_id,
    emc.company_name AS company_name,
    CONCAT(COALESCE(eud1.user_first_name, ''), ' ', COALESCE(eud1.user_last_name, '')) AS cat,
    CONCAT(COALESCE(eud2.user_first_name, ''), ' ', COALESCE(eud2.user_last_name, '')) AS cat_tl,
    CONCAT(COALESCE(eud3.user_first_name, ''), ' ', COALESCE(eud3.user_last_name, '')) AS account_manager,
    ecp.process_name AS process_name,
    ecm.case_ars_no AS case_ars_no,
    ecm.received_date AS received_date,
    CASE ecm.case_status
        WHEN 1 THEN 'New (Incomplete)'
        WHEN 2 THEN 'On Hold'
        WHEN 3 THEN 'Insufficient'
        WHEN 4 THEN 'Work in Progress'
        WHEN 5 THEN 'Pending for report'
        WHEN 6 THEN 'Closed by Client'
        WHEN 7 THEN 'Completed'
        WHEN 8 THEN 'Closed by Authbridge'
        WHEN 9 THEN 'Closed-Case Insufficient'
        WHEN 10 THEN 'HighlighterCase'
        WHEN 11 THEN 'SignOff Pending'
        WHEN 12 THEN 'Pending For Duplicity'
        WHEN 13 THEN 'Duplicity'
        WHEN 14 THEN 'Scrap'
        WHEN 15 THEN 'Excel Pending For Duplicity'
        WHEN 16 THEN 'Escalation Raised'
        WHEN 17 THEN 'Escalation Received'
        WHEN 18 THEN 'Excel Duplicity'
    END AS case_status,
    CASE
        WHEN ecc.insuff_fulfill_date IS NOT NULL
             AND (ecc.check_closure_date IS NULL OR ecc.insuff_fulfill_date <= ecc.check_closure_date)
        THEN ecc.insuff_fulfill_date
        ELSE NULL
    END AS insuff_fulfill_date,
    CASE ecc.check_status
        WHEN 0 THEN 'Documentation Pending'
        WHEN 1 THEN 'New/UnAssigned'
        WHEN 2 THEN 'On Hold'
        WHEN 3 THEN 'Insufficient'
        WHEN 4 THEN 'Work in Progress'
        WHEN 5 THEN 'Awaiting Response'
        WHEN 6 THEN 'Escalated'
        WHEN 7 THEN 'In Research'
        WHEN 8 THEN 'Completed'
        WHEN 9 THEN 'Disabled'
        WHEN 10 THEN 'case closed by client'
        WHEN 11 THEN 'Closed with Insufficiency'
        WHEN 12 THEN 'Closed-case Insufficient'
        WHEN 13 THEN 'Contractually on Hold'
    END AS check_status,
    ecc.check_severity AS check_severity,
    ecc.check_closure_date AS check_closure_date,
    COALESCE(ectc.tat_catg, ecp.tat_days_type) AS check_tat_category,
    COALESCE(ectc1.actual_tat, ecp.tat) AS check_tat,
    CASE
        WHEN ecc.reopen_date IS NOT NULL
             AND (ecc.check_closure_date IS NULL OR ecc.reopen_date <= ecc.check_closure_date)
        THEN ecc.reopen_date
        ELSE NULL
    END AS reopen_date_with_condition,
    CASE
        WHEN ecc.go_ahead_date IS NOT NULL
             AND (ecc.check_closure_date IS NULL OR ecc.go_ahead_date <= ecc.check_closure_date)
        THEN ecc.go_ahead_date
        ELSE NULL
    END AS goahead_date_with_condition,
    ec1.check_name AS `unique check name`,
    ec1.check_ops_name AS check_ops_name
FROM ec_case_checks ecc
LEFT JOIN ec_case_master ecm ON ecc.case_id = ecm.case_id
LEFT JOIN ec_master_company emc ON ecm.client_id = emc.company_id
LEFT JOIN ec_case_tat_config ectc ON (ecm.client_id = ectc.client_id AND ecm.process_id = ectc.process_id)
LEFT JOIN ec_check_tat_config ectc1 ON (ectc.id = ectc1.case_tat_id AND ecc.check_id = ectc1.check_id)
LEFT JOIN ec_client_process ecp ON ecm.process_id = ecp.process_id
INNER JOIN ec_process_checks epc ON (ecm.process_id = epc.process_id AND epc.check_id = ecc.check_id)
LEFT JOIN ec_client ec ON ecm.client_id = ec.client_id
LEFT JOIN ec_checks ec1 ON ecc.check_id = ec1.check_id
LEFT JOIN ec_user_details eud1 ON ec.cat_id = eud1.user_id
LEFT JOIN ec_user_details eud2 ON ec.cat_tl = eud2.user_id
LEFT JOIN ec_user_details eud3 ON ec.cat_account_manager = eud3.user_id
WHERE ecc.check_status <> 9
  AND ecm.case_status <> 8
  AND ec1.check_name NOT LIKE '%%Site Visit%%'
  AND ec.client_external_id IN %(client_external_ids)s
  AND ecm.received_date >= %(date_from)s
  AND ecm.received_date <= %(date_to)s
"""

QUERY_HOLIDAYS = """
SELECT HOLIDAY_DATE AS holiday_date, HOLIDAY_TYPE AS holiday_type
FROM ec_master_holidays
"""

QUERY_AVAILABLE_CLIENTS = """
SELECT DISTINCT ec.client_external_id AS client_external_id, emc.company_name AS company_name
FROM ec_case_master ecm
JOIN ec_client ec ON ecm.client_id = ec.client_id
JOIN ec_master_company emc ON ecm.client_id = emc.company_id
WHERE ec.client_external_id IS NOT NULL
ORDER BY emc.company_name
"""

QUERY_CLIENT_DATE_RANGE = """
SELECT MIN(ecm.received_date) AS min_date, MAX(ecm.received_date) AS max_date
FROM ec_case_master ecm
JOIN ec_client ec ON ecm.client_id = ec.client_id
WHERE ec.client_external_id = %(client_external_id)s
"""
