SELECT
    ca.cust_account_id,
    ca.account_number,
    p.party_name,
    cp.contact_point_type,
    cp.email_address,
    cp.phone_number,
    cp.status,
    cp.primary_flag
FROM 
    AR.HZ_CUSTOMER_PROFILES prof
JOIN 
    AR.HZ_CUST_ACCOUNTS ca ON prof.cust_account_id = ca.cust_account_id
JOIN 
    AR.HZ_PARTIES p ON ca.party_id = p.party_id
JOIN 
    AR.HZ_CONTACT_POINTS cp ON cp.owner_table_id = p.party_id
WHERE 
    prof.collector_id = 256022 -- Needs to be collector id found in AR_COLLECTORS
    AND cp.contact_point_type IN ('EMAIL', 'PHONE')
    AND cp.status = 'A'
    AND p.party_name IS NOT NULL
