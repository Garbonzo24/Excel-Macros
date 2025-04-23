SELECT
    ca.cust_account_id,
    ca.account_number,
    customer.party_name AS customer_name,
    contact.party_name AS contact_name,
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
    AR.HZ_PARTIES customer ON ca.party_id = customer.party_id
JOIN 
    AR.HZ_RELATIONSHIPS rel
        ON rel.object_id = customer.party_id OR rel.subject_id = customer.party_id
JOIN 
    AR.HZ_CONTACT_POINTS cp 
        ON cp.owner_table_id = rel.party_id
LEFT JOIN 
    AR.HZ_PARTIES contact ON rel.subject_id = contact.party_id OR rel.object_id = contact.party_id
WHERE 
    prof.collector_id = 256022 -- Needs to be collector ID found in oracle
    AND cp.contact_point_type IN ('EMAIL', 'PHONE')
    AND cp.status = 'A'
