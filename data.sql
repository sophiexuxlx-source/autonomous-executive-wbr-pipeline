SELECT 
    category,
    DATE(time) AS event_date,
    COUNT(*) AS event_count
FROM 
    events
GROUP BY 
    DATE(time), category
ORDER BY
    event_date, category;