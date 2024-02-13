Simple program written in Python to search tables and columns of the PostgresSQL database

Usage: Enter search term and click search button
Example: Search for all references of Personid:
- Enter "personid" and click search
- All instances of the searched term will be highlighted in the results

Wildcards are not supported at this time
- using "p" as a search term would currently return every table and column that contains "p", no matter where is exists in the name
