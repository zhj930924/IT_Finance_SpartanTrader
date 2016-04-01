/* 1. Who are our customers in Austin, TX? Produce a list of our
customers� first names, last names and customer IDs, sorted by last name.*/

SELECT f_name, l_name, c_id
FROM CUSTOMER
WHERE city = 'Austin' AND state = 'TX'
ORDER BY l_name

/* 2. What rates did we offer on our loans? Produce a list of the loan IDs
and their rates for the loans under $50,000,000 */

SELECT l_id, rate
FROM LOAN
WHERE principal < 50000000

/* 3. What is the first and last name of the of loan officer whose last name begins by �Mon�?
*/

SELECT f_name, l_name
FROM LOAN_OFFICER
WHERE l_name LIKE 'Mon%'

/* 4. How many loans over $100,000,000 do we have? Produce a count, not a list.
*/

/* Simple count by rows */
SELECT COUNT(*)
FROM LOAN
WHERE principal > 100000000

/* Same result count by l_id => Primary Key => Unique */
SELECT COUNT(l_id)
FROM LOAN
WHERE principal > 100000000


/* 5. What is the average rate for the loans higher than $100,000,000?
*/

SELECT AVG(rate)
FROM LOAN
WHERE principal > 100000000

/* 6. What is our total exposure (total principal) per level of interest rate
(hint: see handout on grouping query results) */

SELECT rate, SUM(principal)
FROM LOAN
GROUP BY rate

/* 7. Produce a list containing all customer last names and first names, their loan officers last name,
and the office phone for that loan officer. Sort by Customer last name.
Hint: some customers have multiple loan officers - list them all.
Some customers have multiple loans with the same officer - list him/her only once per customer.*/

SELECT DISTINCT CUSTOMER.l_name, CUSTOMER.f_name, LOAN_OFFICER.l_name, LOAN_OFFICER.phone
FROM CUSTOMER, LOAN, LOAN_OFFICER, CUSTOMER_IN_LOAN
WHERE CUSTOMER.c_id = CUSTOMER_IN_LOAN.c_id
AND LOAN.l_id = CUSTOMER_IN_LOAN.l_id
AND LOAN_OFFICER.lo_id = LOAN.lo_id
ORDER BY CUSTOMER.l_name

/* 8. Who is managing the most money? Produce a list of loan officers� last names
and the total money they manage. Hint: sort it by the total they manage. */

/* Most Money */
SELECT TOP 1 LOAN_OFFICER.l_name, SUM(principal)
FROM LOAN JOIN LOAN_OFFICER
	ON (LOAN.lo_id = LOAN_OFFICER.lo_id)
GROUP BY l_name
ORDER BY SUM(principal) DESC

/* List from most money to least */
SELECT LOAN_OFFICER.l_name, SUM(principal)
FROM LOAN JOIN LOAN_OFFICER
	ON (LOAN.lo_id = LOAN_OFFICER.lo_id)
GROUP BY l_name
ORDER BY SUM(principal) DESC
