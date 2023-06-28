-- Settings update
ALTER SESSION SET NLS_TERRITORY = 'TURKEY';
ALTER SESSION set NLS_DATE_FORMAT = 'DD.MM.YYYY';
ALTER SESSION set NLS_TIME_FORMAT = 'HH24:MI:SSXFF';


CREATE TABLE CALENDAR (
  CALENDAR_DATE DATE PRIMARY KEY,
  CALENDAR_DAY NUMBER(1,0),
  CALENDAR_MONTH NUMBER(2,0),
  CALENDAR_YEAR NUMBER(4,0),
  DAY_NAME VARCHAR2(10),
  MONTH_NAME VARCHAR2(10),
  QUARTER NUMBER(1,0),
  IS_WORKDAY NUMBER(1,0)
);

DECLARE
  START_DATE DATE := TO_DATE('01-JAN-2021', 'DD-MON-YYYY');
  END_DATE DATE := TO_DATE('31-DEC-2022', 'DD-MON-YYYY');
BEGIN
  FOR i IN 0..(END_DATE-START_DATE)
  LOOP
    INSERT INTO CALENDAR (CALENDAR_DATE, CALENDAR_DAY, CALENDAR_MONTH, CALENDAR_YEAR, DAY_NAME, MONTH_NAME, QUARTER, IS_WORKDAY)
    VALUES (
      START_DATE + i,
      TO_NUMBER(TO_CHAR(START_DATE + i, 'D')),
      TO_NUMBER(TO_CHAR(START_DATE + i, 'MM')),
      TO_NUMBER(TO_CHAR(START_DATE + i, 'YYYY')),
      TO_CHAR(START_DATE + i, 'DAY'),
      TO_CHAR(START_DATE + i, 'MONTH'),
      CEIL(TO_NUMBER(TO_CHAR(START_DATE + i, 'MM')) / 3),
      CASE WHEN TO_CHAR(START_DATE + i, 'D') NOT IN (6, 7) THEN 1 ELSE 0 END
    );
  END LOOP;
END;
/


SELECT *
FROM
    (
    SELECT A.*
        	,ROW_NUMBER() OVER (PARTITION BY CALENDAR_YEAR, CALENDAR_MONTH ORDER BY CALENDAR_DATE DESC) AS MY_RANKING
    FROM CALENDAR A
    WHERE 1=1
    	AND IS_WORKDAY = 1
    ) 
WHERE 1=1
    	AND MY_RANKING = 5
ORDER BY 1
;