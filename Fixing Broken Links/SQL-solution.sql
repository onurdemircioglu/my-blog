-- Oracle SQL
SELECT MY_LINK
        ,'servername' || SUBSTR(MY_LINK,INSTR(MY_LINK,'\',1), LENGTH(MY_LINK)-INSTR(MY_LINK,'\',1)+1) AS NEW_LINK
FROM
    (
    SELECT 'K:\MainFolder\SubFolder1\SubFolder2' AS MY_LINK FROM DUAL
    )
;
