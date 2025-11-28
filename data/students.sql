=
DROP TABLE IF EXISTS uni_students;

-- Build unified, cleaned, normalized table
CREATE TABLE uni_students AS
SELECT
    TRIM(`StudentID`)                        AS Student_ID,
    TRIM(`FirstName`)                        AS First_Name,
    TRIM(`LastName`)                         AS Last_Name,
    TRIM(`Class`)                            AS Class,
    TRIM(`Term`)                             AS Term,
    TRIM(`Math`)                             AS Math,
    TRIM(`English`)                          AS English,
    TRIM(`Science`)                          AS Science,
    TRIM(`History`)                          AS History,
    NULLIF(TRIM(`Attendance (%)`), '')       AS Attendance,      -- empty → NULL
    TRIM(`ProjectScore`)                     AS Score,

    -- Passed: Y/N → 1/0
    CASE 
        WHEN UPPER(TRIM(`Passed (Y/N)`)) = 'Y' THEN 1
        WHEN UPPER(TRIM(`Passed (Y/N)`)) = 'N' THEN 0
        ELSE NULL
    END                                       AS Passed,

    -- Incoming_Student: TRUE/FALSE → 1/0
    CASE 
        WHEN UPPER(TRIM(`IncomeStudent`)) = 'TRUE' THEN 1
        WHEN UPPER(TRIM(`IncomeStudent`)) = 'FALSE' THEN 0
        ELSE NULL
    END                                       AS Incoming_Student,

    TRIM(`Cohort`)                            AS Cohort
FROM (
    SELECT * FROM Data
    UNION ALL
    SELECT * FROM Finance
    UNION ALL
    SELECT * FROM BM
) AS all_students
WHERE NULLIF(TRIM(`Attendance (%)`), '') IS NOT NULL;

-- Remove rows with ANY NULLs by creating a cleaned replacement
CREATE TABLE uni_students_clean AS
SELECT *
FROM uni_students
WHERE
    Student_ID IS NOT NULL
    AND First_Name IS NOT NULL
    AND Last_Name IS NOT NULL
    AND Class IS NOT NULL
    AND Term IS NOT NULL
    AND Math IS NOT NULL
    AND English IS NOT NULL
    AND Science IS NOT NULL
    AND History IS NOT NULL
    AND Attendance IS NOT NULL
    AND Score IS NOT NULL
    AND Passed IS NOT NULL
    AND Income_Student IS NOT NULL
    AND Cohort IS NOT NULL;

-- Replace the original with the cleaned table (cross-version safe syntax)
DROP TABLE uni_students;
RENAME TABLE uni_students_clean TO uni_students;

-- Index for performance
CREATE INDEX idx_uni_students_student_id ON uni_students (Student_ID);

-- Verify
SELECT * FROM uni_students;
